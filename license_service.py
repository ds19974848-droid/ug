"""授权码生成与验证，兼容外置授权管理工具。"""

import base64
import hashlib
import json
import os
import socket
import subprocess
from datetime import datetime, timedelta
from math import ceil

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

from .config import config
from .db import get_setting, get_session, set_settings, write_audit
from .models import LicenseGrant


LICENSE_PUBLIC_KEY = "lY3z6QcW/S3Pcuu1GjecHw5NcS8Gnn7LgArzMdDiAbA="
LICENSE_PREFIX = "ICLA2"
MACHINE_CODE_LENGTH = 16
ACTIVE_LICENSE_SETTING = "active_license"


def _b64decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def get_machine_code() -> str:
    """返回与外置工具一致的机器码：硬盘序列号 + 计算机名。"""
    try:
        out = subprocess.check_output(
            "wmic volume where driveletter='C:' get serialnumber",
            shell=True,
            stderr=subprocess.DEVNULL,
        )
        disk_serial = "".join(
            out.decode("utf-8", "ignore").strip().splitlines()[1:]
        ).strip()
    except Exception:
        disk_serial = "UNKNOWN"

    try:
        computer_name = os.environ.get("COMPUTERNAME", "UNKNOWN")
    except Exception:
        computer_name = "UNKNOWN"
    raw = f"{disk_serial}-{computer_name}"
    return hashlib.md5(raw.encode()).hexdigest().upper()[:MACHINE_CODE_LENGTH]


def normalize_machine_code(value: str) -> str:
    return "".join(ch for ch in (value or "").strip().upper() if ch.isalnum())


def _expiry_datetime(expire: str) -> datetime | None:
    expire = (expire or "").strip()
    if not expire:
        return None
    return datetime.strptime(expire, "%Y%m%d")


def verify_license_code(license_key: str) -> dict:
    """Verify a signed license with the public key bundled in the client."""
    try:
        parts = (license_key or "").strip().split(".")
        if len(parts) != 3 or parts[0] != LICENSE_PREFIX:
            raise ValueError("unsupported license format")
        payload_bytes = _b64decode(parts[1])
        Ed25519PublicKey.from_public_bytes(_b64decode(LICENSE_PUBLIC_KEY)).verify(
            _b64decode(parts[2]), payload_bytes
        )
        payload = json.loads(payload_bytes.decode("utf-8"))
        machine_code = payload.get("mc", "")
        expire = payload.get("exp", "")

        days_left = -1
        if expire:
            expire_datetime = datetime.strptime(expire, "%Y%m%d") + timedelta(days=1)
            now = datetime.now()
            if now >= expire_datetime:
                return {
                    "valid": False,
                    "message": f"激活码已过期（过期日期: {expire_datetime.strftime('%Y-%m-%d')}）",
                    "machine_code": machine_code,
                    "expire_date": expire,
                    "days_left": -1,
                }
            days_left = max(0, ceil((expire_datetime - now).total_seconds() / 86400))

        message = "激活成功！"
        if days_left < 0:
            message += " 永久使用"
        else:
            message += f" 剩余 {days_left} 天"
        message += f"\n机器码: {machine_code}"
        return {
            "valid": True,
            "message": message,
            "machine_code": machine_code,
            "expire_date": expire,
            "days_left": days_left,
        }
    except (TypeError, ValueError, KeyError, IndexError, json.JSONDecodeError, InvalidSignature):
        return {
            "valid": False,
            "message": "激活码格式错误",
            "machine_code": "",
            "expire_date": "",
            "days_left": -1,
        }


def plan_to_days(plan: str) -> int:
    return {
        "1天": 1,
        "3天": 3,
        "7天": 7,
        "一个月": 30,
        "永久": 0,
    }.get(plan, 0)


def _serialize_license(license_key: str, result: dict, recipient: str = "") -> dict:
    return {
        "license_key": license_key,
        "machine_code": result.get("machine_code", ""),
        "expire_date": result.get("expire_date", ""),
        "days_left": result.get("days_left", -1),
        "recipient": recipient,
        "activated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def activate_license(license_key: str, recipient: str = "") -> dict:
    """验证并保存本机授权。"""
    license_key = (license_key or "").strip()
    result = verify_license_code(license_key)
    if not result["valid"]:
        return {"ok": False, "message": result["message"], "result": result}
    local_machine = get_machine_code()
    if normalize_machine_code(result.get("machine_code", "")) != normalize_machine_code(local_machine):
        return {
            "ok": False,
            "message": "授权码对应机器码与本机不一致",
            "result": result,
        }
    session = get_session()
    try:
        data = _serialize_license(license_key, result, recipient)
        set_settings({ACTIVE_LICENSE_SETTING: json.dumps(data, ensure_ascii=False)})
        grant = session.query(LicenseGrant).filter(
            LicenseGrant.license_key == license_key
        ).first()
        if grant is None:
            grant = LicenseGrant(
                license_key=license_key,
                machine_code=result.get("machine_code", ""),
                plan_label="已激活",
                days=0 if not result.get("expire_date") else -1,
                expires_at=_expiry_datetime(result.get("expire_date", "")),
                status="activated",
                recipient=recipient,
                activated_at=datetime.now(),
                machine_name=socket.gethostname(),
            )
            session.add(grant)
        else:
            grant.status = "activated"
            grant.activated_at = datetime.now()
            grant.machine_code = result.get("machine_code", "")
            grant.expires_at = _expiry_datetime(result.get("expire_date", ""))
            grant.recipient = recipient or grant.recipient
            grant.machine_name = socket.gethostname()
        write_audit(session, "activate", "license", grant.id, f"激活授权: {result.get('machine_code')}")
        session.commit()
        return {"ok": True, "message": result["message"], "result": result}
    finally:
        session.close()


def get_active_license_state() -> dict:
    """返回当前本机授权状态，过期或签名错误时也给出结果。"""
    raw = get_setting(ACTIVE_LICENSE_SETTING, "")
    if not raw:
        return {
            "valid": False,
            "message": "未激活授权",
            "machine_code": get_machine_code(),
            "expire_date": "",
            "days_left": -1,
            "recipient": "",
        }
    try:
        data = json.loads(raw)
    except (TypeError, json.JSONDecodeError):
        data = {}
    license_key = data.get("license_key", "")
    result = verify_license_code(license_key)
    result.setdefault("recipient", data.get("recipient", ""))
    if not result["valid"]:
        return result
    if normalize_machine_code(result.get("machine_code", "")) != normalize_machine_code(get_machine_code()):
        result.update({"valid": False, "message": "授权码对应机器码与本机不一致"})
    return result


def revoke_local_license() -> bool:
    session = get_session()
    try:
        set_settings({ACTIVE_LICENSE_SETTING: ""})
        raw = get_setting(ACTIVE_LICENSE_SETTING, "")
        data = json.loads(raw) if raw else {}
        license_key = data.get("license_key", "")
        if license_key:
            grant = session.query(LicenseGrant).filter(
                LicenseGrant.license_key == license_key
            ).first()
            if grant:
                grant.status = "revoked"
                session.commit()
        return True
    finally:
        session.close()


def save_issued_license(
    license_key: str,
    machine_code: str,
    plan_label: str,
    days: int,
    recipient: str = "",
) -> LicenseGrant:
    session = get_session()
    try:
        result = verify_license_code(license_key)
        expires_at = _expiry_datetime(result.get("expire_date", ""))
        grant = session.query(LicenseGrant).filter(
            LicenseGrant.license_key == license_key
        ).first()
        if grant is None:
            grant = LicenseGrant(
                license_key=license_key,
                machine_code=normalize_machine_code(machine_code),
                plan_label=plan_label,
                days=days,
                expires_at=expires_at,
                status="issued",
                recipient=recipient,
            )
            session.add(grant)
        else:
            grant.machine_code = normalize_machine_code(machine_code)
            grant.plan_label = plan_label
            grant.days = days
            grant.expires_at = expires_at
            grant.status = "issued"
            grant.recipient = recipient
        write_audit(session, "create", "license", grant.id, f"生成授权: {plan_label} {recipient}")
        session.commit()
        return grant
    finally:
        session.close()


def get_license_grants() -> list[dict]:
    session = get_session()
    try:
        grants = session.query(LicenseGrant).order_by(LicenseGrant.created_at.desc()).all()
        return [
            {
                "id": grant.id,
                "license_key": grant.license_key,
                "machine_code": grant.machine_code,
                "plan_label": grant.plan_label,
                "days": grant.days,
                "issued_at": grant.issued_at,
                "expires_at": grant.expires_at,
                "status": grant.status,
                "recipient": grant.recipient,
                "activated_at": grant.activated_at,
                "machine_name": grant.machine_name,
            }
            for grant in grants
        ]
    finally:
        session.close()


def delete_license_grant(grant_id: int) -> bool:
    session = get_session()
    try:
        grant = session.query(LicenseGrant).filter(LicenseGrant.id == grant_id).first()
        if grant is None:
            return False
        session.delete(grant)
        session.commit()
        return True
    finally:
        session.close()


def set_license_enforcement(enabled: bool) -> None:
    if enabled:
        config.LICENSE_REQUIRED_FILE.parent.mkdir(parents=True, exist_ok=True)
        config.LICENSE_REQUIRED_FILE.write_text("enabled", encoding="utf-8")
    else:
        try:
            config.LICENSE_REQUIRED_FILE.unlink(missing_ok=True)
        except OSError:
            pass


def license_enforcement_enabled() -> bool:
    return config.LICENSE_REQUIRED_FILE.exists()
