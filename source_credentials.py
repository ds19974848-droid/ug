"""使用 Windows DPAPI 保存来源授权信息。"""
from __future__ import annotations

import base64
import ctypes
import json
from ctypes import wintypes

from .db import get_session
from .models import AppSettings


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(data: bytes):
    buffer = ctypes.create_string_buffer(data)
    return _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char))), buffer


def protect_text(value: str) -> str:
    if not value:
        return ""
    input_blob, input_buffer = _blob(value.encode("utf-8"))
    output_blob = _DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(input_blob), "DashuoCostCloud", None, None, None, 0,
        ctypes.byref(output_blob),
    ):
        raise ctypes.WinError()
    try:
        encrypted = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        return base64.b64encode(encrypted).decode("ascii")
    finally:
        ctypes.windll.kernel32.LocalFree(output_blob.pbData)


def unprotect_text(value: str) -> str:
    if not value:
        return ""
    input_blob, input_buffer = _blob(base64.b64decode(value))
    output_blob = _DataBlob()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(input_blob), None, None, None, None, 0,
        ctypes.byref(output_blob),
    ):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData).decode("utf-8")
    finally:
        ctypes.windll.kernel32.LocalFree(output_blob.pbData)


def save_source_credentials(source_id: int, auth_type: str, username: str, secret: str):
    session = get_session()
    key = f"source_auth_{source_id}"
    try:
        setting = session.query(AppSettings).filter(AppSettings.key == key).first()
        payload = json.dumps({
            "auth_type": auth_type,
            "username": username,
            "secret": protect_text(secret),
        }, ensure_ascii=False)
        if setting is None:
            setting = AppSettings(key=key, value=payload)
            session.add(setting)
        else:
            setting.value = payload
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def load_source_credentials(source_id: int, *, include_secret: bool = True) -> dict:
    session = get_session()
    try:
        setting = session.query(AppSettings).filter(AppSettings.key == f"source_auth_{source_id}").first()
        if not setting:
            return {"auth_type": "none", "username": "", "secret": ""}
        payload = json.loads(setting.value or "{}")
        if include_secret and payload.get("secret"):
            payload["secret"] = unprotect_text(payload["secret"])
        elif not include_secret:
            payload["secret"] = ""
        return payload
    except Exception:
        return {"auth_type": "none", "username": "", "secret": ""}
    finally:
        session.close()


def source_auth_headers(source_id: int) -> dict[str, str]:
    credential = load_source_credentials(source_id)
    auth_type = credential.get("auth_type", "none")
    username = credential.get("username", "")
    secret = credential.get("secret", "")
    if auth_type == "basic" and (username or secret):
        token = base64.b64encode(f"{username}:{secret}".encode("utf-8")).decode("ascii")
        return {"Authorization": f"Basic {token}"}
    if auth_type == "bearer" and secret:
        return {"Authorization": f"Bearer {secret}"}
    if auth_type == "cookie" and secret:
        return {"Cookie": secret}
    return {}
