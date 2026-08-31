"""数据库服务层：初始化、Session管理、CRUD操作"""
import json
import sqlite3
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional
from sqlalchemy import create_engine, event, func, text
from sqlalchemy.orm import sessionmaker, Session
from .config import config
from .models import Base, now, Region, RegionStandard, FeeRule
from .models import Material, MaterialAlias, MaterialPrice, PriceHistory
from .models import OfficialSource, SubscriptionTask, SourceDocument, AiParseLog
from .models import Project, ProjectItem, ProjectPriceSnapshot, ProjectQuotaMatch
from .models import AuditLog, BackupRecord, AppSettings, CostItem, ProjectListData
from .models import LicenseGrant, QuotaItem, QuotaComposition

engine = create_engine(
    config.DATABASE_URL,
    connect_args={"check_same_thread": False, "timeout": 60} if "sqlite" in config.DATABASE_URL else {},
    echo=False,
    pool_pre_ping=True,
)


if "sqlite" in config.DATABASE_URL:
    @event.listens_for(engine, "connect")
    def _configure_sqlite_connection(dbapi_connection, _connection_record):
        cursor = dbapi_connection.cursor()
        try:
            cursor.execute("PRAGMA busy_timeout=60000")
            cursor.execute("PRAGMA journal_mode=WAL")
            cursor.execute("PRAGMA synchronous=NORMAL")
        finally:
            cursor.close()

SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)


def init_db():
    """创建所有表 + 种子数据"""
    config.ensure_dirs()
    Base.metadata.create_all(engine)
    _migrate_schema()
    _seed_data()
    _repair_quota_majors()
    _repair_invalid_subscription_data()
    _ensure_shanghai_api_source()
    _ensure_shenzhen_api_source()
    _ensure_chengdu_api_source()
    from .subscription import prune_retrieved_price_storage
    prune_retrieved_price_storage()
    _load_runtime_settings()


def get_session() -> Session:
    return SessionLocal()


def _migrate_schema():
    """执行可重复运行的小型 SQLite 迁移。"""
    if "sqlite" not in config.DATABASE_URL:
        return
    with engine.begin() as connection:
        columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(material_prices)"))
        }
        if "source_key" not in columns:
            connection.execute(text(
                "ALTER TABLE material_prices ADD COLUMN source_key VARCHAR(300) DEFAULT ''"
            ))
        for column, definition in {
            "source_type": "VARCHAR(50) DEFAULT 'official'",
            "price_basis": "VARCHAR(50) DEFAULT 'as_published'",
            "valid_from": "VARCHAR(20) DEFAULT ''",
            "valid_to": "VARCHAR(20) DEFAULT ''",
            "is_withdrawn": "BOOLEAN DEFAULT 0",
        }.items():
            if column not in columns:
                connection.execute(text(f"ALTER TABLE material_prices ADD COLUMN {column} {definition}"))
        history_columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(price_history)"))
        }
        for column, definition in {
            "source_type": "VARCHAR(50) DEFAULT 'official'",
            "price_basis": "VARCHAR(50) DEFAULT 'as_published'",
        }.items():
            if column not in history_columns:
                connection.execute(text(f"ALTER TABLE price_history ADD COLUMN {column} {definition}"))
        project_columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(projects)"))
        }
        for column, definition in {
            "client": "VARCHAR(300) DEFAULT ''",
            "design_unit": "VARCHAR(300) DEFAULT ''",
            "contractor": "VARCHAR(300) DEFAULT ''",
            "supervision_unit": "VARCHAR(300) DEFAULT ''",
            "tender_agent": "VARCHAR(300) DEFAULT ''",
            "project_type": "VARCHAR(100) DEFAULT ''",
            "funding_source": "VARCHAR(100) DEFAULT ''",
            "contract_type": "VARCHAR(100) DEFAULT ''",
            "scale": "VARCHAR(50) DEFAULT ''",
            "daily_capacity": "FLOAT DEFAULT 0",
            "area": "FLOAT DEFAULT 0",
            "structure_type": "VARCHAR(100) DEFAULT ''",
            "process_type": "VARCHAR(100) DEFAULT ''",
            "project_location": "VARCHAR(300) DEFAULT ''",
            "project_address": "VARCHAR(500) DEFAULT ''",
            "pricing_province": "VARCHAR(100) DEFAULT ''",
            "pricing_city": "VARCHAR(100) DEFAULT ''",
            "pricing_district": "VARCHAR(100) DEFAULT ''",
            "pricing_date": "VARCHAR(20) DEFAULT ''",
            "price_year": "VARCHAR(10) DEFAULT ''",
            "stage": "VARCHAR(50) DEFAULT '投标报价'",
            "pricing_basis": "VARCHAR(300) DEFAULT ''",
            "boq_basis": "VARCHAR(300) DEFAULT ''",
            "tax_method": "VARCHAR(100) DEFAULT ''",
            "planned_start_date": "VARCHAR(20) DEFAULT ''",
            "planned_end_date": "VARCHAR(20) DEFAULT ''",
            "specialty": "VARCHAR(100) DEFAULT ''",
            "currency": "VARCHAR(10) DEFAULT 'CNY'",
            "management_rate": "FLOAT DEFAULT 0.05",
            "management_base": "VARCHAR(50) DEFAULT 'direct'",
            "profit_rate": "FLOAT DEFAULT 0.07",
            "profit_base": "VARCHAR(50) DEFAULT 'direct_management'",
            "tax_rate": "FLOAT DEFAULT 0.09",
        }.items():
            if column not in project_columns:
                connection.execute(text(f"ALTER TABLE projects ADD COLUMN {column} {definition}"))
        project_item_columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(project_items)"))
        }
        for column, definition in {
            "labor_unit_price": "FLOAT DEFAULT 0",
            "material_unit_price": "FLOAT DEFAULT 0",
            "machinery_unit_price": "FLOAT DEFAULT 0",
            "unallocated_unit_price": "FLOAT DEFAULT 0",
            "management_unit_price": "FLOAT DEFAULT 0",
            "profit_unit_price": "FLOAT DEFAULT 0",
            "custom_data": "TEXT DEFAULT ''",
        }.items():
            if column not in project_item_columns:
                connection.execute(text(f"ALTER TABLE project_items ADD COLUMN {column} {definition}"))
        subscription_task_columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(subscription_tasks)"))
        }
        for column, definition in {
            "parsed_count": "INTEGER DEFAULT 0",
            "stored_count": "INTEGER DEFAULT 0",
            "rejected_count": "INTEGER DEFAULT 0",
            "phase": "VARCHAR(50) DEFAULT ''",
            "failure_stage": "VARCHAR(50) DEFAULT ''",
            "failure_reason": "TEXT DEFAULT ''",
            "next_action": "TEXT DEFAULT ''",
        }.items():
            if column not in subscription_task_columns:
                connection.execute(text(f"ALTER TABLE subscription_tasks ADD COLUMN {column} {definition}"))
        connection.execute(text(
            "CREATE TABLE IF NOT EXISTS project_quota_matches ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, "
            "project_id INTEGER NOT NULL, source_file VARCHAR(500) DEFAULT '', "
            "data_json TEXT DEFAULT '', total_rows INTEGER DEFAULT 0, "
            "matched_rows INTEGER DEFAULT 0, quota_count INTEGER DEFAULT 0, "
            "created_at DATETIME, updated_at DATETIME, "
            "FOREIGN KEY(project_id) REFERENCES projects(id))"
        ))
        connection.execute(text(
            "CREATE TABLE IF NOT EXISTS project_list_data ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, source_file VARCHAR(500) DEFAULT '', "
            "source_sheet VARCHAR(200) DEFAULT '', source_project VARCHAR(300) DEFAULT '', "
            "region VARCHAR(100) DEFAULT '', period VARCHAR(20) DEFAULT '', seq_no VARCHAR(50) DEFAULT '', "
            "item_code VARCHAR(120) DEFAULT '', item_name VARCHAR(500) NOT NULL, feature TEXT DEFAULT '', "
            "unit VARCHAR(50) DEFAULT '', quantity FLOAT, analysis TEXT DEFAULT '', "
            "comprehensive_price FLOAT, total_price FLOAT, labor_cost FLOAT, material_cost FLOAT, machinery_cost FLOAT, "
            "management_cost FLOAT, profit FLOAT, note TEXT DEFAULT '', raw_data TEXT DEFAULT '', "
            "created_at DATETIME, updated_at DATETIME)"
        ))
        project_list_data_columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(project_list_data)"))
        }
        if "total_price" not in project_list_data_columns:
            connection.execute(text("ALTER TABLE project_list_data ADD COLUMN total_price FLOAT"))
        connection.execute(text(
            "CREATE INDEX IF NOT EXISTS idx_project_list_data_name ON project_list_data (item_name)"
        ))
        connection.execute(text(
            "CREATE INDEX IF NOT EXISTS idx_project_quota_match_project "
            "ON project_quota_matches (project_id)"
        ))
        connection.execute(text(
            "UPDATE material_prices SET source_type = 'market_reference' "
            "WHERE trust_level = 'market_reference' AND (source_type IS NULL OR source_type = 'official')"
        ))
        connection.execute(text(
            "UPDATE price_history SET source_type = 'market_reference' "
            "WHERE trust_level = 'market_reference' AND (source_type IS NULL OR source_type = 'official')"
        ))
        connection.execute(text("DROP INDEX IF EXISTS idx_price_material_region_period"))
        connection.execute(text("DROP INDEX IF EXISTS idx_price_material_region_period_spec"))
        connection.execute(text(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_price_material_region_period_spec "
            "ON material_prices (material_id, region_id, period, spec, source_key)"
        ))


def _seed_data():
    """初始化基础种子数据"""
    session = SessionLocal()
    try:
        if session.query(Region).count() == 0:
            _seed_regions(session)
        if session.query(Material).count() == 0:
            _seed_materials(session)
            session.flush()
        if session.query(FeeRule).count() == 0:
            _seed_fee_rules(session)
            session.flush()
        if session.query(MaterialPrice).count() == 0:
            _seed_material_prices(session)
        if session.query(Project).count() == 0:
            _seed_projects(session)
        if session.query(CostItem).count() == 0:
            _seed_cost_items(session)
        if session.query(QuotaItem).count() == 0:
            _seed_quota_items(session)
        _seed_enterprise_reference_data(session)
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def _seed_cost_items(session: Session):
    source_file = config.BUILTIN_COST_ITEMS_FILE
    if not source_file.exists():
        return
    from .cost_item_service import import_cost_items_from_excel
    result = import_cost_items_from_excel(source_file, session, audit=False)
    if not result.get("success"):
        raise RuntimeError(result.get("error", "内置清单导入失败"))


def _seed_quota_items(session: Session):
    source_file = config.BUILTIN_QUOTA_FILE
    if not source_file.exists():
        return
    from .quota_service import import_quota_excel

    result = import_quota_excel(source_file, session, major="装修", audit=False)
    if not result.get("success"):
        raise RuntimeError(result.get("error", "内置定额导入失败"))


def _seed_enterprise_reference_data(session: Session):
    """Load the bundled enterprise reference table once into the reference library."""
    source_file = config.BUILTIN_ENTERPRISE_REFERENCE_FILE
    marker = "软件内置企业参考定额表"
    if not source_file.exists():
        return
    if session.query(ProjectListData.id).filter(ProjectListData.source_project == marker).first():
        return
    from .project_list_data_service import import_project_list_data

    result = import_project_list_data(
        str(source_file),
        session,
        source_project=marker,
    )
    if not result.get("success"):
        raise RuntimeError(result.get("error", "内置企业参考定额表导入失败"))


def _repair_quota_majors():
    """Repair quota majors that were previously forced to a single value.

    The quota workbook's first column is a fee category.  The actual major is
    encoded by the source filename, for example ``公路成本定额.xlsx``.
    """
    marker_key = "repair_quota_majors_from_filename_v1"
    session = SessionLocal()
    try:
        if session.query(AppSettings).filter(AppSettings.key == marker_key).first():
            return
        from .quota_service import _source_key, infer_quota_major

        changes = []
        for item in session.query(QuotaItem).all():
            inferred = infer_quota_major(item.source, "")
            if not inferred or inferred == item.major:
                continue
            new_source_key = _source_key(
                "excel",
                inferred,
                item.code,
                item.name,
                item.feature,
                item.unit,
            )
            duplicate = session.query(QuotaItem).filter(
                QuotaItem.source_key == new_source_key,
                QuotaItem.id != item.id,
            ).first()
            if duplicate is not None:
                continue
            changes.append((item, inferred, new_source_key))

        if changes:
            backup_database("automatic", "修复定额库专业字段前自动备份")
            for item, inferred, new_source_key in changes:
                item.major = inferred
                item.source_key = new_source_key
            write_audit(
                session,
                "update",
                "quota_item",
                detail=f"按定额文件名修复专业: {len(changes)} 条",
            )
        session.add(AppSettings(key=marker_key, value=str(len(changes))))
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def _load_runtime_settings():
    session = SessionLocal()
    try:
        values = {setting.key: setting.value for setting in session.query(AppSettings).all()}
        if "price_change_threshold" in values:
            config.PRICE_CHANGE_THRESHOLD = float(values["price_change_threshold"])
        if "price_min_sane" in values:
            config.PRICE_MIN_SANE = float(values["price_min_sane"])
        if "price_max_sane" in values:
            config.PRICE_MAX_SANE = float(values["price_max_sane"])
    finally:
        session.close()


def _repair_invalid_subscription_data():
    """清理旧版本从上海住建委首页误识别出的非价格资料。"""
    marker_key = "repair_invalid_homepage_prices_v1"
    session = get_session()
    try:
        if session.query(AppSettings).filter(AppSettings.key == marker_key).first():
            return
        shanghai = session.query(Region).filter(Region.name == "上海").first()
        broad_source = None
        if shanghai:
            broad_source = session.query(OfficialSource).filter(
                OfficialSource.region_id == shanghai.id,
                OfficialSource.url.in_([
                    "http://zjw.sh.gov.cn", "http://zjw.sh.gov.cn/",
                    "https://zjw.sh.gov.cn", "https://zjw.sh.gov.cn/",
                ]),
            ).first()
        needs_cleanup = broad_source is not None
    finally:
        session.close()
    _finish_invalid_subscription_cleanup(needs_cleanup)


def _ensure_shanghai_api_source():
    marker_key = "ensure_shanghai_api_source_v1"
    official_url = "https://ciac.zjw.sh.gov.cn/JGBXMGCZJInterWeb/pc/#/HyxxHynr?bmCode=003002"
    session = get_session()
    try:
        if session.query(AppSettings).filter(AppSettings.key == marker_key).first():
            return
        region = session.query(Region).filter(Region.code == "310000").first()
        if region is None:
            region = Region(name="上海", province="上海", code="310000", is_active=True)
            session.add(region)
            session.flush()
        matching = session.query(OfficialSource).filter(
            OfficialSource.region_id == region.id,
            OfficialSource.url == official_url,
        ).order_by(OfficialSource.id).all()
        if matching:
            primary = matching[0]
        else:
            primary = OfficialSource(
                region_id=region.id,
                name="上海市人工、材料、机械信息价（官方）",
                url=official_url,
                source_type="api",
                is_official=True,
                is_active=True,
            )
            session.add(primary)
            session.flush()
        primary.name = "上海市人工、材料、机械信息价（官方）"
        primary.source_type = "api"
        primary.is_official = True
        primary.is_active = True
        primary.last_result = ""
        primary.notes = "上海住建委官方人工、材料、机械信息价；使用专用 API 下载并保留原始 XLS 文件"

        for duplicate in matching[1:]:
            duplicate.is_active = False
            duplicate.last_result = "disabled"
            duplicate.notes = f"与来源 {primary.id} 重复，已自动停用"
        for source in session.query(OfficialSource).filter(OfficialSource.region_id == region.id).all():
            if source.id == primary.id:
                continue
            if source.url.rstrip("/") in {
                "https://zjw.sh.gov.cn",
                "https://www.shggzy.com",
                "https://www.shjjw.gov.cn",
            }:
                source.is_active = False
                source.last_result = "disabled"
                source.notes = "不是上海材料信息价数据入口，已从自动订阅中停用"
        session.add(AppSettings(key=marker_key, value=str(primary.id)))
        write_audit(session, "update", "official_source", primary.id, "启用上海官方信息价专用 API 来源")
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def _ensure_shenzhen_api_source():
    marker_key = "ensure_shenzhen_api_source_v1"
    official_url = "https://zjj.sz.gov.cn/szzjxx/web/pc/index"
    session = get_session()
    try:
        if session.query(AppSettings).filter(AppSettings.key == marker_key).first():
            return
        region = session.query(Region).filter(Region.code == "440300").first()
        if region is None:
            region = Region(name="深圳", province="广东", code="440300", is_active=True)
            session.add(region)
            session.flush()
        matching = session.query(OfficialSource).filter(
            OfficialSource.region_id == region.id,
            OfficialSource.url == official_url,
        ).order_by(OfficialSource.id).all()
        if matching:
            primary = matching[0]
        else:
            primary = OfficialSource(
                region_id=region.id,
                name="深圳市建设工程造价信息查询系统（官方）",
                url=official_url,
                source_type="api",
                is_official=True,
                is_active=True,
            )
            session.add(primary)
            session.flush()
        primary.name = "深圳市建设工程造价信息查询系统（官方）"
        primary.source_type = "api"
        primary.is_official = True
        primary.is_active = True
        primary.last_result = ""
        primary.notes = "深圳市住房和建设局官方造价信息查询系统；专用 API 接口，含税价，保留官方接口快照"

        for duplicate in matching[1:]:
            duplicate.is_active = False
            duplicate.last_result = "disabled"
            duplicate.notes = f"与来源 {primary.id} 重复，已自动停用"
        for source in session.query(OfficialSource).filter(OfficialSource.region_id == region.id).all():
            if source.id == primary.id:
                continue
            if source.url.rstrip("/") in {
                "http://www.szcost.com", "http://www.szcost.com/",
                "https://www.szcost.com", "https://www.szcost.com/",
                "http://zjj.sz.gov.cn", "http://zjj.sz.gov.cn/",
                "https://zjj.sz.gov.cn", "https://zjj.sz.gov.cn/",
            }:
                source.is_active = False
                source.last_result = "disabled"
                source.notes = "不是深圳材料信息价数据入口，已从自动订阅中停用"
        session.add(AppSettings(key=marker_key, value=str(primary.id)))
        write_audit(session, "update", "official_source", primary.id, "启用深圳官方造价信息专用 API 来源")
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def _ensure_chengdu_api_source():
    marker_key = "ensure_chengdu_api_source_v1"
    official_url = "http://202.61.90.35:8037/jgxx.htm?code=5101"
    session = get_session()
    try:
        if session.query(AppSettings).filter(AppSettings.key == marker_key).first():
            return
        region = session.query(Region).filter(Region.code == "510100").first()
        if region is None:
            region = Region(name="成都", province="四川", code="510100", is_active=True)
            session.add(region)
            session.flush()
        matching = session.query(OfficialSource).filter(
            OfficialSource.region_id == region.id,
            OfficialSource.url == official_url,
        ).order_by(OfficialSource.id).all()
        if matching:
            primary = matching[0]
        else:
            primary = OfficialSource(
                region_id=region.id,
                name="四川省工程造价信息网-成都价格信息（官方）",
                url=official_url,
                source_type="api",
                is_official=True,
                is_active=True,
            )
            session.add(primary)
            session.flush()
        primary.name = "四川省工程造价信息网-成都价格信息（官方）"
        primary.source_type = "api"
        primary.is_official = True
        primary.is_active = True
        primary.last_result = ""
        primary.notes = (
            "四川省建设工程造价总站价格信息查询入口；专用接口按成都代码5101查询。"
            "官网未授权时可读取材料目录，但具体单价会显示“会员查看”；支持配置浏览器Cookie。"
        )

        for duplicate in matching[1:]:
            duplicate.is_active = False
            duplicate.last_result = "disabled"
            duplicate.notes = f"与来源 {primary.id} 重复，已自动停用"
        old_urls = {
            "http://202.61.90.35:8037",
            "http://202.61.90.35:8037/",
            "http://www.sceci.net",
            "http://www.sceci.net/",
            "http://cdzj.chengdu.gov.cn",
            "http://cdzj.chengdu.gov.cn/",
        }
        for source in session.query(OfficialSource).filter(OfficialSource.region_id == region.id).all():
            if source.id == primary.id:
                continue
            if source.url.rstrip("/") in {value.rstrip("/") for value in old_urls}:
                source.is_active = False
                source.last_result = "disabled"
                source.notes = "不是成都材料价格的准确查询入口，已由四川官网专用来源替代"
        session.add(AppSettings(key=marker_key, value=str(primary.id)))
        write_audit(session, "update", "official_source", primary.id, "启用四川官网成都价格信息专用接口")
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def _finish_invalid_subscription_cleanup(needs_cleanup):
    marker_key = "repair_invalid_homepage_prices_v1"
    if needs_cleanup:
        backup_database("automatic", "清理旧版本误抓价格前自动备份")

    session = get_session()
    try:
        removed_prices = 0
        removed_history = 0
        removed_materials = 0
        shanghai = session.query(Region).filter(Region.name == "上海").first()
        broad_source = None
        if shanghai:
            broad_source = session.query(OfficialSource).filter(
                OfficialSource.region_id == shanghai.id,
                OfficialSource.url.in_([
                    "http://zjw.sh.gov.cn", "http://zjw.sh.gov.cn/",
                    "https://zjw.sh.gov.cn", "https://zjw.sh.gov.cn/",
                ]),
            ).first()
        if broad_source:
            document_ids = [row[0] for row in session.query(SourceDocument.id).filter(
                SourceDocument.source_id == broad_source.id,
            ).all()]
            if document_ids:
                bad_prices = session.query(MaterialPrice).filter(
                    MaterialPrice.source_doc_id.in_(document_ids),
                    MaterialPrice.is_confirmed.is_(False),
                ).all()
                material_ids = {price.material_id for price in bad_prices}
                bad_keys = {
                    (price.material_id, price.region_id, price.period, price.spec)
                    for price in bad_prices
                }
                for history in session.query(PriceHistory).filter(
                    PriceHistory.source_doc_id.in_(document_ids),
                ).all():
                    key = (history.material_id, history.region_id, history.period, history.spec)
                    if key in bad_keys:
                        session.delete(history)
                        removed_history += 1
                for price in bad_prices:
                    session.delete(price)
                    removed_prices += 1
                session.flush()
                for material_id in material_ids:
                    material = session.query(Material).filter(Material.id == material_id).first()
                    if not material or material.category:
                        continue
                    has_reference = (
                        session.query(MaterialPrice.id).filter(MaterialPrice.material_id == material_id).first()
                        or session.query(PriceHistory.id).filter(PriceHistory.material_id == material_id).first()
                        or session.query(MaterialAlias.id).filter(MaterialAlias.material_id == material_id).first()
                        or session.query(ProjectItem.id).filter(ProjectItem.material_id == material_id).first()
                    )
                    if not has_reference:
                        session.delete(material)
                        removed_materials += 1
            for task in session.query(SubscriptionTask).filter(
                SubscriptionTask.source_id == broad_source.id,
                SubscriptionTask.result_count > 0,
            ).all():
                task.status = "failed"
                task.result_count = 0
                task.message = "旧版本误把非价格资料识别为材料价格，本次已自动清理"
            broad_source.name = "上海市人工、材料、机械信息价（官方）"
            broad_source.url = "https://ciac.zjw.sh.gov.cn/JGBXMGCZJInterWeb/pc/#/HyxxHynr?bmCode=003002"
            broad_source.last_result = ""
            broad_source.notes = "上海住建委官方人工、材料、机械信息价入口；动态网站需要专用接口适配"

        if shanghai:
            old_sources = session.query(OfficialSource).filter(
                OfficialSource.region_id == shanghai.id,
                OfficialSource.url.in_([
                    "http://www.shjjw.gov.cn", "http://www.shjjw.gov.cn/",
                    "https://www.shjjw.gov.cn", "https://www.shjjw.gov.cn/",
                ]),
            ).all()
            for source in old_sources:
                source.is_active = False
                source.last_result = "disabled"
                source.notes = "旧域名已失效，程序已自动停用"

        marker = AppSettings(
            key=marker_key,
            value=json.dumps({
                "removed_prices": removed_prices,
                "removed_history": removed_history,
                "removed_materials": removed_materials,
            }, ensure_ascii=False),
        )
        session.add(marker)
        if removed_prices:
            write_audit(
                session,
                "delete",
                "invalid_subscription_data",
                detail=f"清理旧版本误抓价格 {removed_prices} 条、历史 {removed_history} 条",
            )
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def get_setting(key: str, default: str = "") -> str:
    session = get_session()
    try:
        setting = session.query(AppSettings).filter(AppSettings.key == key).first()
        return setting.value if setting else default
    finally:
        session.close()


def set_settings(values: dict[str, str]):
    session = get_session()
    try:
        for key, value in values.items():
            setting = session.query(AppSettings).filter(AppSettings.key == key).first()
            if setting is None:
                setting = AppSettings(key=key, value=str(value))
                session.add(setting)
            else:
                setting.value = str(value)
        write_audit(session, "update", "app_settings", detail=f"更新设置: {', '.join(values)}")
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def check_database_integrity() -> tuple[bool, str]:
    try:
        with engine.connect() as connection:
            result = connection.execute(text("PRAGMA integrity_check")).scalar()
        return result == "ok", str(result)
    except Exception as error:
        return False, str(error)


def _seed_regions(session: Session):
    regions = [
        Region(name="北京", province="北京", code="110000"),
        Region(name="天津", province="天津", code="120000"),
        Region(name="石家庄", province="河北", code="130100"),
        Region(name="唐山", province="河北", code="130200"),
        Region(name="保定", province="河北", code="130600"),
        Region(name="太原", province="山西", code="140100"),
        Region(name="呼和浩特", province="内蒙古", code="150100"),
        Region(name="包头", province="内蒙古", code="150200"),
        Region(name="沈阳", province="辽宁", code="210100"),
        Region(name="大连", province="辽宁", code="210200"),
        Region(name="长春", province="吉林", code="220100"),
        Region(name="哈尔滨", province="黑龙江", code="230100"),
        Region(name="上海", province="上海", code="310000"),
        Region(name="南京", province="江苏", code="320100"),
        Region(name="苏州", province="江苏", code="320500"),
        Region(name="无锡", province="江苏", code="320200"),
        Region(name="常州", province="江苏", code="320400"),
        Region(name="徐州", province="江苏", code="320300"),
        Region(name="南通", province="江苏", code="320600"),
        Region(name="连云港", province="江苏", code="320700"),
        Region(name="淮安", province="江苏", code="320800"),
        Region(name="盐城", province="江苏", code="320900"),
        Region(name="扬州", province="江苏", code="321000"),
        Region(name="镇江", province="江苏", code="321100"),
        Region(name="泰州", province="江苏", code="321200"),
        Region(name="宿迁", province="江苏", code="321300"),
        Region(name="杭州", province="浙江", code="330100"),
        Region(name="宁波", province="浙江", code="330200"),
        Region(name="温州", province="浙江", code="330300"),
        Region(name="合肥", province="安徽", code="340100"),
        Region(name="福州", province="福建", code="350100"),
        Region(name="厦门", province="福建", code="350200"),
        Region(name="南昌", province="江西", code="360100"),
        Region(name="济南", province="山东", code="370100"),
        Region(name="青岛", province="山东", code="370200"),
        Region(name="烟台", province="山东", code="370600"),
        Region(name="郑州", province="河南", code="410100"),
        Region(name="洛阳", province="河南", code="410300"),
        Region(name="武汉", province="湖北", code="420100"),
        Region(name="宜昌", province="湖北", code="420500"),
        Region(name="长沙", province="湖南", code="430100"),
        Region(name="广州", province="广东", code="440100"),
        Region(name="深圳", province="广东", code="440300"),
        Region(name="东莞", province="广东", code="441900"),
        Region(name="佛山", province="广东", code="440600"),
        Region(name="珠海", province="广东", code="440400"),
        Region(name="南宁", province="广西", code="450100"),
        Region(name="海口", province="海南", code="460100"),
        Region(name="重庆", province="重庆", code="500000"),
        Region(name="成都", province="四川", code="510100"),
        Region(name="绵阳", province="四川", code="510700"),
        Region(name="贵阳", province="贵州", code="520100"),
        Region(name="昆明", province="云南", code="530100"),
        Region(name="拉萨", province="西藏", code="540100"),
        Region(name="西安", province="陕西", code="610100"),
        Region(name="兰州", province="甘肃", code="620100"),
        Region(name="西宁", province="青海", code="630100"),
        Region(name="银川", province="宁夏", code="640100"),
        Region(name="乌鲁木齐", province="新疆", code="650100"),
    ]
    session.add_all(regions)
    session.flush()

    for r in regions:
        std = RegionStandard(
            region_id=r.id, version="2024版",
            name=f"{r.name}通用计价标准",
            tax_rate=0.09, regulation_fee_rate=0.03,
            measure_fee_rate=0.025, profit_rate=0.07,
            is_active=True,
        )
        session.add(std)

    session.flush()
    _seed_default_sources(session)


def _seed_default_sources(session: Session):
    zhengzhou = session.query(Region).filter(Region.name == "郑州").first()
    if zhengzhou:
        sources = [
            OfficialSource(region_id=zhengzhou.id, name="郑州市城乡建设局",
                          url="http://www.zhengzhou.gov.cn/", source_type="web", is_official=True),
            OfficialSource(region_id=zhengzhou.id, name="河南省工程造价信息网",
                          url="https://hnzj.hnjs.gov.cn/", source_type="web", is_official=True),
        ]
        session.add_all(sources)


def _seed_materials(session: Session):
    materials = [
        Material(name="预拌混凝土 C30", category="土建材料", default_unit="m³", spec_template="C30 泵送"),
        Material(name="预拌混凝土 C25", category="土建材料", default_unit="m³", spec_template="C25 泵送"),
        Material(name="螺纹钢 HRB400E Φ20", category="土建材料", default_unit="t", spec_template="HRB400E Φ20"),
        Material(name="螺纹钢 HRB400E Φ25", category="土建材料", default_unit="t", spec_template="HRB400E Φ25"),
        Material(name="圆钢 HPB300 Φ10", category="土建材料", default_unit="t", spec_template="HPB300 Φ10"),
        Material(name="中砂（河砂）", category="土建材料", default_unit="m³", spec_template="河砂"),
        Material(name="碎石 5-31.5mm", category="土建材料", default_unit="m³", spec_template="5-31.5mm"),
        Material(name="普通硅酸盐水泥 P.O 42.5", category="土建材料", default_unit="t", spec_template="P.O 42.5 袋装"),
        Material(name="SBS改性沥青防水卷材 3mm", category="土建材料", default_unit="m²", spec_template="3mm"),
        Material(name="挤塑聚苯板 50mm", category="土建材料", default_unit="m³", spec_template="50mm B1级"),
        Material(name="建筑人工（综合）", category="人工", default_unit="工日", spec_template="综合工日"),
        Material(name="挖掘机 1m³", category="机械", default_unit="台班", spec_template="1m³ 履带式"),
        Material(name="铜芯电缆 YJV-0.6/1 4×25", category="安装材料", default_unit="m", spec_template="YJV-0.6/1 4×25"),
        Material(name="PPR给水管 DN20", category="安装材料", default_unit="m", spec_template="DN20 S4"),
        Material(name="花岗岩路缘石 1000×300×150", category="市政材料", default_unit="m", spec_template="1000×300×150"),
    ]
    session.add_all(materials)


def _seed_fee_rules(session: Session):
    rules = [
        FeeRule(fee_type="直接费", fee_name="人工费", calc_base="工日数×人工单价", sort_order=1),
        FeeRule(fee_type="直接费", fee_name="材料费", calc_base="材料量×材料价", sort_order=2),
        FeeRule(fee_type="直接费", fee_name="机械费", calc_base="台班数×台班价", sort_order=3),
        FeeRule(fee_type="间接费", fee_name="企业管理费", rate=0.05, calc_base="直接费", sort_order=4),
        FeeRule(fee_type="规费", fee_name="社会保险费", rate=0.03, calc_base="人工费", sort_order=5),
        FeeRule(fee_type="规费", fee_name="住房公积金", rate=0.005, calc_base="人工费", sort_order=6),
        FeeRule(fee_type="利润", fee_name="利润", rate=0.07, calc_base="直接费+间接费", sort_order=7),
        FeeRule(fee_type="税金", fee_name="增值税", rate=0.09, calc_base="税前造价", sort_order=8),
        FeeRule(fee_type="措施费", fee_name="安全文明施工费", rate=0.025, calc_base="直接费", sort_order=9),
    ]
    session.add_all(rules)


def _seed_material_prices(session: Session):
    source_file = config.BUILTIN_INFO_PRICE_FILE
    if source_file.exists():
        from .utils import import_prices_from_excel

        result = import_prices_from_excel(str(source_file), session)
        if not result.get("success"):
            raise RuntimeError(result.get("error", "内置信息价导入失败"))
        return
    _seed_prices(session)

def _seed_prices(session: Session):
    zhengzhou = session.query(Region).filter(Region.name == "郑州").first()
    if not zhengzhou:
        return
    price_rows = [
        ("预拌混凝土 C30", "C30 泵送", "m³", "2026-04", 420.00, True, False),
        ("预拌混凝土 C30", "C30 泵送", "m³", "2026-05", 419.20, True, False),
        ("预拌混凝土 C30", "C30 泵送", "m³", "2026-06", 428.00, True, False),
        ("螺纹钢 HRB400E Φ20", "HRB400E Φ20", "t", "2026-05", 3915.00, True, False),
        ("螺纹钢 HRB400E Φ20", "HRB400E Φ20", "t", "2026-06", 3860.00, True, False),
        ("中砂（河砂）", "河砂", "m³", "2026-06", 156.00, False, True),
        ("普通硅酸盐水泥 P.O 42.5", "P.O 42.5 袋装", "t", "2026-06", 380.00, True, False),
        ("建筑人工（综合）", "综合工日", "工日", "2026-06", 280.00, True, False),
    ]
    for material_name, spec, unit, period, price, confirmed, anomaly in price_rows:
        material = session.query(Material).filter(Material.name == material_name).first()
        if not material:
            continue
        material_price = MaterialPrice(
            material_id=material.id,
            region_id=zhengzhou.id,
            period=period,
            price=price,
            unit=unit,
            spec=spec,
            trust_level="official",
            is_confirmed=confirmed,
            is_anomaly=anomaly,
            anomaly_reason="单月涨幅偏高，需人工复核" if anomaly else "",
        )
        session.add(material_price)
        session.add(PriceHistory(
            material_id=material.id,
            region_id=zhengzhou.id,
            period=period,
            price=price,
            unit=unit,
            spec=spec,
            trust_level="official",
            notes="种子样例数据",
        ))


def _seed_projects(session: Session):
    zhengzhou = session.query(Region).filter(Region.name == "郑州").first()
    project = Project(
        name="郑州示范办公楼项目",
        code="DS-DEMO-001",
        client="示范建设单位",
        project_type="公共建筑",
        scale="中型",
        area=4200.0,
        structure_type="钢筋混凝土框架",
        project_location="河南·郑州",
        pricing_province="河南",
        pricing_city="郑州",
        price_year="2026",
        stage="投标报价",
        specialty="建筑与装饰工程",
        region_id=zhengzhou.id if zhengzhou else None,
        status="locked",
        locked_version="2026-06",
        total_amount=354795.46,
        notes="示范项目，可删除或作为模板参考",
    )
    session.add(project)
    session.flush()
    items = [
        (1, "010101001001", "平整场地", "", "m2", 1500.00, 3.50, 5250.00),
        (2, "010502001001", "现浇C30砼柱", "C30", "m3", 120.00, 856.00, 102720.00),
        (3, "010515001001", "现浇构件钢筋", "HRB400 D12", "t", 15.00, 4250.00, 63750.00),
        (4, "010401001001", "砖基础", "MU10", "m3", 80.00, 350.00, 28000.00),
        (5, "011101001001", "水泥砂浆楼地面", "20mm", "m2", 2000.00, 45.00, 90000.00),
    ]
    item_dicts = []
    for seq_no, item_code, item_name, spec, unit, quantity, direct_unit_price, _direct_total in items:
        labor_unit_price = direct_unit_price * 0.35
        material_unit_price = direct_unit_price * 0.55
        machinery_unit_price = direct_unit_price * 0.10
        management_unit_price = direct_unit_price * 0.05
        profit_unit_price = (direct_unit_price + management_unit_price) * 0.07
        unit_price = direct_unit_price + management_unit_price + profit_unit_price
        total_price = quantity * unit_price
        project_item = ProjectItem(
            project_id=project.id,
            seq_no=seq_no,
            item_code=item_code,
            item_name=item_name,
            spec=spec,
            unit=unit,
            quantity=quantity,
            labor_unit_price=labor_unit_price,
            material_unit_price=material_unit_price,
            machinery_unit_price=machinery_unit_price,
            management_unit_price=management_unit_price,
            profit_unit_price=profit_unit_price,
            unit_price=unit_price,
            total_price=total_price,
        )
        session.add(project_item)
        item_dicts.append({
            "seq_no": seq_no,
            "item_code": item_code,
            "item_name": item_name,
            "spec": spec,
            "unit": unit,
            "quantity": quantity,
            "labor_unit_price": labor_unit_price,
            "material_unit_price": material_unit_price,
            "machinery_unit_price": machinery_unit_price,
            "unallocated_unit_price": 0.0,
            "management_unit_price": management_unit_price,
            "profit_unit_price": profit_unit_price,
            "unit_price": unit_price,
            "total_price": total_price,
        })
    session.add(ProjectPriceSnapshot(
        project_id=project.id,
        version="2026-06",
        total_amount=354795.46,
        data_json=json.dumps(item_dicts, ensure_ascii=False),
        notes="种子项目快照",
    ))


def write_audit(session: Session, action: str, entity_type: str, entity_id: int = None, detail: str = ""):
    log = AuditLog(action=action, entity_type=entity_type, entity_id=entity_id, detail=detail)
    session.add(log)


def backup_database(backup_type: str = "manual", notes: str = "") -> Optional[Path]:
    db_path = config.db_path()
    if not db_path.exists():
        return None
    config.ensure_dirs()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"dashuo_cost_cloud_{ts}.db"
    backup_path = config.BACKUP_DIR / backup_name
    try:
        with sqlite3.connect(db_path) as source_connection:
            with sqlite3.connect(backup_path) as backup_connection:
                source_connection.backup(backup_connection)
        size = backup_path.stat().st_size
        session = get_session()
        try:
            rec = BackupRecord(
                file_name=backup_name,
                file_path=str(backup_path),
                file_size=size,
                backup_type=backup_type,
                notes=notes,
            )
            session.add(rec)
            write_audit(session, "backup", "database", detail=f"备份: {backup_name}")
            session.commit()
            _cleanup_old_backups(session)
        except Exception:
            session.rollback()
            raise
        finally:
            session.close()
        return backup_path
    except OSError:
        return None


def _cleanup_old_backups(session: Session):
    records = session.query(BackupRecord).order_by(BackupRecord.created_at.desc()).all()
    if len(records) > config.MAX_BACKUPS:
        for rec in records[config.MAX_BACKUPS:]:
            path = Path(rec.file_path)
            if path.exists():
                path.unlink(missing_ok=True)
            session.delete(rec)
        session.commit()


def restore_backup(backup_id: int) -> bool:
    session = get_session()
    try:
        rec = session.query(BackupRecord).filter(BackupRecord.id == backup_id).first()
        if not rec:
            return False
        backup_path = Path(rec.file_path)
        if not backup_path.exists():
            return False
        backup_name = rec.file_name
    finally:
        session.close()
    db_path = config.db_path()
    safety_copy = db_path.with_name(f"{db_path.stem}_before_restore_{datetime.now():%Y%m%d_%H%M%S}.db")
    try:
        engine.dispose()
        if db_path.exists():
            shutil.copy2(db_path, safety_copy)
        shutil.copy2(backup_path, db_path)
        with engine.connect() as connection:
            if connection.execute(text("PRAGMA integrity_check")).scalar() != "ok":
                raise RuntimeError("恢复后的数据库完整性检查失败")
        session = get_session()
        try:
            write_audit(session, "restore", "database", detail=f"恢复备份: {backup_name}")
            session.commit()
        finally:
            session.close()
        return True
    except Exception:
        engine.dispose()
        if safety_copy.exists():
            shutil.copy2(safety_copy, db_path)
        return False


def get_db_stats() -> dict:
    session = get_session()
    try:
        return {
            "regions": session.query(func.count(Region.id)).scalar(),
            "materials": session.query(func.count(Material.id)).scalar(),
            "cost_items": session.query(func.count(CostItem.id)).scalar(),
            "quota_items": session.query(func.count(QuotaItem.id)).scalar(),
            "quota_compositions": session.query(func.count(QuotaComposition.id)).scalar(),
            "prices": session.query(func.count(MaterialPrice.id)).scalar(),
            "projects": session.query(func.count(Project.id)).scalar(),
            "sources": session.query(func.count(OfficialSource.id)).scalar(),
            "backups": session.query(func.count(BackupRecord.id)).scalar(),
            "pending_prices": session.query(func.count(MaterialPrice.id)).filter(
                MaterialPrice.is_confirmed == False).scalar(),
            "anomaly_prices": session.query(func.count(MaterialPrice.id)).filter(
                MaterialPrice.is_anomaly == True).scalar(),
        }
    finally:
        session.close()
