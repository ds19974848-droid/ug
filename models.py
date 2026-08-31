"""SQLAlchemy 数据库模型定义"""
from datetime import datetime
from sqlalchemy import (
    Column, Integer, String, Float, Text, Boolean,
    DateTime, ForeignKey, create_engine, event, Index,
)
from sqlalchemy.orm import DeclarativeBase, relationship, Session
from sqlalchemy.engine import Engine


class Base(DeclarativeBase):
    pass


@event.listens_for(Engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.execute("PRAGMA synchronous=NORMAL")
    cursor.close()


now = datetime.utcnow


class Region(Base):
    __tablename__ = "regions"
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(100), nullable=False, unique=True, comment="地区名称")
    province = Column(String(100), default="", comment="所属省份")
    code = Column(String(20), default="", comment="行政区划代码")
    is_active = Column(Boolean, default=True, comment="是否启用")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    standards = relationship("RegionStandard", back_populates="region", cascade="all, delete-orphan")
    subscription_sources = relationship("OfficialSource", back_populates="region_obj")
    prices = relationship("MaterialPrice", back_populates="region_obj")


class RegionStandard(Base):
    __tablename__ = "region_standards"
    id = Column(Integer, primary_key=True, autoincrement=True)
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=False)
    version = Column(String(50), nullable=False, comment="标准版本号")
    name = Column(String(200), default="", comment="标准名称")
    tax_rate = Column(Float, default=0.0, comment="税率")
    regulation_fee_rate = Column(Float, default=0.0, comment="规费费率")
    measure_fee_rate = Column(Float, default=0.0, comment="措施费费率")
    profit_rate = Column(Float, default=0.0, comment="利润率")
    is_active = Column(Boolean, default=False)
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    region = relationship("Region", back_populates="standards")


class FeeRule(Base):
    __tablename__ = "fee_rules"
    id = Column(Integer, primary_key=True, autoincrement=True)
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=True)
    fee_type = Column(String(100), nullable=False, comment="费用类型")
    fee_name = Column(String(200), default="", comment="费用名称")
    rate = Column(Float, default=0.0, comment="费率")
    calc_base = Column(String(100), default="", comment="取费基数")
    sort_order = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)


class Material(Base):
    __tablename__ = "materials"
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(300), nullable=False, unique=True, comment="标准化材料名称")
    category = Column(String(100), default="", comment="材料分类")
    default_unit = Column(String(50), default="", comment="默认单位")
    spec_template = Column(String(200), default="", comment="规格模板")
    description = Column(Text, default="", comment="描述")
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    aliases = relationship("MaterialAlias", back_populates="material", cascade="all, delete-orphan")
    prices = relationship("MaterialPrice", back_populates="material")


class MaterialAlias(Base):
    __tablename__ = "material_aliases"
    id = Column(Integer, primary_key=True, autoincrement=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    alias_name = Column(String(300), nullable=False, comment="别名")
    source = Column(String(200), default="", comment="别名来源")
    created_at = Column(DateTime, default=now)

    material = relationship("Material", back_populates="aliases")


class MaterialPrice(Base):
    __tablename__ = "material_prices"
    id = Column(Integer, primary_key=True, autoincrement=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=False)
    period = Column(String(20), nullable=False, comment="期数，如 2026-06")
    price = Column(Float, nullable=False, comment="价格")
    unit = Column(String(50), default="", comment="单位")
    spec = Column(String(200), default="", comment="规格")
    source_key = Column(String(300), default="", comment="来源行标识；用于保留同名同规格的不同价格")
    trust_level = Column(String(50), default="official", comment="可信等级")
    source_type = Column(String(50), default="official", comment="来源类型: official/market_reference/manual")
    price_basis = Column(String(50), default="as_published", comment="价格依据: 含税/除税/到场/出厂/原文")
    valid_from = Column(String(20), default="", comment="有效期起始")
    valid_to = Column(String(20), default="", comment="有效期结束")
    is_withdrawn = Column(Boolean, default=False, comment="是否撤回")
    source_doc_id = Column(Integer, ForeignKey("source_documents.id"), nullable=True)
    is_confirmed = Column(Boolean, default=False, comment="是否已确认")
    is_anomaly = Column(Boolean, default=False, comment="是否异常")
    anomaly_reason = Column(Text, default="", comment="异常原因")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    material = relationship("Material", back_populates="prices")
    region_obj = relationship("Region", back_populates="prices")

    __table_args__ = (
        Index(
            "idx_price_material_region_period_spec",
            "material_id", "region_id", "period", "spec", "source_key",
            unique=True,
        ),
    )


class CostItem(Base):
    """工程清单与综合价格参考库。"""

    __tablename__ = "cost_items"
    id = Column(Integer, primary_key=True, autoincrement=True)
    source_key = Column(String(120), nullable=False, unique=True)
    data_source = Column(String(100), default="")
    original_seq = Column(Integer, default=0)
    item_name = Column(String(500), nullable=False)
    features = Column(Text, default="")
    unit = Column(String(50), default="")
    comprehensive_price = Column(Float, default=0.0)
    tax_inclusive_price = Column(Float, default=0.0)
    labor_cost = Column(Float, default=0.0)
    material_cost = Column(Float, default=0.0)
    machinery_cost = Column(Float, default=0.0)
    management_cost = Column(Float, default=0.0)
    profit = Column(Float, default=0.0)
    full_category = Column(String(300), default="")
    original_category = Column(String(200), default="")
    price_analysis = Column(Text, default="")
    ai_analysis = Column(Text, default="")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    __table_args__ = (
        Index("idx_cost_item_name", "item_name"),
        Index("idx_cost_item_source_category", "data_source", "original_category"),
    )


class ProjectListData(Base):
    """Historical project BOQ rows used as project-level pricing evidence."""

    __tablename__ = "project_list_data"
    id = Column(Integer, primary_key=True, autoincrement=True)
    source_file = Column(String(500), default="")
    source_sheet = Column(String(200), default="")
    source_project = Column(String(300), default="")
    region = Column(String(100), default="")
    period = Column(String(20), default="")
    seq_no = Column(String(50), default="")
    item_code = Column(String(120), default="")
    item_name = Column(String(500), nullable=False)
    feature = Column(Text, default="")
    unit = Column(String(50), default="")
    quantity = Column(Float, nullable=True)
    analysis = Column(Text, default="")
    comprehensive_price = Column(Float, nullable=True)
    total_price = Column(Float, nullable=True, comment="合价/总价")
    labor_cost = Column(Float, nullable=True)
    material_cost = Column(Float, nullable=True)
    machinery_cost = Column(Float, nullable=True)
    management_cost = Column(Float, nullable=True)
    profit = Column(Float, nullable=True)
    note = Column(Text, default="")
    raw_data = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    __table_args__ = (
        Index("idx_project_list_data_name", "item_name"),
        Index("idx_project_list_data_source", "source_project", "region", "period"),
    )


class QuotaItem(Base):
    """成本定额库主表：一条定额子目及其专业/价格信息。"""

    __tablename__ = "quota_items"
    id = Column(Integer, primary_key=True, autoincrement=True)
    source_key = Column(String(300), nullable=False, unique=True)
    major = Column(String(100), default="装修", comment="定额专业，如装修/土建/安装")
    code = Column(String(120), default="", comment="定额编码")
    name = Column(String(500), nullable=False, comment="定额名称")
    feature = Column(Text, default="", comment="工作内容/项目特征")
    unit = Column(String(50), default="", comment="计量单位")
    tax_price = Column(Float, default=0.0, comment="含税单价")
    no_tax_price = Column(Float, default=0.0, comment="除税单价")
    category = Column(String(200), default="", comment="备注/分类，如地面/墙面")
    source = Column(String(200), default="", comment="数据来源")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    compositions = relationship(
        "QuotaComposition",
        back_populates="quota_item",
        cascade="all, delete-orphan",
        order_by="QuotaComposition.sort_order",
    )

    __table_args__ = (
        Index("idx_quota_item_major_category", "major", "category"),
        Index("idx_quota_item_name", "name"),
    )


class QuotaComposition(Base):
    """定额库组成明细：人工/材料/机械/分包等工料机子目。"""

    __tablename__ = "quota_compositions"
    id = Column(Integer, primary_key=True, autoincrement=True)
    quota_item_id = Column(
        Integer,
        ForeignKey("quota_items.id", ondelete="CASCADE"),
        nullable=False,
    )
    sort_order = Column(Integer, default=0)
    category = Column(String(100), default="", comment="费用类别，如人工费/材料费")
    code = Column(String(120), default="", comment="工料机编码")
    name = Column(String(300), default="", comment="工料机名称")
    feature = Column(Text, default="", comment="工作内容")
    unit = Column(String(50), default="", comment="单位")
    qty = Column(Float, default=0.0, comment="单位含量")
    loss_rate = Column(Float, default=0.0, comment="损耗率%，如 3 表示 3%")
    no_tax_price = Column(Float, default=0.0, comment="除税单价")
    tax_rate = Column(Float, default=0.0, comment="税率%，如 13 表示 13%")
    tax_price = Column(Float, default=0.0, comment="含税单价")
    no_tax_total = Column(Float, default=0.0, comment="除税合价")
    tax_total = Column(Float, default=0.0, comment="含税合价")
    note = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    quota_item = relationship("QuotaItem", back_populates="compositions")

    __table_args__ = (
        Index("idx_quota_composition_item", "quota_item_id"),
    )


class PriceHistory(Base):
    __tablename__ = "price_history"
    id = Column(Integer, primary_key=True, autoincrement=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=False)
    period = Column(String(20), nullable=False)
    price = Column(Float, nullable=False)
    unit = Column(String(50), default="")
    spec = Column(String(200), default="")
    trust_level = Column(String(50), default="official")
    source_type = Column(String(50), default="official")
    price_basis = Column(String(50), default="as_published")
    source_doc_id = Column(Integer, ForeignKey("source_documents.id"), nullable=True)
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)

    __table_args__ = (
        Index("idx_hist_material_region_period", "material_id", "region_id", "period"),
    )


class OfficialSource(Base):
    __tablename__ = "official_sources"
    id = Column(Integer, primary_key=True, autoincrement=True)
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=False)
    name = Column(String(300), nullable=False, comment="来源名称")
    url = Column(String(500), default="", comment="网址")
    source_type = Column(String(50), default="web", comment="来源类型: web/pdf/excel")
    is_official = Column(Boolean, default=True, comment="是否官方来源")
    is_active = Column(Boolean, default=True)
    last_check_at = Column(DateTime, nullable=True)
    last_result = Column(String(100), default="")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    region_obj = relationship("Region", back_populates="subscription_sources")


class SubscriptionTask(Base):
    __tablename__ = "subscription_tasks"
    id = Column(Integer, primary_key=True, autoincrement=True)
    source_id = Column(Integer, ForeignKey("official_sources.id"), nullable=False)
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=False)
    period = Column(String(20), default="", comment="目标期数")
    status = Column(String(50), default="pending", comment="pending/running/success/failed/manual")
    progress = Column(Integer, default=0, comment="进度 0-100")
    message = Column(Text, default="")
    result_count = Column(Integer, default=0)
    parsed_count = Column(Integer, default=0)
    stored_count = Column(Integer, default=0)
    rejected_count = Column(Integer, default=0)
    phase = Column(String(50), default="")
    failure_stage = Column(String(50), default="")
    failure_reason = Column(Text, default="")
    next_action = Column(Text, default="")
    anomaly_count = Column(Integer, default=0)
    started_at = Column(DateTime, nullable=True)
    finished_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)


class SourceDocument(Base):
    __tablename__ = "source_documents"
    id = Column(Integer, primary_key=True, autoincrement=True)
    task_id = Column(Integer, ForeignKey("subscription_tasks.id"), nullable=True)
    source_id = Column(Integer, ForeignKey("official_sources.id"), nullable=True)
    file_name = Column(String(500), default="", comment="原始文件名")
    file_path = Column(String(500), default="", comment="本地保存路径")
    file_type = Column(String(50), default="", comment="文件类型")
    file_size = Column(Integer, default=0)
    period = Column(String(20), default="")
    url = Column(String(500), default="")
    is_parsed = Column(Boolean, default=False)
    parse_result = Column(Text, default="")
    created_at = Column(DateTime, default=now)


class AiParseLog(Base):
    __tablename__ = "ai_parse_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    task_id = Column(Integer, ForeignKey("subscription_tasks.id"), nullable=True)
    doc_id = Column(Integer, ForeignKey("source_documents.id"), nullable=True)
    model = Column(String(100), default="deepseek-chat")
    prompt_tokens = Column(Integer, default=0)
    completion_tokens = Column(Integer, default=0)
    input_text = Column(Text, default="")
    output_text = Column(Text, default="")
    success = Column(Boolean, default=False)
    error_message = Column(Text, default="")
    created_at = Column(DateTime, default=now)


class Project(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(300), nullable=False, comment="项目名称")
    code = Column(String(100), default="", comment="项目编号")
    client = Column(String(300), default="", comment="建设单位/客户")
    design_unit = Column(String(300), default="", comment="设计单位")
    contractor = Column(String(300), default="", comment="施工单位")
    supervision_unit = Column(String(300), default="", comment="监理单位")
    tender_agent = Column(String(300), default="", comment="招标代理/咨询单位")
    project_type = Column(String(100), default="", comment="项目类型")
    funding_source = Column(String(100), default="", comment="资金来源")
    contract_type = Column(String(100), default="", comment="合同类型")
    scale = Column(String(50), default="", comment="规模等级")
    daily_capacity = Column(Float, default=0.0, comment="日处理量，万立方米/日")
    area = Column(Float, default=0.0, comment="建筑面积，平方米")
    structure_type = Column(String(100), default="", comment="结构形式")
    process_type = Column(String(100), default="", comment="工艺类型")
    project_location = Column(String(300), default="", comment="项目所在地")
    project_address = Column(String(500), default="", comment="项目详细地址")
    pricing_province = Column(String(100), default="", comment="计价省份")
    pricing_city = Column(String(100), default="", comment="计价城市")
    pricing_district = Column(String(100), default="", comment="计价区县")
    pricing_date = Column(String(20), default="", comment="计价日期")
    price_year = Column(String(10), default="", comment="价格年份")
    stage = Column(String(50), default="投标报价", comment="计价阶段")
    pricing_basis = Column(String(300), default="", comment="计价依据")
    boq_basis = Column(String(300), default="", comment="清单编制依据")
    tax_method = Column(String(100), default="", comment="税价口径")
    planned_start_date = Column(String(20), default="", comment="计划开工日期")
    planned_end_date = Column(String(20), default="", comment="计划完工日期")
    specialty = Column(String(100), default="", comment="造价专业")
    currency = Column(String(10), default="CNY", comment="项目本位币")
    region_id = Column(Integer, ForeignKey("regions.id"), nullable=True)
    standard_id = Column(Integer, ForeignKey("region_standards.id"), nullable=True)
    status = Column(String(50), default="draft", comment="draft/locked/archived")
    locked_version = Column(String(100), default="", comment="锁定时的价格版本")
    total_amount = Column(Float, default=0.0, comment="总造价")
    management_rate = Column(Float, default=0.05, comment="管理费率")
    management_base = Column(String(50), default="direct", comment="管理费计取基础")
    profit_rate = Column(Float, default=0.07, comment="利润率")
    profit_base = Column(String(50), default="direct_management", comment="利润计取基础")
    tax_rate = Column(Float, default=0.09, comment="税率")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    items = relationship("ProjectItem", back_populates="project", cascade="all, delete-orphan")
    snapshots = relationship("ProjectPriceSnapshot", back_populates="project", cascade="all, delete-orphan")
    quota_matches = relationship("ProjectQuotaMatch", cascade="all, delete-orphan")


class ProjectItem(Base):
    __tablename__ = "project_items"
    id = Column(Integer, primary_key=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id", ondelete="CASCADE"), nullable=False)
    seq_no = Column(Integer, default=0, comment="序号")
    item_code = Column(String(100), default="", comment="清单编号/定额编号")
    item_name = Column(String(500), default="", comment="清单名称")
    spec = Column(String(200), default="", comment="规格")
    unit = Column(String(50), default="", comment="单位")
    quantity = Column(Float, default=0.0, comment="工程量")
    labor_unit_price = Column(Float, default=0.0, comment="人工费单价")
    material_unit_price = Column(Float, default=0.0, comment="材料费单价")
    machinery_unit_price = Column(Float, default=0.0, comment="机械费单价")
    unallocated_unit_price = Column(Float, default=0.0, comment="未拆分直接费单价")
    management_unit_price = Column(Float, default=0.0, comment="管理费单价")
    profit_unit_price = Column(Float, default=0.0, comment="利润单价")
    unit_price = Column(Float, default=0.0, comment="综合单价")
    total_price = Column(Float, default=0.0, comment="合价")
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=True)
    price_source = Column(String(50), default="manual", comment="价格来源")
    custom_data = Column(Text, default="", comment="导入表格自定义字段 JSON")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    project = relationship("Project", back_populates="items")


class ProjectQuotaMatch(Base):
    """项目工作台的定额匹配结果，与项目报价明细分开保存。"""

    __tablename__ = "project_quota_matches"
    id = Column(Integer, primary_key=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id", ondelete="CASCADE"), nullable=False)
    source_file = Column(String(500), default="")
    data_json = Column(Text, default="")
    total_rows = Column(Integer, default=0)
    matched_rows = Column(Integer, default=0)
    quota_count = Column(Integer, default=0)
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    __table_args__ = (
        Index("idx_project_quota_match_project", "project_id"),
    )


class ProjectPriceSnapshot(Base):
    __tablename__ = "project_price_snapshots"
    id = Column(Integer, primary_key=True, autoincrement=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False)
    version = Column(String(100), nullable=False, comment="快照版本")
    total_amount = Column(Float, default=0.0, comment="快照总造价")
    data_json = Column(Text, default="", comment="快照详细数据 JSON")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)

    project = relationship("Project", back_populates="snapshots")


class AuditLog(Base):
    __tablename__ = "audit_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    action = Column(String(50), default="", comment="操作类型")
    entity_type = Column(String(100), default="", comment="实体类型")
    entity_id = Column(Integer, nullable=True)
    detail = Column(Text, default="", comment="操作详情")
    user_name = Column(String(100), default="admin")
    ip_address = Column(String(50), default="")
    created_at = Column(DateTime, default=now)


class BackupRecord(Base):
    __tablename__ = "backups"
    id = Column(Integer, primary_key=True, autoincrement=True)
    file_name = Column(String(500), default="")
    file_path = Column(String(500), default="")
    file_size = Column(Integer, default=0)
    backup_type = Column(String(50), default="auto", comment="auto/manual")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)


class AppSettings(Base):
    __tablename__ = "app_settings"
    id = Column(Integer, primary_key=True, autoincrement=True)
    key = Column(String(200), unique=True, nullable=False)
    value = Column(Text, default="")
    updated_at = Column(DateTime, default=now, onupdate=now)


class LicenseGrant(Base):
    """授权码生成与激活记录，兼容外置授权管理工具。"""

    __tablename__ = "license_grants"
    id = Column(Integer, primary_key=True, autoincrement=True)
    license_key = Column(String(500), unique=True, nullable=False)
    machine_code = Column(String(50), default="", comment="机器码（16位大写十六进制）")
    plan_label = Column(String(50), default="", comment="授权时长，如 1天/7天/一个月/永久")
    days = Column(Integer, default=0, comment="授权天数，0 表示永久")
    issued_at = Column(DateTime, default=now)
    expires_at = Column(DateTime, nullable=True)
    status = Column(String(50), default="issued", comment="issued/activated/revoked/expired")
    recipient = Column(String(300), default="", comment="被授权人/备注")
    activated_at = Column(DateTime, nullable=True)
    machine_name = Column(String(200), default="")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=now)
    updated_at = Column(DateTime, default=now, onupdate=now)

    __table_args__ = (
        Index("idx_license_grants_machine_code", "machine_code"),
        Index("idx_license_grants_status", "status"),
    )
