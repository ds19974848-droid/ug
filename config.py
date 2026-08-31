"""全局配置管理。"""
import os
import sqlite3
import shutil
import sys
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else BASE_DIR
DEFAULT_DATA_DIR = Path(os.getenv("APPDATA") or Path.home() / "AppData" / "Roaming") / "DashuoCostCloud"
DATA_LOCATION_FILE = DEFAULT_DATA_DIR / "data_location.txt"


def _saved_data_dir() -> Path | None:
    try:
        value = DATA_LOCATION_FILE.read_text(encoding="utf-8").strip()
    except (OSError, UnicodeError):
        return None
    return Path(value).expanduser() if value else None


USER_DATA_DIR = Path(os.getenv("DASHUO_DATA_DIR") or _saved_data_dir() or DEFAULT_DATA_DIR)

load_dotenv(APP_DIR / ".env")
load_dotenv(USER_DATA_DIR / ".env")


class Config:
    APP_NAME = "智能工程造价辅助系统"
    VERSION = "0.5.3"

    USER_DATA_DIR = USER_DATA_DIR
    DATA_LOCATION_FILE = DATA_LOCATION_FILE
    LOG_DIR = USER_DATA_DIR / "logs"
    SOURCE_DIR = USER_DATA_DIR / "source_files"
    LICENSE_REQUIRED_FILE = USER_DATA_DIR / "require_license.flag"
    BUILTIN_COST_ITEMS_FILE = BUNDLE_DIR / "assets" / "工程价格数据库_4453条.xlsx"
    BUILTIN_INFO_PRICE_FILE = BUNDLE_DIR / "assets" / "信息价数据库.xlsx"
    BUILTIN_QUOTA_FILE = BUNDLE_DIR / "assets" / "定额库.xlsx"
    BUILTIN_ENTERPRISE_REFERENCE_FILE = BUNDLE_DIR / "assets" / "内蒙古区域企业工程高频项目参考定额8.25(1).xlsx"
    APP_ICON = BUNDLE_DIR / "assets" / "intelligent-cost-logo.svg"
    USING_DEFAULT_DATABASE = os.getenv("DATABASE_URL") is None
    DATABASE_URL = os.getenv(
        "DATABASE_URL",
        f"sqlite:///{(USER_DATA_DIR / 'dashuo_cost_cloud.db').as_posix()}",
    )

    DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
    DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
    DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
    AI_WEB_SEARCH_ENABLED = os.getenv("AI_WEB_SEARCH_ENABLED", "true").lower() != "false"
    AI_WEB_SEARCH_TIMEOUT = int(os.getenv("AI_WEB_SEARCH_TIMEOUT", "20"))

    SUBSCRIPTION_TIMEOUT = int(os.getenv("SUBSCRIPTION_TIMEOUT", "30"))
    MAX_SUBSCRIPTION_RETRIES = int(os.getenv("MAX_SUBSCRIPTION_RETRIES", "3"))
    MAX_SOURCE_PAGES = int(os.getenv("MAX_SOURCE_PAGES", "80"))
    MAX_SOURCE_DOCUMENTS = int(os.getenv("MAX_SOURCE_DOCUMENTS", "300"))
    # 信息价库只保留每个城市最新一期，并限制城市数量，避免联网抓取持续膨胀。
    MAX_PRICE_REGIONS = int(os.getenv("MAX_PRICE_REGIONS", "5"))
    KEEP_LATEST_PRICE_PERIOD_ONLY = os.getenv(
        "KEEP_LATEST_PRICE_PERIOD_ONLY", "true"
    ).lower() != "false"
    USER_AGENT = os.getenv(
        "USER_AGENT",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    )

    PRICE_CHANGE_THRESHOLD = float(os.getenv("PRICE_CHANGE_THRESHOLD", "0.30"))
    PRICE_MIN_SANE = float(os.getenv("PRICE_MIN_SANE", "0.01"))
    PRICE_MAX_SANE = float(os.getenv("PRICE_MAX_SANE", "100000000"))

    BACKUP_DIR = USER_DATA_DIR / "backups"
    MAX_BACKUPS = int(os.getenv("MAX_BACKUPS", "30"))
    AUTO_BACKUP_ENABLED = os.getenv("AUTO_BACKUP_ENABLED", "true").lower() == "true"

    DEFAULT_OFFICIAL_SOURCES = [
        {
            "name": "郑州市城乡建设局",
            "url": "http://www.zhengzhou.gov.cn/",
            "region": "郑州",
            "province": "河南",
        },
        {
            "name": "河南省工程造价信息网",
            "url": "https://hnzj.hnjs.gov.cn/",
            "region": "郑州",
            "province": "河南",
        },
    ]

    OFFICIAL_DOMAIN_PATTERNS = [
        ".gov.cn", "zjw.", "zjz.", "cost.", "zaojia.", "cace.",
    ]

    MATERIAL_CATEGORIES = [
        "土建材料", "装饰材料", "安装材料", "市政材料",
        "园林材料", "人工", "机械", "其他",
    ]

    FEE_TYPES = [
        "直接费", "间接费", "利润", "税金", "规费", "措施费", "其他费用",
    ]

    AUDIT_ACTION_TYPES = [
        "create", "update", "delete", "import", "export",
        "confirm", "restore", "backup", "login", "logout",
    ]

    @classmethod
    def db_path(cls) -> Path:
        url = cls.DATABASE_URL
        if url.startswith("sqlite:///"):
            return Path(url.replace("sqlite:///", ""))
        return Path(url)

    @classmethod
    def ensure_dirs(cls):
        cls.USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
        cls.BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        cls.LOG_DIR.mkdir(parents=True, exist_ok=True)
        cls.SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        cls._migrate_legacy_db()

    @classmethod
    def configure_data_directory(cls, directory: str | Path) -> tuple[bool, str]:
        """Prepare a new local data directory and persist it for the next restart."""
        target = Path(directory).expanduser().resolve()
        current = cls.USER_DATA_DIR.expanduser().resolve()
        if target == current:
            return True, "当前已经使用该数据目录。"
        if target == current or current in target.parents:
            return False, "新数据目录不能放在当前数据目录内部，以免复制时形成递归目录。"
        try:
            target.mkdir(parents=True, exist_ok=True)
            target_db = target / "dashuo_cost_cloud.db"
            current_db = cls.db_path()
            existing_database = target_db.exists()
            if current_db.exists() and not existing_database:
                source_connection = sqlite3.connect(str(current_db), timeout=15)
                destination_connection = sqlite3.connect(str(target_db), timeout=15)
                try:
                    source_connection.backup(destination_connection)
                finally:
                    destination_connection.close()
                    source_connection.close()

            for folder_name in ("source_files", "backups"):
                source_folder = current / folder_name
                target_folder = target / folder_name
                if source_folder.exists() and source_folder.resolve() != target_folder.resolve():
                    shutil.copytree(source_folder, target_folder, dirs_exist_ok=True)

            cls.DATA_LOCATION_FILE.parent.mkdir(parents=True, exist_ok=True)
            cls.DATA_LOCATION_FILE.write_text(str(target), encoding="utf-8")
            if existing_database:
                return True, (
                    f"目标目录已有 dashuo_cost_cloud.db，未覆盖该文件；来源文件和备份已同步。"
                    f"重启后将使用：{target}"
                )
            return True, f"数据已复制到：{target}。重启后将使用新目录。"
        except Exception as error:
            return False, f"数据目录切换准备失败：{error}"

    @classmethod
    def _migrate_legacy_db(cls):
        if not cls.USING_DEFAULT_DATABASE:
            return
        db_path = cls.db_path()
        legacy_path = BASE_DIR / "dashuo_cost_cloud.db"
        if db_path.exists() or not legacy_path.exists() or legacy_path == db_path:
            return
        try:
            shutil.copy2(legacy_path, db_path)
        except OSError:
            pass


config = Config()
