"""大硕造价云库 - 程序入口"""
import logging
import sys
import os
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.config import config


def _setup_logging():
    config.ensure_dirs()
    logging.basicConfig(
        filename=config.LOG_DIR / "app.log",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        encoding="utf-8",
        force=True,
    )


def _show_startup_error(error: Exception):
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox

        QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(
            None,
            "智能工程造价辅助系统启动失败",
            f"软件启动失败：{error}\n\n错误日志：\n{config.LOG_DIR / 'app.log'}",
        )
    except Exception:
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(
                0,
                f"软件启动失败：{error}\n\n错误日志：\n{config.LOG_DIR / 'app.log'}",
                "智能工程造价辅助系统启动失败",
                0x10,
            )
        except Exception:
            pass


def _handle_exception(exc_type, exc_value, exc_traceback):
    logging.critical(
        "Unhandled exception",
        exc_info=(exc_type, exc_value, exc_traceback),
    )


def main():
    _setup_logging()
    sys.excepthook = _handle_exception
    logging.info("Application starting")

    from PySide6.QtCore import QEventLoop, QThread, QTimer, Qt, Signal
    from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap
    from PySide6.QtWidgets import (
        QApplication,
        QDialog,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMessageBox,
        QPushButton,
        QSplashScreen,
        QVBoxLayout,
    )

    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    if config.APP_ICON.exists():
        app.setWindowIcon(QIcon(str(config.APP_ICON)))

    splash_pixmap = QPixmap(620, 360)
    splash_pixmap.fill(QColor("#102b3d"))
    painter = QPainter(splash_pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.fillRect(0, 0, 620, 8, QColor("#49d7d0"))
    painter.fillRect(0, 8, 10, 352, QColor("#ffca70"))
    icon_pixmap = QPixmap(str(config.APP_ICON))
    if not icon_pixmap.isNull():
        icon_pixmap = icon_pixmap.scaled(112, 112, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        painter.drawPixmap(48, 62, icon_pixmap)
    painter.setPen(QColor("#ffffff"))
    painter.setFont(QFont("Microsoft YaHei", 25, QFont.Bold))
    painter.drawText(190, 112, config.APP_NAME)
    painter.setPen(QColor("#b8d0da"))
    painter.setFont(QFont("Microsoft YaHei", 12))
    painter.drawText(193, 148, "清单 · 定额 · 人材机 · 信息价 · AI辅助")
    painter.setPen(QColor("#6fabb0"))
    painter.drawLine(193, 180, 565, 180)
    painter.setPen(QColor("#d9fbf7"))
    painter.setFont(QFont("Microsoft YaHei", 11))
    painter.drawText(48, 292, "正在建立本地工程造价工作区")
    painter.setPen(QColor("#8ca9b5"))
    painter.setFont(QFont("Microsoft YaHei", 10))
    painter.drawText(48, 322, f"智能工程造价辅助系统  v{config.VERSION}")
    painter.end()
    splash = QSplashScreen(splash_pixmap)
    splash.setWindowTitle(f"{config.APP_NAME} - 正在启动")
    splash.showMessage(
        "正在加载数据库与工作区...",
        Qt.AlignLeft | Qt.AlignBottom,
        QColor("#475569"),
    )
    splash.show()
    app.processEvents()

    from src.db import init_db
    from src.ui.style import GLOBAL_QSS
    from src.ui.main_window import MainWindow
    from src.ui.home_page import HomePage
    from src.ui.ai_chat import AIChatPage
    from src.ui.price_info import PriceInfoPage
    from src.ui.project_workspace import ProjectWorkspacePage
    from src.ui.project_info import ProjectInfoPage
    from src.ui.base_database import BaseDatabasePage
    from src.ui.data_records import DataRecordsPage
    from src.ui.system_settings import SystemSettingsPage
    from src.ui.quota_library import QuotaLibraryPage
    from src.ui.project_list_data import ProjectListDataPage
    from src.ui.resource_library import ResourceLibraryPage
    from src.ui.user_guide import show_first_run_guide

    class _DatabaseInitWorker(QThread):
        completed = Signal(object)

        def run(self):
            try:
                init_db()
                self.completed.emit({"ok": True, "error": None})
            except Exception as error:
                logging.critical("Database initialization failed\n%s", traceback.format_exc())
                self.completed.emit({"ok": False, "error": error})

    # Keep Qt's event loop alive while the local database is migrated or
    # seeded. The old synchronous call made Windows label the application as
    # "Not Responding" before the main window was even created.
    init_worker = _DatabaseInitWorker()
    init_loop = QEventLoop()
    init_result = {"ok": False, "error": RuntimeError("数据库初始化未完成")}

    def _finish_database_init(result):
        init_result.update(result or {})
        splash.showMessage(
            "数据库已就绪，正在打开工作区..." if init_result.get("ok") else "数据库初始化失败",
            Qt.AlignLeft | Qt.AlignBottom,
            QColor("#475569"),
        )
        init_loop.quit()

    init_worker.completed.connect(_finish_database_init)
    init_worker.start()
    init_loop.exec()
    init_worker.wait()
    if not init_result.get("ok"):
        splash.close()
        _show_startup_error(init_result.get("error") or RuntimeError("数据库初始化失败"))
        return
    app.setStyleSheet(GLOBAL_QSS)

    from src.license_service import (
        activate_license,
        get_active_license_state,
        get_machine_code,
    )

    class LicenseActivationDialog(QDialog):
        def __init__(self, machine_code: str):
            super().__init__()
            self._machine_code = machine_code
            self.setWindowTitle("软件授权")
            self.setMinimumWidth(540)
            self.setModal(True)

            layout = QVBoxLayout(self)
            layout.setContentsMargins(24, 22, 24, 22)
            layout.setSpacing(14)

            title = QLabel("软件授权")
            title.setStyleSheet("font-size:20px; font-weight:600; color:#0f172a;")
            layout.addWidget(title)

            intro = QLabel("请将本机机器码发送给授权管理员，取得激活码后粘贴到下方。")
            intro.setWordWrap(True)
            intro.setStyleSheet("color:#475569; font-size:13px;")
            layout.addWidget(intro)

            contact = QLabel("获取授权码加V：13362385093，可以帮助安装 Codex，一起开发进步。")
            contact.setWordWrap(True)
            contact.setTextInteractionFlags(Qt.TextSelectableByMouse)
            contact.setStyleSheet(
                "color:#dc2626; font-size:18px; font-weight:700;"
                "background:#fef2f2; border:1px solid #fecaca;"
                "border-radius:8px; padding:10px 12px;"
            )
            layout.addWidget(contact)

            machine_label = QLabel("本机机器码")
            machine_label.setStyleSheet("font-size:13px; font-weight:600; color:#334155;")
            layout.addWidget(machine_label)

            self._machine_edit = QLineEdit(machine_code)
            self._machine_edit.setReadOnly(True)
            self._machine_edit.setStyleSheet(
                "border:1px solid #cbd5e1; border-radius:8px; padding:8px 12px;"
                "background:#f8fafc; color:#0f172a;"
                "font-family:Consolas; font-size:15px;"
            )
            machine_row = QHBoxLayout()
            machine_row.addWidget(self._machine_edit, 1)
            copy_btn = QPushButton("复制机器码")
            copy_btn.setCursor(Qt.PointingHandCursor)
            copy_btn.clicked.connect(self._copy_machine_code)
            machine_row.addWidget(copy_btn)
            layout.addLayout(machine_row)

            activation_label = QLabel("激活码")
            activation_label.setStyleSheet("font-size:13px; font-weight:600; color:#334155;")
            layout.addWidget(activation_label)

            self._activation_input = QLineEdit()
            self._activation_input.setPlaceholderText("粘贴由授权管理工具生成的激活码")
            self._activation_input.setStyleSheet(
                "border:1px solid #cbd5e1; border-radius:8px; padding:8px 12px;"
                "font-family:Consolas; font-size:13px;"
            )
            self._activation_input.returnPressed.connect(self._activate)
            layout.addWidget(self._activation_input)

            self._result_label = QLabel("")
            self._result_label.setWordWrap(True)
            self._result_label.setStyleSheet("color:#475569; font-size:13px;")
            layout.addWidget(self._result_label)

            button_row = QHBoxLayout()
            button_row.addStretch()
            exit_btn = QPushButton("退出")
            exit_btn.clicked.connect(self.reject)
            activate_btn = QPushButton("激活")
            activate_btn.setDefault(True)
            activate_btn.setStyleSheet(
                "background:#0f766e; color:#ffffff; border-radius:8px; padding:8px 24px;"
            )
            activate_btn.clicked.connect(self._activate)
            button_row.addWidget(exit_btn)
            button_row.addWidget(activate_btn)
            layout.addLayout(button_row)

        def _copy_machine_code(self):
            QApplication.clipboard().setText(self._machine_code)
            self._result_label.setText("机器码已复制，可以直接粘贴发送给授权管理员。")
            self._result_label.setStyleSheet("color:#15803d; font-size:13px;")

        def _activate(self):
            code = self._activation_input.text().strip()
            if not code:
                self._show_error("请输入激活码")
                return
            result = activate_license(code)
            if result.get("ok"):
                QMessageBox.information(self, "授权成功", result.get("message", "授权成功"))
                self.accept()
                return
            self._show_error(result.get("message", "激活失败，请检查激活码"))

        def _show_error(self, message: str):
            self._result_label.setText(message)
            self._result_label.setStyleSheet("color:#dc2626; font-size:13px;")

    def _ensure_license_enforced() -> bool:
        while True:
            state = get_active_license_state()
            if state.get("valid"):
                return True
            dialog = LicenseActivationDialog(get_machine_code())
            if dialog.exec() != QDialog.Accepted:
                return False

    if not _ensure_license_enforced():
        splash.close()
        sys.exit(0)

    window = MainWindow()

    def _lazy_page(factory):
        """Cache one page instance while keeping construction off startup."""
        cached = {}

        def create():
            if "page" not in cached:
                page = factory()
                navigate = getattr(page, "navigate_requested", None)
                if navigate is not None:
                    navigate.connect(window.show_page)
                cached["page"] = page
            return cached["page"]

        return create

    window.register_page_factory("home", _lazy_page(HomePage))
    window.register_page_factory("ai_chat", _lazy_page(AIChatPage))

    price_info_factory = _lazy_page(PriceInfoPage)
    for key in ["subscription", "source_config", "source_docs"]:
        window.register_page_factory(key, price_info_factory)

    project_workspace_factory = _lazy_page(ProjectWorkspacePage)
    for key in ["project_quote", "price_linkage", "project_snapshot", "quota_match"]:
        window.register_page_factory(key, project_workspace_factory)

    window.register_page_factory("project_info", _lazy_page(ProjectInfoPage))

    base_database_factory = _lazy_page(BaseDatabasePage)
    window.register_page_factory("price_library", base_database_factory)

    resource_library_factory = _lazy_page(ResourceLibraryPage)
    for key in ["bq_library", "market_library", "material_library", "quota_library", "project_list_data"]:
        window.register_page_factory(key, resource_library_factory)

    # Keep the old standalone routes registered as compatibility aliases for
    # callers that open a page by key instead of using the sidebar.
    for key in []:
        window.register_page_factory(key, base_database_factory)

    data_records_factory = _lazy_page(DataRecordsPage)
    for key in ["history", "audit_logs", "backup_restore", "audit_export"]:
        window.register_page_factory(key, data_records_factory)

    system_settings_factory = _lazy_page(SystemSettingsPage)
    for key in ["user_guide", "deepseek_config", "sub_rules", "anomaly_threshold", "license_management", "db_management"]:
        window.register_page_factory(key, system_settings_factory)

    window.show_page("home")
    window.show()
    splash.finish(window)
    QTimer.singleShot(450, lambda: show_first_run_guide(window))
    exit_code = app.exec()
    logging.info("Application exited with code %s", exit_code)
    sys.exit(exit_code)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        _setup_logging()
        logging.critical("Startup failed\n%s", traceback.format_exc())
        _show_startup_error(error)
        raise
