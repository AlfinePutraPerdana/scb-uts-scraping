# gui/main_gui.py
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QStackedWidget, QPushButton, QFrame, QLabel
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt
from gui.screens.scraping_screen import ScrapingScreen
from gui.screens.uts_scraping_screen import UTSScrapingScreen
from gui.screens.comparing_screen import ComparingScreen
from gui.workflow_manager import WorkflowManager
from gui.preset_manager import PresetManager
from core.utils import resource_path,bundled_resource_path
from core import presets
from core.utils import get_edge_driver_version


class MainGUIWithSidebar(QWidget):
    def __init__(self, driver_path):
        super().__init__()
        driver_version = get_edge_driver_version(driver_path)
        self.setWindowTitle(f"UTS Scraping Tools - {driver_version}")

        # make it resizable instead of fixed
        self.resize(1200, 800)
        self.setMinimumSize(900, 600)

        self.stack = QStackedWidget()

        # Screens
        self.scraping_screen = ScrapingScreen(driver_path)
        self.uts_scraping_screen = UTSScrapingScreen(driver_path)
        self.comparing_screen = ComparingScreen()
        self.preset_manager = PresetManager()
        self.workflow_manager = WorkflowManager()

        self.stack.addWidget(self.scraping_screen)
        self.stack.addWidget(self.uts_scraping_screen)
        self.stack.addWidget(self.comparing_screen)
        self.stack.addWidget(self.preset_manager)
        self.stack.addWidget(self.workflow_manager)

        # Sidebar menu
        sidebar = QVBoxLayout()
        sidebar.setAlignment(Qt.AlignTop)
        vline = QFrame()
        vline.setFrameShape(QFrame.VLine)
        vline.setFrameShadow(QFrame.Sunken)
        vline.setLineWidth(1)

        logo_label = QLabel()
        pixmap = QPixmap(bundled_resource_path("assets/standard-chartered.png"))
        pixmap = pixmap.scaledToWidth(150)
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignCenter)
        sidebar.addWidget(logo_label)

        btn_scrape = QPushButton("🔍 Scraping")
        btn_scrape.setStyleSheet("text-align: left; padding: 5px; font-size : 15px")
        btn_scrape.clicked.connect(lambda: self.stack.setCurrentWidget(self.scraping_screen))
        sidebar.addWidget(btn_scrape)

        btn_uts_scrape = QPushButton("UTS Scraping")
        btn_uts_scrape.setStyleSheet("text-align: left; padding: 5px; font-size : 15px")
        btn_uts_scrape.clicked.connect(lambda: self.stack.setCurrentWidget(self.uts_scraping_screen))
        sidebar.addWidget(btn_uts_scrape)

        sidebar_frame = QFrame()
        sidebar_frame.setFixedWidth(200)
        sidebar_frame.setLayout(sidebar)

        layout = QHBoxLayout(self)
        layout.addWidget(sidebar_frame)
        layout.addWidget(vline)
        layout.addWidget(self.stack)

    def open_preset_manager(self):
        dlg = PresetManager(self)
        dlg.exec()
        self.scraping_screen.team_combo.clear()
        self.scraping_screen.team_combo.addItems(presets.get_team_names())
