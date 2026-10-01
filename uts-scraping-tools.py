import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from gui.main_gui import MainGUIWithSidebar  # Import the Main GUI class
from core.utils import resource_path, bundled_resource_path  # Import the utility function to get resource paths
import os
from PySide6.QtGui import QFont

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("scb-logo.ico"))

    # ✅ Apply global font
    app.setFont(QFont("Segoe UI", 10))

    driver_path = resource_path("msedgedriver.exe")
    if not os.path.exists(driver_path):  # fallback if not found locally
        driver_path = bundled_resource_path("msedgedriver.exe")
    window = MainGUIWithSidebar(driver_path)
    window.show()
    sys.exit(app.exec())