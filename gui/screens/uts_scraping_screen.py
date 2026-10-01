from pathlib import Path
from threading import Event, Thread
from urllib.parse import urlparse

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.scraper import Scraper
from core.uts_customer_runner import run_uts_workflow, validate_uts_workbook


class UTSScrapingScreen(QWidget):
    log_message = Signal(str)
    browser_started = Signal()
    browser_failed = Signal(str)
    workflow_finished = Signal()
    browser_closed = Signal()

    def __init__(self, driver_path):
        super().__init__()
        self.scraper = Scraper(driver_path)
        self.stop_event = Event()
        self._running = False
        self._browser_starting = False
        self._build_ui()

        self.log_message.connect(self.log_output.append)
        self.browser_started.connect(self._on_browser_started)
        self.browser_failed.connect(self._on_browser_failed)
        self.workflow_finished.connect(self._on_workflow_finished)
        self.browser_closed.connect(self._on_browser_closed)

    def _build_ui(self):
        heading = QLabel("UTS Scraping")
        heading.setStyleSheet("font-size: 22px; font-weight: 600;")
        instruction = QLabel(
            "Open UTS in the WebDriver browser, sign in, and navigate to the selected "
            "customer or account maker/checker screen. Customer workflows read CIF/status/"
            "branch/officer from B:E and write result/start/end to F:H (maker starts at "
            "row 4; checker at row 5). Account workflows read account/status/branch from "
            "B:D and write result/start/end to E:G (row 5 onward). Completed rows are "
            "skipped. Checker approval/undo is reported for manual action."
        )
        instruction.setWordWrap(True)

        self.url_entry = QLineEdit()
        self.url_entry.setPlaceholderText("https://...")
        self.file_entry = QLineEdit()
        self.file_entry.setReadOnly(True)
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self._browse_workbook)
        self.validate_button = QPushButton("Validate")
        self.validate_button.clicked.connect(self._validate_selected_workbook)
        file_row = QHBoxLayout()
        file_row.addWidget(self.file_entry, 1)
        file_row.addWidget(self.browse_button)
        file_row.addWidget(self.validate_button)

        self.mode_combo = QComboBox()
        self.mode_combo.addItem("Customer Maker", "maker")
        self.mode_combo.addItem("Customer Checker", "checker")
        self.mode_combo.addItem("Account Maker", "account_maker")
        self.mode_combo.addItem("Account Checker", "account_checker")
        self.mode_combo.currentIndexChanged.connect(self._validate_selected_workbook)

        form = QFormLayout()
        form.addRow("UTS website URL", self.url_entry)
        form.addRow("Input Excel file", file_row)
        form.addRow("Workflow", self.mode_combo)

        self.open_browser_button = QPushButton("Open Browser")
        self.start_button = QPushButton("Start Scraping")
        self.stop_button = QPushButton("Stop After Current Row")
        self.close_browser_button = QPushButton("Close Browser")
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.close_browser_button.setEnabled(False)
        self.open_browser_button.clicked.connect(self._open_browser)
        self.start_button.clicked.connect(self._start_workflow)
        self.stop_button.clicked.connect(self.stop_event.set)
        self.close_browser_button.clicked.connect(self._close_browser)

        actions = QHBoxLayout()
        actions.addWidget(self.open_browser_button)
        actions.addWidget(self.start_button)
        actions.addWidget(self.stop_button)
        actions.addWidget(self.close_browser_button)

        self.status_label = QLabel("Browser not started")
        self.validation_label = QLabel("Select a workbook to validate its layout.")
        self.validation_label.setWordWrap(True)
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText("Workflow progress will appear here.")

        layout = QVBoxLayout(self)
        layout.addWidget(heading)
        layout.addWidget(instruction)
        layout.addLayout(form)
        layout.addLayout(actions)
        layout.addWidget(self.validation_label)
        layout.addWidget(self.status_label)
        layout.addWidget(self.log_output, 1)

    def _browse_workbook(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select UTS input workbook",
            "",
            "Excel workbooks (*.xlsx *.xlsm *xls)",
        )
        if file_path:
            self.file_entry.setText(file_path)
            self._validate_selected_workbook()

    def _validate_selected_workbook(self):
        workbook_path = self.file_entry.text().strip()
        if not workbook_path:
            self.validation_label.setText("Select a workbook to validate its layout.")
            self.validation_label.setStyleSheet("")
            return False
        is_valid, message = validate_uts_workbook(
            workbook_path, self.mode_combo.currentData()
        )
        self.validation_label.setText(message)
        color = "#26734d" if is_valid else "#a52a2a"
        self.validation_label.setStyleSheet(f"color: {color};")
        return is_valid

    def _open_browser(self):
        if self._browser_starting or self.scraper.driver:
            return
        url = self.url_entry.text().strip()
        if not url:
            QMessageBox.warning(self, "Missing URL", "Enter the UTS website URL first.")
            return
        parsed_url = urlparse(url if "://" in url else f"https://{url}")
        if parsed_url.scheme not in {"http", "https"} or not parsed_url.netloc:
            QMessageBox.warning(self, "Invalid URL", "Enter a valid website URL.")
            return
        if "://" not in url:
            url = f"https://{url}"

        self._browser_starting = True
        self.open_browser_button.setEnabled(False)
        self.status_label.setText("Starting WebDriver browser...")
        Thread(target=self._launch_browser, args=(url,), daemon=True).start()

    def _launch_browser(self, url):
        try:
            self.scraper.start_browser(url)
            self.browser_started.emit()
        except Exception as error:
            self.browser_failed.emit(str(error))

    def _on_browser_started(self):
        self._browser_starting = False
        self.status_label.setText(
            "Browser ready. Sign in and open the required customer or account section in UTS."
        )
        self.start_button.setEnabled(True)
        self.close_browser_button.setEnabled(True)
        self.log_message.emit(
            "WebDriver browser opened. Navigate to the required UTS customer or account screen."
        )

    def _on_browser_failed(self, message):
        self._browser_starting = False
        self.open_browser_button.setEnabled(True)
        self.status_label.setText("Browser could not be started")
        QMessageBox.critical(self, "Browser Error", message)

    def _start_workflow(self):
        workbook_path = Path(self.file_entry.text())
        if not self._validate_selected_workbook():
            QMessageBox.warning(
                self,
                "Invalid Workbook",
                self.validation_label.text(),
            )
            return
        if not self.scraper.driver:
            QMessageBox.warning(self, "Browser Not Ready", "Open the WebDriver browser first.")
            return
        if self._running:
            return

        self._running = True
        self.stop_event.clear()
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.open_browser_button.setEnabled(False)
        self.close_browser_button.setEnabled(False)
        mode = self.mode_combo.currentData()
        self.status_label.setText(f"Running {self.mode_combo.currentText().lower()} workflow...")
        Thread(
            target=self._run_workflow,
            args=(str(workbook_path), mode),
            daemon=True,
        ).start()

    def _run_workflow(self, workbook_path, mode):
        try:
            run_uts_workflow(
                self.scraper.driver,
                workbook_path,
                mode,
                log_callback=self.log_message.emit,
                stop_event=self.stop_event,
            )
        except Exception as error:
            self.log_message.emit(f"Workflow stopped: {error}")
        finally:
            self.workflow_finished.emit()

    def _on_workflow_finished(self):
        self._running = False
        self.status_label.setText("Workflow finished. Check the log and workbook results.")
        self.start_button.setEnabled(bool(self.scraper.driver))
        self.stop_button.setEnabled(False)
        self.close_browser_button.setEnabled(bool(self.scraper.driver))
        self.open_browser_button.setEnabled(not bool(self.scraper.driver))

    def _close_browser(self):
        if self._running or not self.scraper.driver:
            return
        self.close_browser_button.setEnabled(False)
        Thread(target=self._quit_browser, daemon=True).start()

    def _quit_browser(self):
        self.scraper.close_browser()
        self.browser_closed.emit()

    def _on_browser_closed(self):
        self.status_label.setText("Browser closed")
        self.start_button.setEnabled(False)
        self.close_browser_button.setEnabled(False)
        self.open_browser_button.setEnabled(True)

    def closeEvent(self, event):
        if self._running:
            QMessageBox.information(
                self,
                "Workflow Running",
                "Stop the workflow and wait for it to finish before closing the application.",
            )
            event.ignore()
            return
        if self.scraper.driver:
            self.scraper.close_browser()
        event.accept()
