from PySide6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QComboBox,
    QMessageBox, QHBoxLayout, QToolButton, QFormLayout, QSpacerItem, QSizePolicy, QGroupBox,
    QProgressBar, QTextEdit
)
from PySide6.QtGui import QIcon
from PySide6.QtCore import Signal, QTimer, QMetaObject, Qt
from core.scraper import Scraper
from core import presets
from core.workflow import Workflow
import openpyxl
from openpyxl.utils import get_column_letter
import threading

class ScrapingScreen(QWidget):
    enquiry_opened = Signal(str)
    enquiry_failed = Signal()

    def __init__(self, driver_path):
        super().__init__()
        self.driver_path = driver_path
        self.scraper = Scraper(driver_path)
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Enquiry Scraper")
        self.setMaximumHeight(350)  # Slight limit to avoid extra vertical space

        # === FORM INPUTS ===
        self.url_entry = QLineEdit()
        self.team_combo = QComboBox()
        all_workflows = Workflow.load_all_workflows("workflow.json")
        team_names = sorted(set(w.name for w in all_workflows))
        self.team_combo.addItems(team_names)

        
        # Interchangeable input layout for other screens
        form_layout = QFormLayout()
        form_layout.addRow("Web App URL:", self.url_entry)        # <-- Interchangeable: URL input
        form_layout.addRow("Workflow:", self.team_combo)       # <-- Interchangeable: Team selector

        # Input File browse
        self.file_path_entry = QLineEdit()
        self.file_path_entry.setReadOnly(True)
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_file)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_entry)
        file_layout.addWidget(self.browse_button)
        form_layout.addRow("Input File:", file_layout)
        
        # === ACTION BUTTONS ===
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Start")
        self.start_button.setStyleSheet("font-size : 15px")
        self.resume_button = QPushButton("Resume")
        self.resume_button.setStyleSheet("font-size : 15px")
        self.reset_button = QPushButton("Reset")
        self.reset_button.setStyleSheet("font-size : 15px")
        self.resume_button.setEnabled(False)

        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.resume_button)
        button_layout.addWidget(self.reset_button)
        
        # GroupBox for visual grouping (optional but helps with modularity)
        input_group = QGroupBox("Scraping Configuration")  # <-- Interchangeable: Rename for other screens
        input_group.setLayout(form_layout)

        # === MAIN LAYOUT ===
        main_layout = QVBoxLayout()
        main_layout.addWidget(input_group)
        main_layout.addLayout(button_layout)
        main_layout.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding))  # Pushes UI up
        self.setLayout(main_layout)

        # === BUTTON CONNECTIONS ===
        self.start_button.clicked.connect(self.start_browser)
        self.resume_button.clicked.connect(self.resume_scraping)
        self.reset_button.clicked.connect(self.reset_app)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        # Log Output
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("font-family: Consolas; font-size: 11px;")

        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.log_output)
        
        self.log_output.append("🔥 GUI is alive and can print this!")

    def browse_file(self):
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*.*)")
        if file_path:
            self.file_path_entry.setText(file_path)
    
    def start_browser(self):
        url = self.url_entry.text().strip()
        if not url:
            QMessageBox.critical(self, "Input Error", "Please enter a valid URL.")
            return
        if not url.startswith("http://") and not url.startswith("https://"):
            url = "https://" + url

        self.start_button.setEnabled(False)
        self.resume_button.setEnabled(True)

        QMessageBox.information(
            self,
            "Next Step",
            "Please log in to your web app and open the Enquiry page.\n\nClick 'Resume' once the Enquiry window is fully opened."
        )

        threading.Thread(target=self.scraper.start_browser, args=(url,), daemon=True).start()

    def resume_scraping(self):
        self.resume_button.setEnabled(False)

        # ✅ Get workflow name from combo box
        selected_name = self.team_combo.currentText()

        # Load presets and workflow
        from core import presets
        self.preset_data = presets.load_presets()

        from core.workflow import Workflow
        all_workflows = Workflow.load_all_workflows("workflow.json")

        # ✅ Match workflow by name
        matching = [w for w in all_workflows if w.name == selected_name]
        if not matching:
            QMessageBox.warning(self, "Missing Workflow", f"No workflow found with name {selected_name}.")
            self.reset_app()
            return

        workflow = matching[0]

        # ✅ Assign details from workflow
        self.team = workflow.team
        # keep workflow name as well because UI now selects by workflow
        # and some workflows may have team values that don't match the old hardcoded strings
        self.workflow_name = workflow.name
        self.screen = workflow.screen
        self.tab = workflow.tab
        self.steps = [step.to_dict() for step in workflow.steps]

        # ✅ Start the thread here
        self.start_scraping()

    
    def start_scraping(self):
        threading.Thread(target=self.scrape_thread, daemon=True).start()

    def scrape_thread(self):
        import sys
        import os
        from datetime import datetime
        from core.automation_runner import run_workflow

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_file_path = os.path.join(log_dir, f"scrape_log_{timestamp}.txt")
        print(f"\U0001F4C4 Log output will be saved to: {os.path.abspath(log_file_path)}")

        sys.stdout = open(log_file_path, "w", encoding="utf-8")
        sys.stderr = sys.stdout
        self.log("🟢 Scrape thread started")

        title = self.scraper.wait_for_new_window()
        if not title:
            self.enquiry_failed.emit()
            self.scraper.close_browser()
            self.call_in_main(self.reset_app)
            return

        self.enquiry_opened.emit(title)

        file_path = self.file_path_entry.text()
        try:
            wb = openpyxl.load_workbook(file_path)
        except Exception as e:
            self.call_in_main(lambda: QMessageBox.critical(None, "Error", f"Failed to open Excel file:\n{e}"))
            self.scraper.close_browser()
            self.call_in_main(self.reset_app)
            return

        ws = wb.active

        header = [cell.value for cell in ws[1]]
        header_map = {col: idx + 1 for idx, col in enumerate(header)}

        required_inputs = []
        for step in self.steps:
            if step.get("action") == "input":
                context_key = step.get("params", {}).get("context_key") or step.get("value_from")
                if context_key:
                    required_inputs.append(context_key)

        missing = [key for key in required_inputs if key not in header_map]
        if missing:
            self.call_in_main(lambda: QMessageBox.critical(None, "Error",
                f"Missing required input columns in Excel: {', '.join(missing)}"))
            self.scraper.close_browser()
            self.call_in_main(self.reset_app)
            return

        status_col_name = "SCRAPE_STATUS"
        if status_col_name not in header_map:
            new_col_idx = len(header_map) + 1
            ws.cell(row=1, column=new_col_idx).value = status_col_name
            header_map[status_col_name] = new_col_idx

        total_rows = max(1, ws.max_row - 1)
        processed_count = 0

        result_ws = None
        self.atm_header_written = False  # Reset setiap scrape baru
        result_sheet_name = ""
        # Support matching by team OR workflow name (UI now selects workflow name)
        identifier = (self.team or "", getattr(self, "workflow_name", ""))

        if "ATM-Inquiry-SCB" in identifier or "ATM Inquiry Workflow" in identifier:
            result_sheet_name = "ATM Result"
        elif "Relationship Single Team" in identifier:
            result_sheet_name = "Relationship Result"
        elif "test_new_close_window" in identifier:
            result_sheet_name = "testing Result"
        elif "General Information - Sweeps Team" in identifier:
            result_sheet_name = "Sweeps Result"
        elif "General Information - Risk Details" in identifier:
            result_sheet_name = "Risk Result"
        elif "General Information - Charge Details" in identifier:
            result_sheet_name = "Charge Result"
        elif "General Information - Block Details" in identifier:
            result_sheet_name = "Block Result"
        elif "General Information - InfoType Details" in identifier:
            result_sheet_name = "InfoType Result"
        elif "General Information - LINKAGES" in identifier:
            result_sheet_name = "LINKAGES Result"
        elif "General Information - Other Details" in identifier:
            result_sheet_name = "OtherDetails Result"
        elif "General Informtion - Interest Rate" in identifier:
            result_sheet_name = "Credit Result"

        if result_sheet_name:
            if result_sheet_name not in wb.sheetnames:
                result_ws = wb.create_sheet(result_sheet_name)
                
                if self.team == "Relationship Single Team":
                    result_ws.append([
                        "rel_id", "sys_id", "prod_cd_deal_type", "prod_desc_deal_type_desc", "account_cca_rla_deal_number",
                        "account_ccy", "atm_debit_card_no", "card_seq_no", "prim_sec_link_type", "operating_instructions",
                        "consol_statement_indicator", "sort_code", "payroll_indicator", "bonus_payroll", "open_date",
                        "close_or_maturity_date", "available_balance_without_limits", "ledger_balance", "limit", "lien",
                        "related_relationship_no", "status", "holder", "old_account_cca_rla_deal_number"
                    ])
            else:
                result_ws = wb[result_sheet_name]

                # ⛔ Bersihkan isi sheet untuk ATM agar header lama tidak tertinggal
                if self.team == "ATM-Inquiry-SCB" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "test_new_close_window" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - Sweeps Team" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - Risk Details" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - Charge Details" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - Block Details" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - InfoType Details" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - LINKAGES" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Information - Other Details" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)
                elif self.team == "General Informtion - Interest Rate" and result_ws.max_row > 0:
                    result_ws.delete_rows(1, result_ws.max_row)

        save_interval = 5  # Save every 5 processed rows
        rows_since_save = 0
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            row_number = row[0].row
            context = {}
            # ✅ Skip already processed rows
            status_value = row[header_map[status_col_name] - 1].value
            if status_value in ("Success", "Error"):
                self.log(f"⏩ Row {row_number}: Already processed — skipped.")
                continue
            
            missing_input = False

            for key in required_inputs:
                value = row[header_map[key] - 1].value
                if not value:
                    missing_input = True
                    break
                context[key] = value

            if missing_input:
                self.log(f"⚠️ Row {row_number}: Missing required input — skipped.")
                continue

            print("──────────────────────────────────────────")
            print(f"🔁 Mulai scraping baris {row_number}")

            try:
                def log_cb(msg, row_number=row_number):
                    self.log(f"📝 Row {row_number}: {msg}")

                context["current_excel_row"] = {}  # ✅ Fix: allow dynamic table fields to be populated
                print(f"➡️ Executing workflow for row {row_number} with context: {context}")
                results = run_workflow(
                    driver=self.scraper.driver,
                    workflow=self.steps,
                    context=context,
                    presets=self.preset_data,
                    team=self.team,
                    screen=self.screen,
                    tab=self.tab,
                    log_callback=log_cb
                )

                if self.team == "ATM-Inquiry-SCB":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")

                elif self.team == "test_new_close_window":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_atm_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")


                elif self.team == "General Information - Sweeps Team":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")
                elif self.team == "General Information - Risk Details":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")
                elif self.team == "General Information - Charge Details":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")
                        
                elif self.team == "General Information - Block Details":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")
                
                elif self.team == "General Information - InfoType Details":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")

                elif self.team == "General Information - LINKAGES":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_atm_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")
                
                elif self.team == "General Information - Other Details":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")
                
                elif self.team == "General Informtion - Interest Rate":
                    atm_rows = results.get("ATM_Table", [])
                    print(f"🔍 Row {row_number}: Found {len(atm_rows)} ATM row(s) for Account: {context.get('Account No', 'N/A')}")
                    if isinstance(atm_rows, list) and atm_rows:
                        for row_data in atm_rows:
                            self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                    else:
                        print(f"⚠️ Row {row_number}: ATM_Table is empty or invalid — skipped writing.")

                elif self.team == "Relationship Single Team":
                    relationship_rows = results.get("rel_single_table", [])
                    print(f"🔍 Row {row_number}: Found {len(relationship_rows)} relationship row(s) for RELNO: {context.get('RELNO', 'N/A')}")
                    if isinstance(relationship_rows, list) and relationship_rows:
                        for row_data in relationship_rows:
                            # row_data["rel_id"] = context.get("RELNO", "")
                            result_ws.append([
                                context.get("RELNO", ""), row_data.get("sys_id", ""), row_data.get("prod_cd_deal_type", ""),
                                row_data.get("prod_desc_deal_type_desc", ""), row_data.get("account_cca_rla_deal_number", ""),
                                row_data.get("account_ccy", ""), row_data.get("atm_debit_card_no", ""), row_data.get("card_seq_no", ""),
                                row_data.get("prim_sec_link_type", ""), row_data.get("operating_instructions", ""),
                                row_data.get("consol_statement_indicator", ""), row_data.get("sort_code", ""),
                                row_data.get("payroll_indicator", ""), row_data.get("bonus_payroll", ""), row_data.get("open_date", ""),
                                row_data.get("close_or_maturity_date", ""), row_data.get("available_balance_without_limits", ""),
                                row_data.get("ledger_balance", ""), row_data.get("limit", ""), row_data.get("lien", ""),
                                row_data.get("related_relationship_no", ""), row_data.get("status", ""), row_data.get("holder", ""),
                                row_data.get("old_account_cca_rla_deal_number", "")
                            ])
                    else:
                        print(f"⚠️ Row {row_number}: rel_single_table is empty or invalid — skipped writing.")
                else:
                    for key, value in results.items():
                        # Special-case: if the result is a table (list of dicts) and we have a result sheet,
                        # write each dict as a separate row in the result worksheet instead of stringifying it.
                        if isinstance(value, list) and value and isinstance(value[0], dict) and result_ws:
                            # If this is the ATM table, prefer the ATM-specific writer so AccountNo/CCY
                            # and headers are handled consistently with existing code.
                            if key == "ATM_Table":
                                for row_data in value:
                                    # Use the same writer used elsewhere; prefer general info writer for ATM-Inquiry flows
                                    if self.team == "ATM-Inquiry-SCB" or self.team.startswith("General Information"):
                                        self.append_general_information_result(result_ws, row_data, context, row_number, log_cb)
                                    else:
                                        self.append_atm_result2(result_ws, row_data, context, row_number, log_cb)
                                # done handling this key
                                continue

                            # Generic list-of-dicts handling for other table results
                            headers = list(value[0].keys())
                            # If sheet is empty, write headers first
                            if result_ws.max_row == 0:
                                result_ws.append(headers)
                            for rowd in value:
                                result_ws.append([rowd.get(h, "") for h in headers])
                            continue

                        # Fallback: write scalar or unknown structures into the input worksheet (as before)
                        if key not in header_map:
                            new_col_idx = len(header_map) + 1
                            ws.cell(row=1, column=new_col_idx).value = key
                            header_map[key] = new_col_idx

                        # Don't write raw dict/list into a single cell; stringify instead
                        if isinstance(value, (dict, list)):
                            value = str(value)

                        ws.cell(row=row_number, column=header_map[key], value=value)

                # 🔹 Write dynamic table values
                if "current_excel_row" in context:
                    for key, value in context["current_excel_row"].items():
                        if key not in header_map:
                            new_col_idx = len(header_map) + 1
                            ws.cell(row=1, column=new_col_idx).value = key
                            header_map[key] = new_col_idx

                        if isinstance(value, (dict, list)):
                            value = str(value)

                        ws.cell(row=row_number, column=header_map[key], value=value)

                # ✅ Pastikan kolom status ada
                if status_col_name not in header_map:
                    new_col_idx = len(header_map) + 1
                    ws.cell(row=1, column=new_col_idx).value = status_col_name
                    header_map[status_col_name] = new_col_idx

                ws.cell(row=row_number, column=header_map[status_col_name], value="Success")
                self.log(f"✅ Row {row_number}: Extracted {len(results)} fields.")

            except Exception as e:
                import traceback
                concise_msg = str(e).split("\n")[0]
                self.log(f"❌ Error scraping row {row_number}: {concise_msg}")
                ws.cell(row=row_number, column=header_map[status_col_name], value=f"Error: {concise_msg}")

            processed_count += 1
            rows_since_save += 1
            progress = int((processed_count / total_rows) * 100)
            self.call_in_main(lambda: self.progress_bar.setValue(progress))

            # ✅ Auto-save every N rows
            if rows_since_save >= save_interval:
                try:
                    saved_rows = rows_since_save  # store before resetting
                    wb.save(file_path)
                    self.log(f"💾 Auto-saved after {saved_rows} row(s)")
                except Exception as e:
                    self.log(f"⚠️ Auto-save failed: {e}")
                rows_since_save = 0

        try:
            wb.save(file_path)
        except Exception as e:
            self.call_in_main(lambda: QMessageBox.critical(None, "Error", f"Failed to save Excel file:\n{e}"))

        self.call_in_main(lambda: QMessageBox.information(None, "Done", "Scraping finished. Excel file updated."))
        self.scraper.close_browser()
        sys.stdout.close()
        self.call_in_main(self.reset_app)

    def append_dynamic_table_result(self, ws, table_data, context, row_number, log_cb):
        """
        Menambahkan hasil dynamic table ke worksheet.
        table_data: list of dict hasil extract_dynamic_table_multiple
        """
        if not table_data:
            log_cb(f"⚠️ Row {row_number}: Table data kosong — skipped writing.")
            return

        # Ambil AccountNo dan CCY dari context (default kosong kalau tidak ada)
        account_no = context.get("account_no", "")
        ccy = context.get("ccy", "")

        # Pastikan header ditulis dulu
        headers = ["AccountNo", "CCY"] + list(table_data[0].keys())
        if ws.max_row == 0:
            ws.append(headers)

        # Append setiap baris data
        for row in table_data:
            ws.append([account_no, ccy] + [row.get(col, "") for col in headers])

        log_cb(f"✅ Row {row_number}: {len(table_data)} baris berhasil ditambahkan "
        f"dengan AccountNo={account_no}, CCY={ccy}.")

    def append_atm_result(self, result_ws, row_data, context, row_number, log):
        if not row_data:
            print(f"⚠️ Row {row_number}: Data ATM kosong, tidak diproses.")
            return

        # ✅ Tulis header jika belum pernah ditulis
        if not hasattr(self, "atm_header_written") or not self.atm_header_written:
            headers = list(row_data.keys())
            if result_ws:
                result_ws.append(headers)
            self.atm_header_written = True
            print("📝 Header ATM ditambahkan:", headers)

        row_values = []

        # Tambahkan semua field dari row_data tanpa header
        for i, (field, value) in enumerate(row_data.items(), start=1):
            if isinstance(value, str):
                value = value.strip()
            row_values.append(value) # ⬅️ Tambahkan ke list
            print(f"    [{i:02}] {field}: '{value}'")
        # ✅ Simpan row_values jika result_ws tersedia
        if result_ws:
            result_ws.append(row_values) # ⬅️ Tambahkan baris ke Excel

        # ✅ Log tambahan
        print(f"🧾 Row {row_number}: Menulis ATM (AccountNo: {context.get('AccountNo', '-')}, {len(row_data)} kolom)")

    def append_atm_result2(self, result_ws, row_data, context, row_number, log):
        if not row_data:
            print(f"⚠️ Row {row_number}: Data ATM kosong, tidak diproses.")
            return
        
        # Ambil dari context
        account_no = context.get("Account No", "-")
        ccy = context.get("CCY", "-")

        combined_row_data = {
            "AccountNo": account_no,
            "CCY": ccy,
            **row_data
        }

        # ✅ Gunakan flag terpisah biar ga tabrakan sama ATM
        if not hasattr(self, "general_info_header_written") or not self.general_info_header_written:
            headers = list(combined_row_data.keys())
            if result_ws:
                result_ws.append(headers)
            self.general_info_header_written = True
            print("📝 Header ATM ditambahkan:", headers)

        row_values = []

        # Gunakan combined_row_data supaya AccountNo & CCY ikut ditulis
        for i, (field, value) in enumerate(combined_row_data.items(), start=1):
            if isinstance(value, str):
                value = value.strip()
            row_values.append(value)
            print(f"    [{i:02}] {field}: '{value}'")

        if result_ws:
            result_ws.append(row_values)

        print(f"🧾 Row {row_number}: Menulis ATM (AccountNo: {account_no}, {len(combined_row_data)} kolom)")

    def append_general_information_result(self, result_ws, row_data, context, row_number, log):
        if not row_data:
            print(f"⚠️ Row {row_number}: Data General Information kosong, tidak diproses.")
            return
        
        # Ambil dari context
        account_no = context.get("Account No", "-")
        ccy = context.get("CCY", "-")

        combined_row_data = {
            "AccountNo": account_no,
            "CCY": ccy,
            **row_data
        }

        # ✅ Gunakan flag terpisah biar ga tabrakan sama ATM
        if not hasattr(self, "general_info_header_written") or not self.general_info_header_written:
            headers = list(combined_row_data.keys())
            if result_ws:
                result_ws.append(headers)
            self.general_info_header_written = True
            print("📝 Header General Information ditambahkan:", headers)

        row_values = []

        # Gunakan combined_row_data supaya AccountNo & CCY ikut ditulis
        for i, (field, value) in enumerate(combined_row_data.items(), start=1):
            if isinstance(value, str):
                value = value.strip()
            row_values.append(value)
            print(f"    [{i:02}] {field}: '{value}'")

        if result_ws:
            result_ws.append(row_values)

        print(f"🧾 Row {row_number}: Menulis General Information (AccountNo: {account_no}, {len(combined_row_data)} kolom)")

    def log(self, message):
        def append():
            print(f"[GUI log] {message}")
            self.log_output.append(message)
            self.log_output.ensureCursorVisible()
        self.call_in_main(append)


    def call_in_main(self, func):
        QTimer.singleShot(0, func)

    
    def reset_app(self):
        self.log_output.clear()  # Clear the log output

        self.url_entry.clear()
        self.team_combo.setCurrentIndex(0)
        self.start_button.setEnabled(True)
        self.resume_button.setEnabled(False)

        if self.scraper.driver:
            self.scraper.close_browser()
        self.scraper = Scraper(self.driver_path)

        print("Application has been reset.")
