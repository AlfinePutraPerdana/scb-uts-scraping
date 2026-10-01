from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QHBoxLayout,
    QMessageBox, QDialog, QLabel, QLineEdit, QFormLayout, QTextEdit, QSplitter, QComboBox, QDialogButtonBox
)
from PySide6.QtGui import QFont, QColor
from PySide6.QtCore import Qt
from core.workflow import Workflow, WorkflowStep
from gui.step_editor import StepEditorDialog
import tempfile
import os
import json





class WorkflowManager(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Workflow Manager")
        self.all_workflows = []
        self.selected_index = -1
        self.workflow = None

        font = QFont()
        font.setPointSize(11)
        self.setFont(font)

        self.setStyleSheet("""
            QPushButton {
                min-height: 32px;
                min-width: 120px;
                font-size: 11pt;
            }
            QSplitter::handle {
                background-color: #666;
                border: none;
            }
            QSplitter::handle:hover {
                background-color: #aaa;
            }
        """)

        # --- Workflow Section ---
        workflow_title = QLabel("📂 Workflows")
        workflow_title.setFont(QFont("", 12, QFont.Bold))

        from PySide6.QtWidgets import QListWidget
        self.workflow_list = QListWidget()
        self.workflow_list.setToolTip("List of all workflows. Select one to view or edit its steps.")

        self.add_workflow_btn = QPushButton("Add Workflow")
        self.add_workflow_btn.setToolTip("Create a new workflow from presets")

        self.delete_workflow_btn = QPushButton("Remove Workflow")
        self.delete_workflow_btn.setToolTip("Delete the selected workflow")

        workflow_layout = QVBoxLayout()
        workflow_layout.addWidget(workflow_title)
        workflow_layout.addWidget(self.workflow_list)
        workflow_layout.addWidget(self.add_workflow_btn)
        workflow_layout.addWidget(self.delete_workflow_btn)
        workflow_container = QWidget()
        workflow_container.setLayout(workflow_layout)

        # --- Steps Section ---
        step_title = QLabel("📝 Workflow Steps")
        step_title.setFont(QFont("", 12, QFont.Bold))

        # 📊 Table for step list
        self.step_table = QTableWidget()
        self.step_table.setColumnCount(4)
        self.step_table.setHorizontalHeaderLabels(["#", "Action", "Selector", "Parameters"])
        self.step_table.horizontalHeader().setStretchLastSection(True)
        self.step_table.setToolTip("List of steps for the selected workflow")

        # Step buttons
        step_btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Step")
        self.add_btn.setToolTip("Add a new step to the workflow")
        self.edit_btn = QPushButton("Edit Step")
        self.edit_btn.setToolTip("Modify the selected step")
        self.delete_btn = QPushButton("Remove Step")
        self.delete_btn.setToolTip("Delete the selected step")
        step_btn_layout.addWidget(self.add_btn)
        step_btn_layout.addWidget(self.edit_btn)
        step_btn_layout.addWidget(self.delete_btn)

        # Move buttons
        move_btn_layout = QHBoxLayout()
        self.move_up_btn = QPushButton("Move Up")
        self.move_up_btn.setToolTip("Move selected step up")
        self.move_down_btn = QPushButton("Move Down")
        self.move_down_btn.setToolTip("Move selected step down")
        move_btn_layout.addWidget(self.move_up_btn)
        move_btn_layout.addWidget(self.move_down_btn)

        # Details area
        detail_title = QLabel("🔍 Step Details")
        detail_title.setFont(QFont("", 12, QFont.Bold))
        self.detail_action = QLineEdit()
        self.detail_selector = QLineEdit()
        self.detail_params = QTextEdit()
        self.detail_params.setFont(QFont("Consolas", 10))
        detail_layout = QFormLayout()
        detail_layout.addRow("Action:", self.detail_action)
        detail_layout.addRow("Selector:", self.detail_selector)
        detail_layout.addRow("Parameters (JSON):", self.detail_params)
        detail_container = QWidget()
        vbox = QVBoxLayout()
        vbox.addWidget(detail_title)
        vbox.addLayout(detail_layout)
        detail_container.setLayout(vbox)

        # Combine step table + controls
        step_splitter = QSplitter(Qt.Vertical)
        top_widget = QWidget()
        top_layout = QVBoxLayout()
        top_layout.addWidget(step_title)
        top_layout.addWidget(self.step_table)
        top_layout.addLayout(step_btn_layout)
        top_layout.addLayout(move_btn_layout)
        top_widget.setLayout(top_layout)

        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout()
        bottom_layout.addWidget(detail_container)
        bottom_widget.setLayout(bottom_layout)

        step_splitter.addWidget(top_widget)
        step_splitter.addWidget(bottom_widget)

        steps_container = QWidget()
        steps_layout = QVBoxLayout()
        steps_layout.addWidget(step_splitter)
        steps_container.setLayout(steps_layout)

        # --- Main Layout ---
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(workflow_container)
        splitter.addWidget(steps_container)

        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)
        self.save_btn = QPushButton("Save Workflows")
        self.save_btn.setToolTip("Save all workflows to workflow.json")
        main_layout.addWidget(self.save_btn)
        self.setLayout(main_layout)

        # --- Connections ---
        self.add_workflow_btn.clicked.connect(self.add_workflow)
        self.delete_workflow_btn.clicked.connect(self.delete_workflow)
        self.workflow_list.currentRowChanged.connect(self.on_workflow_selected)
        self.add_btn.clicked.connect(self.add_step)
        self.edit_btn.clicked.connect(self.edit_step)
        self.delete_btn.clicked.connect(self.delete_step)
        self.save_btn.clicked.connect(self.save_workflow)
        self.move_up_btn.clicked.connect(self.move_step_up)
        self.move_down_btn.clicked.connect(self.move_step_down)
        self.step_table.cellClicked.connect(self.show_step_details)
        # allow inline editing of params column and handle changes
        self._suppress_table_change = False
        self.step_table.itemChanged.connect(self.on_table_item_changed)

        self.load_workflow()
    # --- Existing methods (unchanged except for detail refresh) ---

    def load_workflow(self):
        try:
            self.all_workflows = Workflow.load_all_workflows("workflow.json")
            self.workflow_list.clear()
            for wf in self.all_workflows:
                self.workflow_list.addItem(f"{wf.name} ({wf.team})")
        except Exception as e:
            QMessageBox.warning(self, "Load Error", f"Could not load workflows: {e}")

    def add_workflow(self):
        try:
            with open("presets.json", "r", encoding="utf-8") as f:
                presets = json.load(f)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load presets.json: {e}")
            return

        dialog = AddWorkflowDialog(presets, self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()

            # Define default start and end steps for each supported screen
            DEFAULT_STEPS = {
                "Customer Information": {
                    "start": [
                        {"action": "switch_to_frame", "selector_key": "EUC Team/Customer Information/Input/iframe", "params": {}},
                        {"action": "input", "selector_key": "EUC Team/Customer Information/Input/Rel ID", "params": {"context_key": "RELNO"}},
                        {"action": "click", "selector_key": "EUC Team/Customer Information/Input/Submit_btn", "params": {}},
                        {"action": "wait_for", "selector_key": "EUC Team/Customer Information/details/details_button", "params": {}}
                    ],
                    "end": [
                        {"action": "click", "selector_key": "EUC Team/Customer Information/Input/cust_informationbtn", "params": {}}
                    ]
                },
                "ATM Inquiry Screen": {
                    "start": [
                        {"action": "switch_to_frame", "selector_key": "ATM-Inquiry-SCB/ATM Inquiry Screen/input/iframe", "params": {}},
                        {"action": "input", "selector_key": "ATM-Inquiry-SCB/ATM Inquiry Screen/input/currency", "params": {"context_key": "CCY"}},
                        {"action": "input", "selector_key": "ATM-Inquiry-SCB/ATM Inquiry Screen/input/account_number", "params": {"context_key": "Account No"}},
                        {"action": "click", "selector_key": "ATM-Inquiry-SCB/ATM Inquiry Screen/input/submit", "params": {}}
                    ],
                    "end": [
                        {"action": "switch_to_frame", "selector_key": "ATM-Inquiry-SCB/ATM Inquiry Screen/input/menuframe", "params": {}},
                        {"action": "click", "selector_key": "ATM-Inquiry-SCB/ATM Inquiry Screen/input/btn_back", "params": {}}
                    ]
                },
                "Master Information": {
                    "start": [
                        {"action": "switch_to_frame", "selector_key": "ARM Team/Master Information/input/iframe", "params": {}},
                        {"action": "input", "selector_key": "ARM Team/Master Information/input/master_number", "params": {"context_key": "MasterNo"}},
                        {"action": "click", "selector_key": "ARM Team/Master Information/input/submit_button", "params": {}},
                        {"action": "switch_to_frame", "selector_key": "ARM Team/Master Information/input/iframe", "params": {}}
                    ],
                    "end": [
                        {"action": "switch_to_frame", "selector_key": "ARM Team/Master Information/input/menuframe", "params": {}},
                        {"action": "click", "selector_key": "ARM Team/Master Information/AccountMasterTab5/btn_back_go", "params": {}}
                    ]
                },
                "Relationship Inquiry": {
                    "start": [
                        {"action": "switch_to_frame", "selector_key": "Relationship Single Team/Relationship Inquiry/Input/iframe", "params": {}},
                        {"action": "input", "selector_key": "Relationship Single Team/Relationship Inquiry/Input/rel_id", "params": {"context_key": "RELNO"}},
                        {"action": "click", "selector_key": "Relationship Single Team/Relationship Inquiry/Input/Submit_btn", "params": {}},
                        {"action": "wait_for", "selector_key": "Relationship Single Team/Relationship Inquiry/relationship_inquiry_page/rel_single_table", "params": {}}
                    ],
                    "end": [
                        {"action": "switch_to_frame", "selector_key": "Relationship Single Team/Relationship Inquiry/relationship_inquiry_page/header_frame", "params": {}},
                        {"action": "click", "selector_key": "Relationship Single Team/Relationship Inquiry/relationship_inquiry_page/Btn_Back", "params": {}}
                    ]
                },
                "general_information_inquiry_DetailsTeam": {
                    "start": [
                        {"action": "switch_to_frame", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/input/iframe", "params": {}},
                        {"action": "input", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/input/currency", "params": {"context_key": "CCY"}},
                        {"action": "input", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/input/account_number", "params": {"context_key": "Account No"}},
                        {"action": "click", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/input/submit", "params": {}},
                        {"action": "wait_for", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/general_information_page/relationship_no", "params": {}}
                    ],
                    "end": [
                        {"action": "switch_to_frame", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/input/iframe_menu", "params": {}},
                        {"action": "click", "selector_key": "General Information - Details Team/general_information_inquiry_DetailsTeam/input/btn_back_go", "params": {}}
                    ]
                }
            }

            defaults = DEFAULT_STEPS.get(data["screen"], {"start": [], "end": []})
            # create steps and tag phase explicitly so we can reliably render start vs end
            start_steps = []
            for s in defaults.get("start", []):
                st = WorkflowStep(auto_generated=True, **s)
                st.auto_phase = "start"
                start_steps.append(st)
            end_steps = []
            for s in defaults.get("end", []):
                st = WorkflowStep(auto_generated=True, **s)
                st.auto_phase = "end"
                end_steps.append(st)

            new_workflow = Workflow(
                name=data["name"] or f"Workflow {len(self.all_workflows) + 1}",
                team=data["team"],
                screen=data["screen"],
                tab=data["tab"],
                steps=start_steps + end_steps
            )

            self.all_workflows.append(new_workflow)
            self.workflow_list.addItem(f"{new_workflow.name} ({new_workflow.team})")
            self.workflow_list.setCurrentRow(len(self.all_workflows) - 1)

            # ✅ Refresh the step table immediately
            self.workflow = new_workflow
            self._refresh_step_table()

            # ✅ Visually mark auto-generated steps (gray italic)
            for i, step in enumerate(new_workflow.steps):
                if getattr(step, "auto_generated", False):
                    for col in range(self.step_table.columnCount()):
                        item = self.step_table.item(i, col)
                        if item:
                            font = item.font()
                            font.setItalic(True)
                            item.setFont(font)
                            item.setForeground(Qt.gray)


    
    def delete_workflow(self):
        row = self.workflow_list.currentRow()
        if row < 0:
            return
        wf = self.all_workflows[row]
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete workflow '{wf.name}'?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            del self.all_workflows[row]
            self.workflow_list.takeItem(row)
            if self.all_workflows:
                self.selected_index = 0
                self.workflow = self.all_workflows[0]
                self._refresh_step_table()
            else:
                self.selected_index = -1
                self.workflow = None
                self.step_table.clear()



    def on_workflow_selected(self, index):
        if index < 0 or index >= len(self.all_workflows):
            self.workflow = None
            self.step_table.setRowCount(0)
            return
        self.selected_index = index
        self.workflow = self.all_workflows[index]
        self._refresh_step_table()

    def move_step_up(self):
        row = self.step_table.currentRow()
        if row > 0:
            self.workflow.steps[row - 1], self.workflow.steps[row] = self.workflow.steps[row], self.workflow.steps[row - 1]
            self._refresh_step_table()
            self.step_table.selectRow(row - 1)

    def move_step_down(self):
        row = self.step_table.currentRow()
        if row < len(self.workflow.steps) - 1:
            self.workflow.steps[row + 1], self.workflow.steps[row] = self.workflow.steps[row], self.workflow.steps[row + 1]
            self._refresh_step_table()
            self.step_table.selectRow(row + 1)

    def add_step(self):
        dialog = StepEditorDialog(self, workflow_screen=self.workflow.screen)
        if dialog.exec() == QDialog.Accepted:
            step = dialog.get_step()
            self.workflow.steps.append(step)
            self._refresh_step_table()

    def edit_step(self):
        row = self.step_table.currentRow()
        if row < 0:
            return
        old_step = self.workflow.steps[row]
        dialog = StepEditorDialog(self, old_step, workflow_screen=self.workflow.screen)
        if dialog.exec() == QDialog.Accepted:
            self.workflow.steps[row] = dialog.get_step()
            self._refresh_step_table()


    def delete_step(self):
        current = self.step_table.currentRow()
        if current >= 0:
            del self.workflow.steps[current]
            self.step_table.takeItem(current)
            self.show_step_details(-1)
    
    def _refresh_step_table(self):
        self.step_table.setRowCount(0)
        if not self.workflow:
            return
        # prevent itemChanged handler firing while we populate the table
        self._suppress_table_change = True

        for i, step in enumerate(self.workflow.steps):
            self.step_table.insertRow(i)

            # --- Simplify selector display (only show Tab/Field)
            short_selector = step.selector_key
            if step.selector_key:
                parts = step.selector_key.split("/")
                if len(parts) >= 2:
                    short_selector = "/".join(parts[-2:])  # last two parts: tab/field

            # --- Detect auto-generated start/end
            is_auto = getattr(step, "auto_generated", False)
            phase = getattr(step, "auto_phase", None)
            is_start = is_auto and phase == "start"
            is_end = is_auto and phase == "end"

            # --- Populate cells
            index_item = QTableWidgetItem(str(i + 1))
            action_item = QTableWidgetItem(step.action)
            selector_item = QTableWidgetItem(short_selector)
            selector_item.setToolTip(step.selector_key)
            params_item = QTableWidgetItem(json.dumps(step.params))
            # Make params cell editable even for auto-generated/default steps
            params_item.setFlags(params_item.flags() | Qt.ItemIsEditable)

            # --- Apply special styling
            if is_auto:
                font = QFont()
                font.setItalic(True)
                action_item.setFont(font)
                selector_item.setFont(font)
                params_item.setFont(font)

                # Background color per type
                if is_start:
                    bg_color = QColor("#eaffea")  # light green
                    action_item.setText(f" {step.action}")
                    tip = "Auto-generated Start Step"
                elif is_end:
                    bg_color = QColor("#ffecec")  # light red
                    action_item.setText(f" {step.action}")
                    tip = "Auto-generated End Step"
                else:
                    bg_color = QColor("#f5f5f5")
                    tip = "Auto-generated Step"

                for col, item in enumerate([index_item, action_item, selector_item, params_item]):
                    item.setBackground(bg_color)
                    item.setToolTip(tip)
                    item.setForeground(QColor("#555"))

            # Insert items
            self.step_table.setItem(i, 0, index_item)
            self.step_table.setItem(i, 1, action_item)
            self.step_table.setItem(i, 2, selector_item)
            self.step_table.setItem(i, 3, params_item)

        # Optional: auto-resize for clarity
        self.step_table.resizeColumnsToContents()
        self.step_table.horizontalHeader().setStretchLastSection(True)
        # re-enable itemChanged processing after table populated
        self._suppress_table_change = False

    # The following method is replaced by more efficient update in edit_step and add_step
    # def _refresh_step_list(self):
    #     self.step_table.clear()
    #     for i, step in enumerate(self.workflow.steps):
    #         self.step_table.addItem(f"{i+1}. {step.action} → {step.selector_key}")
    #     self.show_step_details(self.step_table.currentRow())
    
    def update_step_details(self):
        row = self.step_table.currentRow()
        if row < 0 or row >= len(self.workflow.steps):
            return

        step = self.workflow.steps[row]
        step.action = self.detail_action.text().strip()
        step.selector_key = self.detail_selector.text().strip()

        text = self.detail_params.toPlainText().strip()
        if text:
            try:
                step.params = json.loads(text)
            except json.JSONDecodeError:
                QMessageBox.warning(self, "Invalid JSON", "Parameters must be valid JSON.")
                return
        else:
            step.params = {}

        # 🔥 langsung update tampilan tabel tanpa refresh total (biar edit gak hilang)
        self._suppress_table_change = True
        self.step_table.item(row, 1).setText(step.action)
        short_selector = "/".join(step.selector_key.split("/")[-2:])
        self.step_table.item(row, 2).setText(short_selector)
        self.step_table.item(row, 3).setText(json.dumps(step.params))
        self._suppress_table_change = False



    def show_step_details(self, row):
        if row < 0 or row >= len(self.workflow.steps):
            return

        step = self.workflow.steps[row]

        # block signals to prevent recursive updates
        self.detail_action.blockSignals(True)
        self.detail_selector.blockSignals(True)
        self.detail_params.blockSignals(True)

        self.detail_action.setText(step.action or "")
        self.detail_selector.setText(step.selector_key or "")
        self.detail_params.setPlainText(json.dumps(step.params, indent=2))

        # unblock after setting values
        self.detail_action.blockSignals(False)
        self.detail_selector.blockSignals(False)
        self.detail_params.blockSignals(False)


    def save_workflow(self):
        if 0 <= self.selected_index < len(self.all_workflows):
            self.all_workflows[self.selected_index] = self.workflow
        try:
            data = json.dumps([wf.to_dict() for wf in self.all_workflows], indent=2)
            temp_fd, temp_path = tempfile.mkstemp()
            with os.fdopen(temp_fd, 'w', encoding='utf-8') as tmp_file:
                tmp_file.write(data)
            os.replace(temp_path, "workflow.json")
            QMessageBox.information(self, "Saved", "Workflows saved successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save workflows: {e}")
    
    def on_table_item_changed(self, item):
        """Handle inline edits to the step table (params column)."""
        if getattr(self, "_suppress_table_change", False):
            return
        if not self.workflow:
            return
        row = item.row()
        col = item.column()
        if row < 0 or row >= len(self.workflow.steps):
            return
        # params column index is 3
        if col == 3:
            text = item.text().strip()
            if not text:
                self.workflow.steps[row].params = {}
                return
            try:
                parsed = json.loads(text)
            except json.JSONDecodeError:
                QMessageBox.warning(self, "Invalid JSON", "Parameters must be valid JSON. Reverting.")
                # revert text to current params
                self._suppress_table_change = True
                item.setText(json.dumps(self.workflow.steps[row].params))
                self._suppress_table_change = False
                return
            # valid JSON -> update underlying step
            self.workflow.steps[row].params = parsed

from PySide6.QtWidgets import (
    QDialog, QFormLayout, QLineEdit, QComboBox, 
    QDialogButtonBox
)

class AddWorkflowDialog(QDialog):
    def __init__(self, presets, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Workflow")

        # Presets structure: team -> screen -> tab -> fields
        self.presets = presets or {}

        # Map screen -> team (first level mapping)
        self.screen_to_team = {}
        self.tabs_by_screen = {}
        for team, screens in self.presets.items():
            for screen, tabs in screens.items():
                self.screen_to_team[screen] = team
                self.tabs_by_screen[screen] = sorted(tabs.keys())

        layout = QFormLayout(self)

        self.name_edit = QLineEdit(self)
        layout.addRow("Workflow Name:", self.name_edit)

        # Screen dropdown (no team shown)
        self.screen_combo = QComboBox()
        self.screen_combo.addItems(sorted(self.tabs_by_screen.keys()))
        layout.addRow("Screen:", self.screen_combo)

        # Tab dropdown
        self.tab_combo = QComboBox()
        layout.addRow("Tab:", self.tab_combo)

        # Populate tabs initially
        self.populate_tabs(self.screen_combo.currentText())
        self.screen_combo.currentTextChanged.connect(self.populate_tabs)

        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def populate_tabs(self, screen):
        self.tab_combo.clear()
        if screen in self.tabs_by_screen:
            self.tab_combo.addItems(self.tabs_by_screen[screen])

    def get_data(self):
        screen = self.screen_combo.currentText().strip()
        tab = self.tab_combo.currentText().strip()
        team = self.screen_to_team.get(screen, "")
        return {
            "name": self.name_edit.text().strip(),
            "team": team,   # hidden from user, but still included internally
            "screen": screen,
            "tab": tab
        }
    
    


