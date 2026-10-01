# step_editor_dialog.py
import json
import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QDialogButtonBox,
    QPlainTextEdit, QComboBox, QPushButton, QHBoxLayout, QMessageBox
)
from core.workflow import WorkflowStep

# import the new dialog (make sure the module path matches where you save it)
from gui.dynamic_table_config_dialog import DynamicTableConfigDialog


class StepEditorDialog(QDialog):
    def __init__(self, parent=None, step=None, workflow_screen=None, presets=None, presets_path="presets.json"):
        super().__init__(parent)
        self.setWindowTitle("Step Editor")
        self.step = step
        self.workflow_screen = workflow_screen  # 👈 screen fixed by workflow

        # --- Load presets ---
        if presets is not None:
            self.presets = presets
        else:
            self.presets = {}
            if os.path.exists(presets_path):
                try:
                    with open(presets_path, "r", encoding="utf-8") as f:
                        self.presets = json.load(f)
                except Exception as e:
                    print("⚠️ Failed to load presets.json:", e)

        # Build reverse lookup: screen → team
        self.screen_to_team = {}
        for team, screens in self.presets.items():
            for screen in screens:
                self.screen_to_team[screen] = team

        # --- Safe defaults if step is None ---
        selector_key = step.selector_key if step else ""
        action = step.action if step else ""
        params = step.params if step else {}

        parts = selector_key.split("/") if selector_key else ["", "", "", ""]
        # parts[0] = team, parts[1] = screen, parts[2] = tab, parts[3] = field

        # --- Layout ---
        main = QVBoxLayout(self)

        # --- Screen (removed combo, just label) ---
        screen_label = QLabel(f"Screen: {self.workflow_screen}")
        main.addWidget(screen_label)

        # --- Tab Combo ---
        self.tab_combo = QComboBox()
        main.addWidget(QLabel("Tab:"))
        main.addWidget(self.tab_combo)

        # --- Field Combo ---
        self.field_combo = QComboBox()
        main.addWidget(QLabel("Field:"))
        main.addWidget(self.field_combo)

        # --- Action Combo ---
        self.action_combo = QComboBox()

        self.selector_required_actions = {
            "click", "input", "switch_to_frame", "wait_for",
            "extract", "extract_dynamic_table", "extract_dynamic_table2",
            "extract_table_atm", "extract_dynamic_table_flat", "extract_table_risk_code",
            "extract_table_general_information", "extract_table_general_sweep",
            "extract_table_general_others", "extract_table_general_linkages",
            "extract_table_general_infotype", "extract_table_general_block",
            "extract_table_general_charge", "extract_fallback", "extract_table_new_atm"
        }

        self.selectorless_actions = {
            "switch_to_new_window",
            "close_window_and_return",
            "close_popup_and_return"
        }

        all_actions = sorted(self.selector_required_actions | self.selectorless_actions)
        self.action_combo.addItems(all_actions)
        self.action_combo.setCurrentText(action if action else "click")

        main.addWidget(QLabel("Action:"))
        main.addWidget(self.action_combo)

        # --- Parameter templates controls ---
        param_controls = QHBoxLayout()
        self.param_template_combo = QComboBox()
        self.param_template_combo.addItems([
            "conditions", "slices", "order_by",
            "target_columns", "output_map", "suffix_spacing"
        ])
        self.add_param_btn = QPushButton("Insert Parameter")
        param_controls.addWidget(QLabel("Insert template:"))
        param_controls.addWidget(self.param_template_combo)
        param_controls.addWidget(self.add_param_btn)
        main.addLayout(param_controls)

        # --- Dynamic Table Configurator button (hidden by default) ---
        self.dynamic_table_btn = QPushButton("🧠 Open Dynamic Table Configurator…")
        self.dynamic_table_btn.setToolTip("Open the visual configurator for dynamic table extraction")
        self.dynamic_table_btn.hide()
        main.addWidget(self.dynamic_table_btn)

        # --- Params JSON editor ---
        self.params_edit = QPlainTextEdit()
        self.params_edit.setPlainText(json.dumps(params, indent=2))
        main.addWidget(QLabel("Parameters (JSON):"))
        main.addWidget(self.params_edit)

        # --- Dialog buttons ---
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        main.addWidget(buttons)

        # --- Connectors ---
        self.add_param_btn.clicked.connect(self.insert_param_template)
        self.tab_combo.currentTextChanged.connect(self.populate_fields)
        self.action_combo.currentTextChanged.connect(self.on_action_changed)
        self.dynamic_table_btn.clicked.connect(self.open_dynamic_table_configurator)

        # --- Initialize cascading combos ---
        self.populate_tabs(self.workflow_screen)

        if parts[2]:
            self.tab_combo.setCurrentText(parts[2])
        self.populate_fields(self.tab_combo.currentText())

        if parts[3]:
            self.field_combo.setCurrentText(parts[3])

        # Apply action-specific logic on load
        self.on_action_changed(self.action_combo.currentText())

    def populate_tabs(self, screen):
        self.tab_combo.clear()
        team = self.screen_to_team.get(screen)
        if team and screen in self.presets.get(team, {}):
            self.tab_combo.addItems(sorted(self.presets[team][screen].keys()))
        if self.tab_combo.count() > 0:
            self.populate_fields(self.tab_combo.currentText())
        else:
            self.field_combo.clear()

    def populate_fields(self, tab):
        self.field_combo.clear()
        screen = self.workflow_screen
        team = self.screen_to_team.get(screen)
        if team and screen and tab and tab in self.presets.get(team, {}).get(screen, {}):
            self.field_combo.addItems(sorted(self.presets[team][screen][tab].keys()))

    def get_field_info(self, team, screen, tab, field):
        """Get full field info dict from presets, if available."""
        return self.presets.get(team, {}).get(screen, {}).get(tab, {}).get(field, {})

    def get_preset_columns_for_current_field(self):
        """Return a list of available column names for the currently selected field from presets."""
        screen = self.workflow_screen
        tab = self.tab_combo.currentText().strip()
        field = self.field_combo.currentText().strip()
        team = self.screen_to_team.get(screen, "")
        field_info = self.get_field_info(team, screen, tab, field)
        columns = []
        if field_info:
            # Prefer explicit column_map keys
            column_map = field_info.get("column_map", {})
            if isinstance(column_map, dict) and column_map:
                columns = list(column_map.keys())
            else:
                # Fallback: header_keywords (display them as available 'columns' to pick)
                header_keywords = field_info.get("header_keywords", [])
                if header_keywords:
                    columns = header_keywords.copy()
        return columns

    def insert_param_template(self):
        template_name = self.param_template_combo.currentText()
        templates = {
            "conditions": {"conditions": [{"field": "", "op": "", "value": []}]},
            "slices": {"slices": [{"field": "", "slice": [0, 0]}]},
            "order_by": {"order_by": {"field": "", "values": []}},
            "target_columns": {"target_columns": []},
            "output_map": {"output_map": {"": ""}},
            "suffix_spacing": {"suffix_spacing": {"": False}}
        }

        try:
            current = json.loads(self.params_edit.toPlainText().strip() or "{}")
        except json.JSONDecodeError:
            current = {}

        snippet = templates[template_name]
        for key, value in snippet.items():
            if key not in current:
                current[key] = value
            elif isinstance(current[key], list) and isinstance(value, list):
                current[key].extend(value)
            else:
                current[key] = value

        self.params_edit.setPlainText(json.dumps(current, indent=2))

    def on_action_changed(self, action_name: str):
        if action_name in self.selectorless_actions:
            self.tab_combo.setEnabled(False)
            self.field_combo.setEnabled(False)
            self.dynamic_table_btn.hide()
            return
        else:
            self.tab_combo.setEnabled(True)
            self.field_combo.setEnabled(True)

        # Show configurator button only when dynamic_table field is selected
        if action_name == "extract_dynamic_table":
            screen = self.workflow_screen
            tab = self.tab_combo.currentText().strip()
            field = self.field_combo.currentText().strip()
            team = self.screen_to_team.get(screen, "")
            field_info = self.get_field_info(team, screen, tab, field)

            if field_info and field_info.get("type") == "dynamic_table":
                self.dynamic_table_btn.show()
            else:
                self.dynamic_table_btn.hide()
        else:
            self.dynamic_table_btn.hide()

    def open_dynamic_table_configurator(self):
        """Open the standalone Dynamic Table Configurator dialog."""
        screen = self.workflow_screen
        tab = self.tab_combo.currentText().strip()
        field = self.field_combo.currentText().strip()
        team = self.screen_to_team.get(screen, "")
        field_info = self.get_field_info(team, screen, tab, field) or {}

        # get preset columns (list of names)
        preset_columns = self.get_preset_columns_for_current_field()

        # pass existing params (if valid JSON) to dialog
        try:
            existing_params = json.loads(self.params_edit.toPlainText() or "{}")
        except Exception:
            existing_params = {}

        dlg = DynamicTableConfigDialog(self, preset_columns=preset_columns, preset_field_info=field_info, existing_config=existing_params)
        if dlg.exec() == QDialog.Accepted:
            try:
                new_config = dlg.get_config()
                # merge into current params: replace top-level keys with ones returned
                try:
                    current = json.loads(self.params_edit.toPlainText().strip() or "{}")
                except Exception:
                    current = {}
                current.update(new_config)
                self.params_edit.setPlainText(json.dumps(current, indent=2))
            except Exception as e:
                QMessageBox.warning(self, "Configurator Error", f"Configurator returned invalid config: {e}")

    def get_step(self):
        if self.step is None:
            self.step = WorkflowStep(action="", selector_key="", params={})

        action = self.action_combo.currentText().strip()
        self.step.action = action

        if action in self.selectorless_actions:
            self.step.selector_key = None
        else:
            screen = self.workflow_screen
            tab = self.tab_combo.currentText().strip()
            field = self.field_combo.currentText().strip()
            team = self.screen_to_team.get(screen, "")
            self.step.selector_key = "/".join([team, screen, tab, field])

        try:
            self.step.params = json.loads(self.params_edit.toPlainText())
        except json.JSONDecodeError:
            self.step.params = {}

        return self.step
