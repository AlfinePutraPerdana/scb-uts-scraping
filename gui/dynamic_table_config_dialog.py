# gui/dynamic_table_config_dialog.py
import json
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QDialogButtonBox, QPushButton,
    QTableWidget, QTableWidgetItem, QHBoxLayout, QWidget, QComboBox, QLineEdit,
    QHeaderView, QToolTip, QFormLayout, QSpinBox, QCheckBox, QAbstractItemView
)
from PySide6.QtCore import Qt


class DynamicTableConfigDialog(QDialog):
    """Visual configurator for extract_dynamic_table parameters."""
    def __init__(self, parent=None, preset_columns=None, preset_field_info=None, existing_config=None):
        super().__init__(parent)
        self.setWindowTitle("Dynamic Table Configurator")
        self.resize(850, 600)

        self.preset_columns = preset_columns or []
        self.existing_config = existing_config or {}
        self.conditions = self.existing_config.get("conditions", [])

        # --- Operator list & tooltips ---
        self.operator_tooltips = {
            "equals": "Checks if the field value is exactly equal to the given value.",
            "in": "Checks if the field value is one of the values in the list.",
            "not_in": "Checks if the field value is NOT in the list.",
            "startswith": "Checks if the field value starts with the given text.",
            "contains": "Checks if the field value contains the given word or phrase.",
            "exists": "Passes if the field exists and is not empty, '-', or 'N/A'.",
            "not_exists": "Passes if the field is missing, empty, '-', or 'N/A'."
        }
        self.supported_operators = list(self.operator_tooltips.keys())

        # --------------------------
        # MAIN LAYOUT
        # --------------------------
        main_layout = QVBoxLayout(self)

        # --- General Parameters ---
        param_form = QFormLayout()
        self.max_rows_spin = QSpinBox()
        self.max_rows_spin.setRange(1, 9999)
        self.max_rows_spin.setValue(self.existing_config.get("max_rows", 15))
        self.max_rows_spin.setToolTip("Maximum number of table rows to process.")

        self.max_matches_spin = QSpinBox()
        self.max_matches_spin.setRange(1, 9999)
        self.max_matches_spin.setValue(self.existing_config.get("max_matches", 5))
        self.max_matches_spin.setToolTip("Maximum number of matched records allowed.")

        param_form.addRow("Max Rows:", self.max_rows_spin)
        param_form.addRow("Max Matches:", self.max_matches_spin)
        main_layout.addLayout(param_form)

        # --------------------------
        # TARGET COLUMNS
        # --------------------------
        main_layout.addWidget(QLabel("Target Columns (select and rename outputs):"))
        self.column_table = QTableWidget(0, 3)
        self.column_table.setHorizontalHeaderLabels(["Extract?", "Field Name", "Output Map"])
        self.column_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.column_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.column_table.setSelectionBehavior(QAbstractItemView.SelectRows)

        self._populate_target_columns()
        main_layout.addWidget(self.column_table)

        # --------------------------
        # CONDITIONS
        # --------------------------
        main_layout.addWidget(QLabel("Conditions (optional):"))
        self.condition_table = QTableWidget(0, 3)
        self.condition_table.setHorizontalHeaderLabels(["Field", "Operator", "Value / List"])
        self.condition_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.condition_table)

        btns_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Add Condition")
        remove_btn = QPushButton("➖ Remove Selected")
        btns_layout.addWidget(add_btn)
        btns_layout.addWidget(remove_btn)
        main_layout.addLayout(btns_layout)

        add_btn.clicked.connect(self.add_condition_row)
        remove_btn.clicked.connect(self.remove_selected_rows)

        # --- Load existing conditions ---
        for cond in self.conditions:
            self.add_condition_row(cond)

        # --------------------------
        # DIALOG BUTTONS
        # --------------------------
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        main_layout.addWidget(buttons)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

    # --------------------------
    # POPULATE TABLES
    # --------------------------
    def _populate_target_columns(self):
        """Fill target column table with selectable, editable mappings."""
        target_columns = self.existing_config.get("target_columns", [])
        output_map = self.existing_config.get("output_map", {})

        self.column_table.setRowCount(len(self.preset_columns))
        for i, col in enumerate(self.preset_columns):
            chk = QCheckBox()
            chk.setChecked(col in target_columns or not target_columns)
            chk.setToolTip("Check to include this column in extraction.")
            self.column_table.setCellWidget(i, 0, chk)

            item = QTableWidgetItem(col)
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            item.setToolTip(f"Source column name: {col}")
            self.column_table.setItem(i, 1, item)

            mapped_name = output_map.get(col, col.upper())
            mapped_item = QTableWidgetItem(mapped_name)
            mapped_item.setToolTip("Output column name in Excel (editable)")
            self.column_table.setItem(i, 2, mapped_item)

        self.column_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.column_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)

    # --------------------------
    # CONDITIONS
    # --------------------------
    def add_condition_row(self, cond=None):
        row = self.condition_table.rowCount()
        self.condition_table.insertRow(row)

        # --- Field combo ---
        field_combo = QComboBox()
        field_combo.addItems(self.preset_columns)
        if cond and cond.get("field"):
            field_combo.setCurrentText(cond["field"])
        self.condition_table.setCellWidget(row, 0, field_combo)

        # --- Operator combo ---
        op_combo = QComboBox()
        for op in self.supported_operators:
            op_combo.addItem(op)
            idx = op_combo.findText(op)
            op_combo.setItemData(idx, self.operator_tooltips[op], Qt.ToolTipRole)
        if cond and cond.get("op"):
            op_combo.setCurrentText(cond["op"])
        self.condition_table.setCellWidget(row, 1, op_combo)

        # --- Value edit ---
        value_edit = QLineEdit()
        if cond:
            value = cond.get("value", "")
            if isinstance(value, list):
                value_edit.setText(", ".join(map(str, value)))
            else:
                value_edit.setText(str(value))
        self.condition_table.setCellWidget(row, 2, value_edit)

        # Tooltip preview
        op_combo.currentTextChanged.connect(
            lambda op, combo=op_combo: QToolTip.showText(combo.mapToGlobal(combo.rect().center()), self.operator_tooltips.get(op, ""))
        )

    def remove_selected_rows(self):
        selected_rows = sorted({i.row() for i in self.condition_table.selectedIndexes()}, reverse=True)
        for row in selected_rows:
            self.condition_table.removeRow(row)

    # --------------------------
    # OUTPUT
    # --------------------------
    def get_config(self):
        """Return JSON-like dict representing the full params for extract_dynamic_table."""
        # General params
        max_rows = self.max_rows_spin.value()
        max_matches = self.max_matches_spin.value()

        # Target columns
        target_columns = []
        output_map = {}
        suffix_spacing = {}

        for i in range(self.column_table.rowCount()):
            chk = self.column_table.cellWidget(i, 0)
            if chk.isChecked():
                field = self.column_table.item(i, 1).text()
                mapped = self.column_table.item(i, 2).text().strip()
                target_columns.append(field)
                output_map[field] = mapped
                suffix_spacing[mapped] = True  # Default

        # Conditions
        conditions = []
        for row in range(self.condition_table.rowCount()):
            field = self.condition_table.cellWidget(row, 0).currentText().strip()
            op = self.condition_table.cellWidget(row, 1).currentText().strip()
            value_text = self.condition_table.cellWidget(row, 2).text().strip()

            if "," in value_text:
                value = [v.strip() for v in value_text.split(",") if v.strip()]
            else:
                value = value_text

            conditions.append({"field": field, "op": op, "value": value})

        # Combine into params dict
        params = {
            "max_rows": max_rows,
            "max_matches": max_matches,
            "conditions": conditions,
            "slices": self.existing_config.get("slices", []),
            "order_by": self.existing_config.get("order_by", {}),
            "target_columns": target_columns,
            "output_map": output_map,
            "suffix_spacing": suffix_spacing
        }

        return params
