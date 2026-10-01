from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QComboBox, QFormLayout, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,QTreeWidget, QTreeWidgetItem,QTabWidget
)



from core import presets  # Import the presets module for managing preset data

# Preset Manager Dialog
class PresetManager(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Presets")
        self.setFixedSize(800, 500)

        # Tree widget for Team > Screen > Tab
        self.preset_tree = QTreeWidget()
        self.preset_tree.setHeaderLabels(["Team / Screen / Tab"])
        self.preset_tree.setColumnCount(1)

        self.load_presets()

        # Buttons
        self.add_button = QPushButton("Add Preset")
        self.edit_button = QPushButton("Edit Preset")
        self.delete_button = QPushButton("Delete Preset")

        # Field Table (initially empty)
        self.fields_table = QTableWidget(0, 4)
        self.fields_table.setHorizontalHeaderLabels(["Field Name", "Type", "Selector", "Use?"])
        self.fields_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.fields_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # Layouts
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)

        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)  # Reduce margins (left, top, right, bottom)
        layout.setSpacing(6)  # Reduce spacing between widgets
        layout.addWidget(QLabel("Preset Navigator:"))
        layout.addWidget(self.preset_tree)
        layout.addLayout(button_layout)
        layout.addWidget(QLabel("Fields:"))
        layout.addWidget(self.fields_table)
        self.setLayout(layout)

        # Connections
        self.preset_tree.itemClicked.connect(self.on_tree_item_clicked)
        self.add_button.clicked.connect(self.add_preset)
        self.edit_button.clicked.connect(self.edit_preset)
        self.delete_button.clicked.connect(self.delete_preset)

    def load_presets(self):
        self.preset_tree.clear()
        all_presets = presets.load_presets()

        for team, screens in all_presets.items():
            team_item = QTreeWidgetItem([team])
            for screen, tabs in screens.items():
                screen_item = QTreeWidgetItem([screen])
                for tab in tabs:
                    tab_item = QTreeWidgetItem([tab])
                    screen_item.addChild(tab_item)
                team_item.addChild(screen_item)
            self.preset_tree.addTopLevelItem(team_item)

    def on_tree_item_clicked(self, item):
        """When user selects a tab node, load its fields."""
        if not item.parent() or not item.parent().parent():
            return  # Only proceed if it's a tab level node

        tab = item.text(0)
        screen = item.parent().text(0)
        team = item.parent().parent().text(0)

        fields = presets.get_fields(team, screen, tab)
        self.populate_fields_table(fields)

    def populate_fields_table(self, fields):
        self.fields_table.setRowCount(0)
        for name, data in fields.items():
            if name == "iframe":
                continue  # Skip iframe selector
            row = self.fields_table.rowCount()
            self.fields_table.insertRow(row)
            self.fields_table.setItem(row, 0, QTableWidgetItem(name))
            self.fields_table.setItem(row, 1, QTableWidgetItem(data.get("type", "")))

            selector_item = QTableWidgetItem()
            if data.get("type") == "css":
                selector = data.get("selector", "")
                selector_item.setText(selector)
            elif data.get("type") == "table_lookup":
                selector = f"{data.get('table_id', '')} | {data.get('label', '')}"
                selector_item.setText(selector)
            elif data.get("type") == "dynamic_table":
                # Show summary in cell, full config in tooltip
                summary = f"{data.get('selector', '')} | {','.join(data.get('header_keywords', []))}"
                selector_item.setText(summary)
                # Tooltip with all config
                tooltip = (
                    f"Selector: {data.get('selector', '')}\n"
                    f"Header Keywords: {', '.join(data.get('header_keywords', []))}\n"
                    f"Row Selector: {data.get('row_selector', '')}\n"
                    f"Column Map: {data.get('column_map', {})}"
                )
                selector_item.setToolTip(tooltip)
            else:
                selector_item.setText("")
            self.fields_table.setItem(row, 2, selector_item)

            use_item = QTableWidgetItem()
            use_item.setCheckState(Qt.Checked if data.get("enabled", True) else Qt.Unchecked)
            self.fields_table.setItem(row, 3, use_item)

    
    def add_preset(self):
        form = PresetForm(self)
        if form.exec() == QDialog.Accepted:
            data = form.get_data()
            if not data:
                return

            team = data["team"]
            screen = data["screen"]
            tab_fields = data["fields"]  # {tab_name: field_dict}

            all_presets = presets.load_presets()

            if team not in all_presets:
                all_presets[team] = {}
            if screen not in all_presets[team]:
                all_presets[team][screen] = {}

            duplicate_tabs = []
            for tab, fields in tab_fields.items():
                if tab in all_presets[team][screen]:
                    duplicate_tabs.append(tab)
                else:
                    all_presets[team][screen][tab] = fields

            if duplicate_tabs:
                QMessageBox.warning(self, "Duplicate Tabs",
                                    f"The following tabs already exist and were skipped: {', '.join(duplicate_tabs)}")

            presets.save_presets(all_presets)
            self.load_presets()
            QMessageBox.information(self, "Success", f"Preset(s) added to {team} > {screen}")



    
    def edit_preset(self):
        item = self.preset_tree.currentItem()
        if not item:
            QMessageBox.warning(self, "Selection Error", "Please select a team, screen, or tab.")
            return

        # Determine what level was clicked
        if item.parent() and item.parent().parent():
            # Tab level
            tab = item.text(0)
            screen = item.parent().text(0)
            team = item.parent().parent().text(0)
        elif item.parent():
            # Screen level
            tab = None
            screen = item.text(0)
            team = item.parent().text(0)
        else:
            QMessageBox.warning(self, "Selection Error", "Please select at least a screen under a team.")
            return

        all_presets = presets.load_presets()
        existing_tabs = all_presets.get(team, {}).get(screen, {})

        if not existing_tabs:
            QMessageBox.warning(self, "Not Found", f"No tabs found under {team} > {screen}")
            return

        form = PresetForm(self)
        form.team_input.setText(team)
        form.screen_input.setText(screen)
        form.team_input.setDisabled(True)
        form.screen_input.setDisabled(True)

        # Load all tabs into form
        form.tab_widget.clear()
        for tab_name, fields in existing_tabs.items():
            table = QTableWidget(0, 3)
            table.setHorizontalHeaderLabels(["Field Name", "Type", "Selector"])
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setSelectionBehavior(QAbstractItemView.SelectRows)

            for name, data in fields.items():
                if name == "iframe":
                    continue
                row = table.rowCount()
                table.insertRow(row)
                table.setItem(row, 0, QTableWidgetItem(name))
                table.setItem(row, 1, QTableWidgetItem(data.get("type", "")))

                if data.get("type") == "css":
                    selector = data.get("selector", "")
                elif data.get("type") == "table_lookup":
                    selector = f"{data.get('table_id', '')} | {data.get('label', '')}"
                else:
                    selector = ""
                table.setItem(row, 2, QTableWidgetItem(selector))

            form.tab_widget.addTab(table, tab_name)

        if form.exec() == QDialog.Accepted:
            data = form.get_data()
            if not data:
                return

            new_tabs = data["fields"]
            all_presets[team][screen] = new_tabs
            presets.save_presets(all_presets)

            self.load_presets()
            QMessageBox.information(self, "Success", f"Preset updated for {team} > {screen}")


        
    def delete_preset(self):
        QMessageBox.information(self, "Not Implemented", "Deleting presets will be implemented next.")

# Preset Form for adding/editing presets
class PresetForm(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preset Form")
        self.setMinimumSize(700, 550)

        self.team_input = QLineEdit()
        self.screen_input = QLineEdit()
        self.iframe_selector_input = QLineEdit()

        form_layout = QFormLayout()
        form_layout.setContentsMargins(8, 8, 8, 8)
        form_layout.setSpacing(6)
        form_layout.addRow("Team:", self.team_input)
        form_layout.addRow("Screen:", self.screen_input)
        form_layout.addRow("Iframe Selector (optional):", self.iframe_selector_input)

        # Tab Widget for Tabs (each with its own field table)
        self.tab_widget = QTabWidget()

        self.add_tab_button = QPushButton("Add Tab")
        self.remove_tab_button = QPushButton("Remove Tab")

        tab_button_layout = QHBoxLayout()
        tab_button_layout.addWidget(self.add_tab_button)
        tab_button_layout.addWidget(self.remove_tab_button)

        self.add_tab_button.clicked.connect(self.add_tab)
        self.remove_tab_button.clicked.connect(self.remove_selected_tab)

        # Field Buttons
        self.add_field_button = QPushButton("Add Field")
        self.edit_field_button = QPushButton("Edit Selected")
        self.remove_field_button = QPushButton("Remove Selected")

        field_buttons_layout = QHBoxLayout()
        field_buttons_layout.addWidget(self.add_field_button)
        field_buttons_layout.addWidget(self.edit_field_button)
        field_buttons_layout.addWidget(self.remove_field_button)

        self.add_field_button.clicked.connect(self.add_field)
        self.edit_field_button.clicked.connect(self.edit_selected_field)
        self.remove_field_button.clicked.connect(self.remove_selected_field)

        # Dialog buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Layout
        layout = QVBoxLayout()
        layout.addLayout(form_layout)
        layout.addWidget(QLabel("Tabs & Fields:"))
        layout.addWidget(self.tab_widget)
        layout.addLayout(tab_button_layout)
        layout.addLayout(field_buttons_layout)
        layout.addWidget(self.button_box)

        self.setLayout(layout)

    def add_tab(self):
        from PySide6.QtWidgets import QInputDialog

        tab_name, ok = QInputDialog.getText(self, "Add Tab", "Tab name:")
        if ok and tab_name:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == tab_name:
                    QMessageBox.warning(self, "Duplicate", "Tab name already exists.")
                    return
            table = QTableWidget(0, 3)
            table.setHorizontalHeaderLabels(["Field Name", "Type", "Selector"])
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.tab_widget.addTab(table, tab_name)

    def remove_selected_tab(self):
        index = self.tab_widget.currentIndex()
        if index >= 0:
            self.tab_widget.removeTab(index)

    def get_current_table(self):
        return self.tab_widget.currentWidget()

    def add_field(self):
        from PySide6.QtWidgets import QDialog, QFormLayout, QDialogButtonBox, QListWidget, QListWidgetItem

        table = self.get_current_table()
        if not table:
            QMessageBox.warning(self, "No Tab", "Please create/select a tab first.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Add Field")
        layout = QFormLayout(dialog)

        field_name = QLineEdit()
        field_type = QComboBox()
        field_type.addItems(["css", "table_lookup", "dynamic_table"])
        field_selector = QLineEdit()
        label_input = QLineEdit()

        # Dynamic table extras
        header_keywords_input = QLineEdit()
        row_selector_input = QLineEdit()
        column_map_input = QLineEdit()  # JSON string: {"col1": 1, ...}

        layout.addRow("Field Name:", field_name)
        layout.addRow("Type:", field_type)
        layout.addRow("Selector / Table ID:", field_selector)
        layout.addRow("Label (only for table_lookup):", label_input)
        layout.addRow("Header Keywords (comma):", header_keywords_input)
        layout.addRow("Row Selector:", row_selector_input)
        layout.addRow("Column Map (JSON):", column_map_input)

        # Hide dynamic_table extras by default
        header_keywords_input.hide()
        row_selector_input.hide()
        column_map_input.hide()

        def on_type_change(idx):
            is_dynamic = field_type.currentText() == "dynamic_table"
            header_keywords_input.setVisible(is_dynamic)
            row_selector_input.setVisible(is_dynamic)
            column_map_input.setVisible(is_dynamic)
            label_input.setVisible(field_type.currentText() == "table_lookup")

        field_type.currentIndexChanged.connect(on_type_change)
        on_type_change(0)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(buttons)

        def accept():
            name = field_name.text().strip()
            type_ = field_type.currentText()
            selector = field_selector.text().strip()
            label = label_input.text().strip()

            if not name or not selector:
                QMessageBox.warning(dialog, "Input Error", "Field name and selector are required.")
                return

            # Check for duplicate
            for r in range(table.rowCount()):
                if table.item(r, 0).text() == name:
                    QMessageBox.warning(dialog, "Duplicate Field", f"Field '{name}' already exists.")
                    return

            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(name))
            table.setItem(row, 1, QTableWidgetItem(type_))

            if type_ == "css":
                table.setItem(row, 2, QTableWidgetItem(selector))
            elif type_ == "table_lookup":
                table.setItem(row, 2, QTableWidgetItem(f"{selector} | {label}"))
            elif type_ == "dynamic_table":
                # Show summary in selector column
                summary = f"{selector} | {header_keywords_input.text()}"
                table.setItem(row, 2, QTableWidgetItem(summary))
                # Store extra config as Qt.UserRole
                table.item(row, 2).setData(Qt.UserRole, {
                    "selector": selector,
                    "header_keywords": [k.strip() for k in header_keywords_input.text().split(",") if k.strip()],
                    "row_selector": row_selector_input.text().strip(),
                    "column_map": column_map_input.text().strip(),
                })

            dialog.accept()

        buttons.accepted.connect(accept)
        buttons.rejected.connect(dialog.reject)
        dialog.exec()

    def edit_selected_field(self):
        table = self.get_current_table()
        if not table:
            QMessageBox.warning(self, "No Tab", "Please select a tab first.")
            return

        selected = table.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Selection Error", "Please select a field to edit.")
            return

        name = table.item(selected, 0).text()
        type_ = table.item(selected, 1).text()
        raw_selector = table.item(selected, 2).text()

        selector = raw_selector
        label = ""
        if type_ == "table_lookup" and " | " in raw_selector:
            selector, label = raw_selector.split(" | ", 1)

        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Field")
        layout = QFormLayout(dialog)

        field_name = QLineEdit(name)
        field_type = QComboBox()
        field_type.addItems(["css", "table_lookup"])
        field_type.setCurrentText(type_)
        field_selector = QLineEdit(selector)
        label_input = QLineEdit(label)

        layout.addRow("Field Name:", field_name)
        layout.addRow("Type:", field_type)
        layout.addRow("Selector / Table ID:", field_selector)
        layout.addRow("Label (only for table_lookup):", label_input)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(buttons)

        def accept():
            new_name = field_name.text().strip()
            new_type = field_type.currentText()
            new_selector = field_selector.text().strip()
            new_label = label_input.text().strip()

            if not new_name or not new_selector:
                QMessageBox.warning(dialog, "Input Error", "Field name and selector are required.")
                return

            table.setItem(selected, 0, QTableWidgetItem(new_name))
            table.setItem(selected, 1, QTableWidgetItem(new_type))
            table.setItem(selected, 2, QTableWidgetItem(
                new_selector if new_type == "css" else f"{new_selector} | {new_label}"
            ))

            dialog.accept()

        buttons.accepted.connect(accept)
        buttons.rejected.connect(dialog.reject)
        dialog.exec()

    def remove_selected_field(self):
        table = self.get_current_table()
        if table:
            selected = table.currentRow()
            if selected >= 0:
                table.removeRow(selected)

    def get_data(self):
        """Return all data in JSON-compatible format."""
        team = self.team_input.text().strip()
        screen = self.screen_input.text().strip()
        iframe_selector = self.iframe_selector_input.text().strip()

        if not team or not screen:
            QMessageBox.warning(self, "Input Error", "Team and Screen are required.")
            return None
        
        if self.tab_widget.count() == 0:
            QMessageBox.warning(self, "Missing Tab", "Please add at least one tab before saving.")
            return None


        result = {
            "team": team,
            "screen": screen,
            "fields": {},  # tab name => fields dict
        }

        for i in range(self.tab_widget.count()):
            tab_name = self.tab_widget.tabText(i)
            table = self.tab_widget.widget(i)

            tab_fields = {}
            for row in range(table.rowCount()):
                name = table.item(row, 0).text()
                type_ = table.item(row, 1).text()
                raw = table.item(row, 2).text()

                if type_ == "css":
                    tab_fields[name] = {"type": "css", "selector": raw}
                elif type_ == "table_lookup":
                    if " | " in raw:
                        table_id, label = raw.split(" | ", 1)
                    else:
                        table_id, label = raw, ""
                    tab_fields[name] = {"type": "table_lookup", "table_id": table_id, "label": label}
                elif type_ == "dynamic_table":
                    # Get config from Qt.UserRole
                    config = table.item(row, 2).data(Qt.UserRole)
                    if config:
                        import json
                        try:
                            column_map = json.loads(config.get("column_map", "{}"))
                        except Exception:
                            column_map = {}
                        tab_fields[name] = {
                            "type": "dynamic_table",
                            "selector": config.get("selector", ""),
                            "header_keywords": config.get("header_keywords", []),
                            "row_selector": config.get("row_selector", ""),
                            "column_map": column_map,
                        }
                    else:
                        # fallback: parse from summary string
                        tab_fields[name] = {"type": "dynamic_table", "selector": raw}

            if iframe_selector:
                tab_fields["iframe"] = {"type": "css", "selector": iframe_selector}

            result["fields"][tab_name] = tab_fields

        return result