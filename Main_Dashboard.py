import sys
import json
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QPushButton, QStackedWidget, QMessageBox, QLabel,
    QLineEdit, QFileDialog, QFormLayout, QGroupBox,
    QDialog, QDialogButtonBox, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt

import FKBP5_workflow
import SingleStudy_workflow

CONFIG_PATH = "studies_config.json"


class StudyEditorDialog(QDialog):
    def __init__(self, parent=None, study_name="", study_data=None):
        super().__init__(parent)
        self.setWindowTitle("Add Study" if not study_name else "Edit Study")
        self.setMinimumWidth(620)
        self.is_edit = bool(study_name)

        layout = QVBoxLayout(self)

        form_group = QGroupBox("Study Details")
        form = QFormLayout(form_group)

        self.name_field = QLineEdit(study_name)
        self.name_field.setPlaceholderText("e.g. IFS+SHP099+Medigel 2026")
        if self.is_edit:
            self.name_field.setEnabled(False)
        form.addRow(QLabel("Study Name:"), self.name_field)

        self.path_fields = {}
        data = study_data or {}
        paths_meta = [
            ("raw_data_root",    "Raw Data Folder:",          True),
            ("master_file",      "Master Excel File:",        False),
            ("measurement_file", "Measurements Excel File:",  False),
            ("output_folder",    "CSV Output Folder:",        True),
        ]
        for key, label, is_dir in paths_meta:
            row = QHBoxLayout()
            le = QLineEdit(data.get(key, ""))
            btn = QPushButton("Browse")
            btn.clicked.connect(lambda checked, l=le, d=is_dir: self._browse(l, d))
            row.addWidget(le)
            row.addWidget(btn)
            form.addRow(QLabel(label), row)
            self.path_fields[key] = le

        layout.addWidget(form_group)

        gm_group = QGroupBox("Group Map  (file-prefix  →  group name)")
        gm_layout = QVBoxLayout(gm_group)

        self.group_table = QTableWidget(0, 2)
        self.group_table.setHorizontalHeaderLabels(["Prefix", "Group Name"])
        self.group_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch)
        self.group_table.setMinimumHeight(130)

        for prefix, name in data.get("group_map", {}).items():
            self._add_group_row(prefix, name)

        gm_layout.addWidget(self.group_table)

        gm_btn_row = QHBoxLayout()
        add_row_btn = QPushButton("+ Add Row")
        add_row_btn.clicked.connect(lambda: self._add_group_row())
        rm_row_btn = QPushButton("- Remove Selected Row")
        rm_row_btn.clicked.connect(self._remove_group_row)
        gm_btn_row.addWidget(add_row_btn)
        gm_btn_row.addWidget(rm_row_btn)
        gm_btn_row.addStretch()
        gm_layout.addLayout(gm_btn_row)

        layout.addWidget(gm_group)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _browse(self, line_edit, is_directory):
        if is_directory:
            path = QFileDialog.getExistingDirectory(self, "Select Directory")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select File", "", "Excel Files (*.xlsx *.xls)")
        if path:
            line_edit.setText(os.path.normpath(path))

    def _add_group_row(self, prefix="", group_name=""):
        row = self.group_table.rowCount()
        self.group_table.insertRow(row)
        self.group_table.setItem(row, 0, QTableWidgetItem(prefix))
        self.group_table.setItem(row, 1, QTableWidgetItem(group_name))

    def _remove_group_row(self):
        selected = self.group_table.currentRow()
        if selected >= 0:
            self.group_table.removeRow(selected)

    def _validate_and_accept(self):
        if not self.name_field.text().strip():
            QMessageBox.warning(self, "Validation Error",
                                "Study name cannot be empty.")
            return
        if not self.path_fields["raw_data_root"].text().strip():
            QMessageBox.warning(self, "Validation Error",
                                "Raw Data Folder cannot be empty.")
            return
        self.accept()

    def get_study_name(self):
        return self.name_field.text().strip()

    def get_study_data(self):
        group_map = {}
        for row in range(self.group_table.rowCount()):
            p = self.group_table.item(row, 0)
            n = self.group_table.item(row, 1)
            prefix = p.text().strip() if p else ""
            name   = n.text().strip() if n else ""
            if prefix:
                group_map[prefix] = name
        return {
            "raw_data_root":    self.path_fields["raw_data_root"].text().strip(),
            "master_file":      self.path_fields["master_file"].text().strip(),
            "measurement_file": self.path_fields["measurement_file"].text().strip(),
            "output_folder":    self.path_fields["output_folder"].text().strip(),
            "group_map":        group_map,
        }


class MainDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("3-Point Bending Workflow Hub")
        self.setMinimumWidth(650)
        self.config = self.load_config()

        self.central_widget = QWidget()
        self.layout = QVBoxLayout(self.central_widget)

        # Study selector row
        selector_layout = QHBoxLayout()
        selector_layout.addWidget(QLabel("Select Analysis Mode or Study:"))
        self.study_selector = QComboBox()
        self.study_selector.addItem("FKBP5 Genotyping")
        self.study_selector.addItems(self.config.keys())
        self.study_selector.currentIndexChanged.connect(self.toggle_ui)
        selector_layout.addWidget(self.study_selector, 1)
        self.layout.addLayout(selector_layout)

        # Study management buttons
        mgmt_layout = QHBoxLayout()
        btn_add = QPushButton("+ Add Study")
        btn_add.clicked.connect(self.add_study)
        btn_edit = QPushButton("Edit Study")
        btn_edit.clicked.connect(self.edit_study)
        btn_remove = QPushButton("Remove Study")
        btn_remove.clicked.connect(self.remove_study)
        mgmt_layout.addWidget(btn_add)
        mgmt_layout.addWidget(btn_edit)
        mgmt_layout.addWidget(btn_remove)
        mgmt_layout.addStretch()
        self.layout.addLayout(mgmt_layout)

        # Stacked views
        self.stack = QStackedWidget()
        self.layout.addWidget(self.stack)

        self.setup_fkbp5_ui()
        self.setup_single_study_ui()

        self.run_button = QPushButton("Execute Workflow")
        self.run_button.setMinimumHeight(40)
        self.run_button.clicked.connect(self.run_pipeline)
        self.layout.addWidget(self.run_button)

        self.setCentralWidget(self.central_widget)
        self.toggle_ui()

    # ------------------------------------------------------------------
    # Config persistence
    # ------------------------------------------------------------------

    def load_config(self):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            return {}

    def save_config(self):
        with open(CONFIG_PATH, "w") as f:
            json.dump(self.config, f, indent=4)

    def refresh_selector(self, select_name=None):
        self.study_selector.blockSignals(True)
        self.study_selector.clear()
        self.study_selector.addItem("FKBP5 Genotyping")
        self.study_selector.addItems(self.config.keys())
        if select_name and select_name in self.config:
            idx = self.study_selector.findText(select_name)
            self.study_selector.setCurrentIndex(idx)
        self.study_selector.blockSignals(False)
        self.toggle_ui()

    # ------------------------------------------------------------------
    # Study management actions
    # ------------------------------------------------------------------

    def add_study(self):
        dlg = StudyEditorDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        name = dlg.get_study_name()
        if name in self.config:
            QMessageBox.warning(self, "Duplicate Study",
                                f'A study named "{name}" already exists.')
            return
        self.config[name] = dlg.get_study_data()
        self.save_config()
        self.refresh_selector(select_name=name)

    def edit_study(self):
        name = self.study_selector.currentText()
        if name == "FKBP5 Genotyping":
            QMessageBox.information(self, "Not Editable",
                                    "FKBP5 Genotyping is a built-in mode and cannot be edited here.")
            return
        if name not in self.config:
            return
        dlg = StudyEditorDialog(self, study_name=name,
                                study_data=self.config[name])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self.config[name] = dlg.get_study_data()
        self.save_config()
        self.refresh_selector(select_name=name)

    def remove_study(self):
        name = self.study_selector.currentText()
        if name == "FKBP5 Genotyping":
            QMessageBox.information(self, "Not Removable",
                                    "FKBP5 Genotyping is a built-in mode and cannot be removed.")
            return
        if name not in self.config:
            return
        reply = QMessageBox.question(
            self, "Confirm Removal",
            f'Remove "{name}" from the study list?\n\nThis only removes the entry from studies_config.json — no files will be deleted.',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        del self.config[name]
        self.save_config()
        self.refresh_selector()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------

    def setup_fkbp5_ui(self):
        self.fkbp5_widget = QWidget()
        layout = QVBoxLayout(self.fkbp5_widget)

        group = QGroupBox("FKBP5 Genotyping Pipeline Configurations")
        form = QFormLayout(group)

        self.fkbp5_fields = {}
        fields_meta = [
            ("raw_dir",      "Raw Data Folder:",       FKBP5_workflow.DEFAULT_RAW_DATA_ROOT,    True),
            ("tibia_master", "Tibia Master Excel File:",FKBP5_workflow.DEFAULT_TIBIA_MASTER_FILE,False),
            ("femur_master", "Femur Master Excel File:",FKBP5_workflow.DEFAULT_FEMUR_MASTER_FILE,False),
            ("meas_file",    "Measurements File:",      FKBP5_workflow.DEFAULT_MEASUREMENT_FILE, False),
            ("csv_dir",      "CSV Export Folder:",      FKBP5_workflow.DEFAULT_CSV_OUTPUT_DIR,   True),
        ]

        for key, label_text, default_val, is_dir in fields_meta:
            row_layout = QHBoxLayout()
            line_edit = QLineEdit(default_val)
            btn = QPushButton("Browse")
            btn.clicked.connect(lambda checked, le=line_edit,
                                d=is_dir: self.browse_path(le, d))
            row_layout.addWidget(line_edit)
            row_layout.addWidget(btn)
            form.addRow(QLabel(label_text), row_layout)
            self.fkbp5_fields[key] = line_edit

        self.fkbp5_fallback = QComboBox()
        self.fkbp5_fallback.addItems(["Tibia", "Femur"])
        form.addRow(QLabel("Default Bone Fallback:"), self.fkbp5_fallback)

        layout.addWidget(group)
        self.stack.addWidget(self.fkbp5_widget)

    def setup_single_study_ui(self):
        self.single_study_widget = QWidget()
        layout = QVBoxLayout(self.single_study_widget)

        group = QGroupBox("Custom Study Template Parameters")
        form = QFormLayout(group)

        self.ss_fields = {}

        self.ss_sex = QComboBox()
        self.ss_sex.addItems(["Male", "Female", "Mixed"])
        form.addRow(QLabel("Cohort Sex:"), self.ss_sex)

        self.ss_age = QLineEdit("16 Weeks")
        form.addRow(QLabel("Cohort Age:"), self.ss_age)

        self.ss_bone = QComboBox()
        self.ss_bone.addItems(["Femur", "Tibia"])
        form.addRow(QLabel("Target Bone Structural Type:"), self.ss_bone)

        paths_meta = [
            ("raw_dir",  "Raw Data Folder:",          True),
            ("meas_file","Measurements Excel File:",   False),
            ("csv_dir",  "CSV/Master Output Directory:",True),
        ]

        for key, label_text, is_dir in paths_meta:
            row_layout = QHBoxLayout()
            line_edit = QLineEdit()
            btn = QPushButton("Browse")
            btn.clicked.connect(lambda checked, le=line_edit,
                                d=is_dir: self.browse_path(le, d))
            row_layout.addWidget(line_edit)
            row_layout.addWidget(btn)
            form.addRow(QLabel(label_text), row_layout)
            self.ss_fields[key] = line_edit

        layout.addWidget(group)
        self.stack.addWidget(self.single_study_widget)

    def browse_path(self, line_edit, is_directory):
        if is_directory:
            path = QFileDialog.getExistingDirectory(self, "Select Directory")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select File", "", "Excel Files (*.xlsx *.xls)")
        if path:
            line_edit.setText(os.path.normpath(path))

    def toggle_ui(self):
        current_selection = self.study_selector.currentText()
        if current_selection == "FKBP5 Genotyping":
            self.stack.setCurrentIndex(0)
        else:
            self.stack.setCurrentIndex(1)
            study_data = self.config.get(current_selection, {})
            self.ss_fields["raw_dir"].setText(
                study_data.get("raw_data_root", ""))
            self.ss_fields["meas_file"].setText(
                study_data.get("measurement_file", ""))
            self.ss_fields["csv_dir"].setText(
                study_data.get("output_folder", ""))

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def show_message_box(self, title, message, details="", is_error=False):
        msg = QMessageBox(self)
        msg.setIcon(
            QMessageBox.Icon.Critical if is_error else QMessageBox.Icon.Information)
        msg.setWindowTitle(title)
        msg.setText(message)
        if details:
            msg.setInformativeText(details)
        msg.setStyleSheet("""
            QMessageBox { background-color: #ffffff; }
            QLabel { color: #000000; }
            QPushButton { background-color: #f0f0f0; color: #000000; min-width: 75px; }
        """)
        msg.exec()

    # ------------------------------------------------------------------
    # Pipeline execution
    # ------------------------------------------------------------------

    def run_pipeline(self):
        current_study = self.study_selector.currentText()
        self.run_button.setEnabled(False)
        self.run_button.setText("Processing...")
        QApplication.processEvents()

        try:
            if current_study == "FKBP5 Genotyping":
                inputs = {
                    "data_folder":       self.fkbp5_fields["raw_dir"].text(),
                    "tibia_master":      self.fkbp5_fields["tibia_master"].text(),
                    "femur_master":      self.fkbp5_fields["femur_master"].text(),
                    "measurement_path":  self.fkbp5_fields["meas_file"].text(),
                    "csv_out_dir":       self.fkbp5_fields["csv_dir"].text(),
                    "structure_type":    "Split Genders Male/Female",
                    "fallback_bone":     self.fkbp5_fallback.currentText()
                }
                success = FKBP5_workflow.run_workflow(inputs)
                if success:
                    self.show_message_box(
                        "Success",
                        "FKBP5 calculation finalized and master files successfully updated.")
                else:
                    self.show_message_box(
                        "Execution Incomplete",
                        "Check data directory contents.", is_error=True)
            else:
                study_data = self.config.get(current_study, {})
                meas_path = self.ss_fields["meas_file"].text()
                master_dir = os.path.dirname(meas_path)
                master_filename = os.path.basename(
                    study_data.get("master_file", "Default_Master.xlsx"))
                inputs = {
                    "study_name":       current_study,
                    "data_folder":      self.ss_fields["raw_dir"].text(),
                    "measurement_path": meas_path,
                    "csv_out_dir":      self.ss_fields["csv_dir"].text(),
                    "sex":              self.ss_sex.currentText(),
                    "age":              self.ss_age.text(),
                    "bone":             self.ss_bone.currentText(),
                    "group_map":        study_data.get("group_map", {})
                }
                success = SingleStudy_workflow.run_workflow(inputs)
                if success:
                    self.show_message_box(
                        "Success",
                        f"Custom study pipeline executed successfully for: {current_study}")
                else:
                    self.show_message_box(
                        "Execution Incomplete",
                        "Check data directory contents.", is_error=True)

        except Exception as e:
            self.show_message_box(
                "Execution Failure", "An unhandled error occurred:", str(e), is_error=True)
        finally:
            self.run_button.setEnabled(True)
            self.run_button.setText("Execute Workflow")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainDashboard()
    window.show()
    sys.exit(app.exec())
