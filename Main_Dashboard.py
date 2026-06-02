import sys
import json
import os
import glob
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QPushButton, QStackedWidget, QMessageBox, QLabel,
    QLineEdit, QFileDialog, QFormLayout, QGroupBox, QCheckBox,
    QDialog, QDialogButtonBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QColorDialog,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

import FKBP5Heat_workflow
import FKBP5New_workflow
import SingleStudy_workflow
import OverlayGraphs_workflow
import prism_export
from OverlayGraphs_workflow import DEFAULT_COLORS as _OVERLAY_COLORS

import sys as _sys
_BASE_DIR = os.path.dirname(
    _sys.executable if getattr(_sys, "frozen", False) else os.path.abspath(__file__)
)
CONFIG_PATH = os.path.join(_BASE_DIR, "studies_config.json")


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

        self.sex_field = QComboBox()
        self.sex_field.addItems(["Male", "Female", "Mixed"])
        self.sex_field.setCurrentText(data.get("sex", "Male"))
        form.addRow(QLabel("Cohort Sex:"), self.sex_field)

        self.age_field = QLineEdit(data.get("age", "16 Weeks"))
        form.addRow(QLabel("Cohort Age:"), self.age_field)

        self.bone_field = QComboBox()
        self.bone_field.addItems(["Femur", "Tibia", "Humerus"])
        self.bone_field.setCurrentText(data.get("bone", "Femur"))
        form.addRow(QLabel("Target Bone:"), self.bone_field)

        paths_meta = [
            ("raw_data_root",    "Raw Data Folder:",          True),
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

        is_mixed_init = data.get("sex", "Male") == "Mixed"
        self.group_table = QTableWidget(0, 3 if is_mixed_init else 2)
        if is_mixed_init:
            self.group_table.setHorizontalHeaderLabels(["Prefix", "Sex", "Group Name"])
            self.group_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        else:
            self.group_table.setHorizontalHeaderLabels(["Prefix", "Group Name"])
            self.group_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.group_table.setMinimumHeight(130)

        for prefix, val in data.get("group_map", {}).items():
            if isinstance(val, dict):
                self._add_group_row(prefix, val.get("group", ""), val.get("sex", "Male"))
            else:
                self._add_group_row(prefix, val)

        self.sex_field.currentTextChanged.connect(self._on_sex_changed)

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

    def _add_group_row(self, prefix="", group_name="", sex="Male"):
        row = self.group_table.rowCount()
        self.group_table.insertRow(row)
        self.group_table.setItem(row, 0, QTableWidgetItem(prefix))
        if self.group_table.columnCount() == 3:
            sex_combo = QComboBox()
            sex_combo.addItems(["Male", "Female", "Deduce from code"])
            sex_combo.setCurrentText(sex)
            self.group_table.setCellWidget(row, 1, sex_combo)
            self.group_table.setItem(row, 2, QTableWidgetItem(group_name))
        else:
            self.group_table.setItem(row, 1, QTableWidgetItem(group_name))

    def _on_sex_changed(self, sex_text):
        is_mixed = sex_text == "Mixed"
        old_is_mixed = self.group_table.columnCount() == 3
        if is_mixed == old_is_mixed:
            return

        current_rows = []
        for row in range(self.group_table.rowCount()):
            p_item = self.group_table.item(row, 0)
            prefix = p_item.text().strip() if p_item else ""
            if old_is_mixed:
                sex_w = self.group_table.cellWidget(row, 1)
                sex_val = sex_w.currentText() if sex_w else "Male"
                n_item = self.group_table.item(row, 2)
                name = n_item.text().strip() if n_item else ""
            else:
                sex_val = "Male"
                n_item = self.group_table.item(row, 1)
                name = n_item.text().strip() if n_item else ""
            current_rows.append((prefix, name, sex_val))

        self.group_table.setRowCount(0)
        if is_mixed:
            self.group_table.setColumnCount(3)
            self.group_table.setHorizontalHeaderLabels(["Prefix", "Sex", "Group Name"])
            self.group_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        else:
            self.group_table.setColumnCount(2)
            self.group_table.setHorizontalHeaderLabels(["Prefix", "Group Name"])
            self.group_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)

        for prefix, name, sex_val in current_rows:
            self._add_group_row(prefix, name, sex_val)

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
        is_mixed = self.sex_field.currentText() == "Mixed"
        for row in range(self.group_table.rowCount()):
            p = self.group_table.item(row, 0)
            prefix = p.text().strip() if p else ""
            if not prefix:
                continue
            if is_mixed:
                sex_w = self.group_table.cellWidget(row, 1)
                sex_val = sex_w.currentText() if sex_w else "Male"
                n = self.group_table.item(row, 2)
                name = n.text().strip() if n else ""
                group_map[prefix] = {"group": name, "sex": sex_val}
            else:
                n = self.group_table.item(row, 1)
                name = n.text().strip() if n else ""
                group_map[prefix] = name
        return {
            "raw_data_root":    self.path_fields["raw_data_root"].text().strip(),
            "measurement_file": self.path_fields["measurement_file"].text().strip(),
            "output_folder":    self.path_fields["output_folder"].text().strip(),
            "sex":              self.sex_field.currentText(),
            "age":              self.age_field.text().strip(),
            "bone":             self.bone_field.currentText(),
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
        self.study_selector.addItem("FKBP5 Heat")
        self.study_selector.addItem("FKBP5 New")
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
        btn_prism = QPushButton("Export CSVs → Prism")
        btn_prism.clicked.connect(self.export_csvs_to_prism)
        mgmt_layout.addWidget(btn_prism)
        self.overlay_btn = QPushButton("Overlay Graphs")
        self.overlay_btn.setCheckable(True)
        self.overlay_btn.clicked.connect(self.toggle_overlay_mode)
        mgmt_layout.addWidget(self.overlay_btn)
        self.layout.addLayout(mgmt_layout)

        self.overlay_mode = False

        # Stacked views
        self.stack = QStackedWidget()
        self.layout.addWidget(self.stack)

        self.setup_fkbp5_ui()        # stack index 0
        self.setup_single_study_ui() # stack index 1
        self.setup_overlay_ui()      # stack index 2
        self.setup_fkbp5new_ui()     # stack index 3

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
        self.study_selector.addItem("FKBP5 Heat")
        self.study_selector.addItem("FKBP5 New")
        self.study_selector.addItems(self.config.keys())
        if select_name:
            idx = self.study_selector.findText(select_name)
            if idx >= 0:
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
        if name in ("FKBP5 Heat", "FKBP5 New"):
            QMessageBox.information(self, "Not Editable",
                                    f"{name} is a built-in mode and cannot be edited here.")
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
        if name in ("FKBP5 Heat", "FKBP5 New"):
            QMessageBox.information(self, "Not Removable",
                                    f"{name} is a built-in mode and cannot be removed.")
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

        group = QGroupBox(
            "FKBP5 Heat Pipeline Configurations  "
            "(W/2W = Wildtype · M/2M = Mutant · H/2H = Heterozygous · format: W.12.F23)")
        form = QFormLayout(group)

        self.fkbp5_fields = {}
        fields_meta = [
            ("raw_dir",      "Raw Data Folder:",       FKBP5Heat_workflow.DEFAULT_RAW_DATA_ROOT,    True),
            ("tibia_master", "Tibia Master Excel File:",FKBP5Heat_workflow.DEFAULT_TIBIA_MASTER_FILE,False),
            ("femur_master", "Femur Master Excel File:",FKBP5Heat_workflow.DEFAULT_FEMUR_MASTER_FILE,False),
            ("meas_file",    "Measurements File:",      FKBP5Heat_workflow.DEFAULT_MEASUREMENT_FILE, False),
            ("csv_dir",      "CSV Export Folder:",      FKBP5Heat_workflow.DEFAULT_CSV_OUTPUT_DIR,   True),
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

        self.fkbp5_prism = QCheckBox("Export to GraphPad Prism (.pzfx) after run")
        form.addRow(QLabel(""), self.fkbp5_prism)

        layout.addWidget(group)
        self.stack.addWidget(self.fkbp5_widget)

    def setup_fkbp5new_ui(self):
        self.fkbp5new_widget = QWidget()
        layout = QVBoxLayout(self.fkbp5new_widget)

        group = QGroupBox(
            "FKBP5 New Pipeline Configurations  "
            "(W = Wildtype · Z = Heterozygous · X = Mutant   |   format: W.12.F23)")
        form = QFormLayout(group)

        self.fkbp5new_fields = {}
        fields_meta = [
            ("raw_dir",      "Raw Data Folder:",        FKBP5New_workflow.DEFAULT_RAW_DATA_ROOT,    True),
            ("tibia_master", "Tibia Master Excel File:", FKBP5New_workflow.DEFAULT_TIBIA_MASTER_FILE, False),
            ("femur_master", "Femur Master Excel File:", FKBP5New_workflow.DEFAULT_FEMUR_MASTER_FILE, False),
            ("meas_file",    "Measurements File:",       FKBP5New_workflow.DEFAULT_MEASUREMENT_FILE,  False),
            ("csv_dir",      "CSV Export Folder:",       FKBP5New_workflow.DEFAULT_CSV_OUTPUT_DIR,    True),
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
            self.fkbp5new_fields[key] = line_edit

        self.fkbp5new_fallback = QComboBox()
        self.fkbp5new_fallback.addItems(["Tibia", "Femur"])
        form.addRow(QLabel("Default Bone Fallback:"), self.fkbp5new_fallback)

        self.fkbp5new_prism = QCheckBox("Export to GraphPad Prism (.pzfx) after run")
        form.addRow(QLabel(""), self.fkbp5new_prism)

        layout.addWidget(group)
        self.stack.addWidget(self.fkbp5new_widget)

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
        self.ss_bone.addItems(["Femur", "Tibia", "Humerus"])
        form.addRow(QLabel("Target Bone Structural Type:"), self.ss_bone)

        # Raw Data Folder
        raw_layout = QHBoxLayout()
        raw_le = QLineEdit()
        raw_btn = QPushButton("Browse")
        raw_btn.clicked.connect(lambda checked, le=raw_le: self.browse_path(le, True))
        raw_layout.addWidget(raw_le)
        raw_layout.addWidget(raw_btn)
        form.addRow(QLabel("Raw Data Folder:"), raw_layout)
        self.ss_fields["raw_dir"] = raw_le

        # Measurements Excel File (browse wired to sheet refresh)
        meas_layout = QHBoxLayout()
        meas_le = QLineEdit()
        meas_btn = QPushButton("Browse")
        meas_btn.clicked.connect(self._browse_ss_meas_file)
        meas_le.editingFinished.connect(self._refresh_ss_sheet_selector)
        meas_layout.addWidget(meas_le)
        meas_layout.addWidget(meas_btn)
        form.addRow(QLabel("Measurements Excel File:"), meas_layout)
        self.ss_fields["meas_file"] = meas_le

        # Sheet selector — hidden until the file has multiple sheets
        self.ss_sheet_label = QLabel("Measurement Sheet:")
        self.ss_sheet_selector = QComboBox()
        form.addRow(self.ss_sheet_label, self.ss_sheet_selector)
        self.ss_sheet_label.setVisible(False)
        self.ss_sheet_selector.setVisible(False)

        # CSV Output Folder
        csv_layout = QHBoxLayout()
        csv_le = QLineEdit()
        csv_btn = QPushButton("Browse")
        csv_btn.clicked.connect(lambda checked, le=csv_le: self.browse_path(le, True))
        csv_layout.addWidget(csv_le)
        csv_layout.addWidget(csv_btn)
        form.addRow(QLabel("CSV/Master Output Directory:"), csv_layout)
        self.ss_fields["csv_dir"] = csv_le

        # CSV subfolder checkbox
        self.ss_auto_csvfolder = QCheckBox("Auto-create StudyName_CSVFiles subfolder")
        self.ss_auto_csvfolder.setChecked(True)
        form.addRow(QLabel(""), self.ss_auto_csvfolder)

        # Anatomical Diameters checkbox
        self.ss_anat_diam_checkbox = QCheckBox("Generate Anatomical Diameter Folders")
        form.addRow(QLabel(""), self.ss_anat_diam_checkbox)

        self.ss_prism = QCheckBox("Export to GraphPad Prism (.pzfx) after run")
        form.addRow(QLabel(""), self.ss_prism)

        layout.addWidget(group)
        self.stack.addWidget(self.single_study_widget)

    def _browse_ss_meas_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.ss_fields["meas_file"].setText(os.path.normpath(path))
            self._refresh_ss_sheet_selector()

    def _refresh_ss_sheet_selector(self):
        path = self.ss_fields["meas_file"].text().strip()
        self.ss_sheet_selector.clear()
        if not path or not os.path.exists(path):
            self.ss_sheet_label.setVisible(False)
            self.ss_sheet_selector.setVisible(False)
            return
        try:
            sheets = pd.ExcelFile(path).sheet_names
            self.ss_sheet_selector.addItems(sheets)
            multi = len(sheets) > 1
            self.ss_sheet_label.setVisible(multi)
            self.ss_sheet_selector.setVisible(multi)
        except Exception:
            self.ss_sheet_label.setVisible(False)
            self.ss_sheet_selector.setVisible(False)

    def browse_path(self, line_edit, is_directory):
        if is_directory:
            path = QFileDialog.getExistingDirectory(self, "Select Directory")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select File", "", "Excel Files (*.xlsx *.xls)")
        if path:
            line_edit.setText(os.path.normpath(path))

    # ------------------------------------------------------------------
    # Overlay Graphs UI
    # ------------------------------------------------------------------

    def setup_overlay_ui(self):
        self.overlay_widget = QWidget()
        outer = QVBoxLayout(self.overlay_widget)

        config_group = QGroupBox("Overlay Configuration  —  study selected above auto-fills the scan folder")
        form = QFormLayout(config_group)

        # Scan folder — auto-filled from the main study selector; user can also browse freely
        scan_row = QHBoxLayout()
        self.ov_scan_folder = QLineEdit()
        scan_browse_btn = QPushButton("Browse")
        scan_browse_btn.clicked.connect(lambda: self.browse_path(self.ov_scan_folder, True))
        scan_trigger_btn = QPushButton("Scan .txt Files")
        scan_trigger_btn.clicked.connect(self._scan_overlay_files)
        scan_row.addWidget(self.ov_scan_folder)
        scan_row.addWidget(scan_browse_btn)
        scan_row.addWidget(scan_trigger_btn)
        form.addRow(QLabel("Scan Folder:"), scan_row)

        # Graph title
        self.ov_title = QLineEdit()
        self.ov_title.setPlaceholderText("e.g. FKBP5 WT vs KO — Left Femur")
        form.addRow(QLabel("Graph Title:"), self.ov_title)

        # Save folder
        save_row = QHBoxLayout()
        self.ov_save_folder = QLineEdit()
        save_browse_btn = QPushButton("Browse")
        save_browse_btn.clicked.connect(lambda: self.browse_path(self.ov_save_folder, True))
        save_row.addWidget(self.ov_save_folder)
        save_row.addWidget(save_browse_btn)
        form.addRow(QLabel("Save Folder:"), save_row)

        outer.addWidget(config_group)

        # File table
        files_group = QGroupBox("Files  —  check to include, set label and color per file")
        files_layout = QVBoxLayout(files_group)

        # Search / filter row
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Search:"))
        self.ov_search = QLineEdit()
        self.ov_search.setPlaceholderText("filter by filename or label…")
        self.ov_search.textChanged.connect(self._filter_overlay_table)
        search_row.addWidget(self.ov_search)
        files_layout.addLayout(search_row)

        self.ov_table = QTableWidget(0, 4)
        self.ov_table.setHorizontalHeaderLabels(["✓", "File", "Mouse Code / Label", "Color"])
        self.ov_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.ov_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.ov_table.setColumnWidth(0, 36)
        self.ov_table.setColumnWidth(3, 90)
        self.ov_table.setMinimumHeight(200)
        files_layout.addWidget(self.ov_table)

        sel_row = QHBoxLayout()
        sel_all_btn = QPushButton("Select All")
        sel_all_btn.clicked.connect(lambda: self._ov_set_all_checked(True))
        sel_none_btn = QPushButton("Select None")
        sel_none_btn.clicked.connect(lambda: self._ov_set_all_checked(False))
        sel_row.addWidget(sel_all_btn)
        sel_row.addWidget(sel_none_btn)
        sel_row.addStretch()
        files_layout.addLayout(sel_row)

        outer.addWidget(files_group)
        self.stack.addWidget(self.overlay_widget)

    def export_csvs_to_prism(self):
        study = self.study_selector.currentText()

        if self.overlay_mode or study == "Overlay Graphs":
            self.show_message_box("Not Applicable",
                                  "Select a study or FKBP5 mode first, then export.",
                                  is_error=True)
            return

        if study == "FKBP5 Heat":
            csv_dir    = self.fkbp5_fields["csv_dir"].text().strip()
            study_name = "FKBP5Heat"
        elif study == "FKBP5 New":
            csv_dir    = self.fkbp5new_fields["csv_dir"].text().strip()
            study_name = "FKBP5New"
        else:
            csv_dir = self.ss_fields["csv_dir"].text().strip()
            if self.ss_auto_csvfolder.isChecked():
                csv_dir = os.path.join(csv_dir, f"{study}_CSVFiles")
            study_name = study

        if not csv_dir or not os.path.isdir(csv_dir):
            self.show_message_box("Folder Not Found",
                                  f"CSV Export Folder does not exist:\n{csv_dir}",
                                  is_error=True)
            return

        created = prism_export.workflow_output_to_pzfx(csv_dir, study_name)

        if created:
            self.show_message_box(
                "Prism Export Complete",
                f"Created {len(created)} file(s):\n\n" + "\n".join(created))
        else:
            self.show_message_box(
                "Nothing Exported",
                f"No *_Analysis_By_Genotype subfolders with CSVs found in:\n{csv_dir}",
                is_error=True)

    def _sync_overlay_scan_folder(self):
        """Keep the overlay scan folder in sync with the main study selector."""
        study = self.study_selector.currentText()
        if study == "FKBP5 Heat":
            folder = FKBP5Heat_workflow.DEFAULT_RAW_DATA_ROOT
        else:
            folder = self.config.get(study, {}).get("raw_data_root", "")
        self.ov_scan_folder.setText(folder)

    def toggle_overlay_mode(self):
        self.overlay_mode = self.overlay_btn.isChecked()
        if self.overlay_mode:
            self.stack.setCurrentIndex(2)
            self._sync_overlay_scan_folder()
            self.run_button.setText("Generate Overlay")
        else:
            self.run_button.setText("Execute Workflow")
            self.toggle_ui()

    def _scan_overlay_files(self):
        folder = self.ov_scan_folder.text().strip()
        if not folder or not os.path.isdir(folder):
            QMessageBox.warning(self, "Scan Error", "Please set a valid scan folder first.")
            return

        # Recursive scan — finds .txt files at any depth (top-level and all subfolders)
        all_files = sorted(glob.glob(os.path.join(folder, "**", "*.txt"), recursive=True))

        if not all_files:
            QMessageBox.information(self, "No Files Found",
                                    f"No .txt files found in:\n{folder}")
            return

        self.ov_search.clear()
        self.ov_table.setRowCount(0)
        for full_path in all_files:
            rel_path = os.path.relpath(full_path, folder)
            stem     = os.path.splitext(os.path.basename(full_path))[0]
            row      = self.ov_table.rowCount()
            self.ov_table.insertRow(row)

            # Unchecked by default — user picks which files to include
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk_item.setCheckState(Qt.CheckState.Unchecked)
            self.ov_table.setItem(row, 0, chk_item)

            # Relative path displayed; full path stored in UserRole
            path_item = QTableWidgetItem(rel_path)
            path_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            path_item.setData(Qt.ItemDataRole.UserRole, full_path)
            self.ov_table.setItem(row, 1, path_item)

            self.ov_table.setItem(row, 2, QTableWidgetItem(stem))

            # Blank color button — user must choose a color before including this file
            color_btn = QPushButton("Pick")
            color_btn.setFixedHeight(24)
            color_btn.setProperty("hex_color", "")
            color_btn.setStyleSheet(
                "background-color: #e8e8e8; border: 1px dashed #aaa; "
                "border-radius: 3px; color: #777; font-size: 10px;")
            color_btn.clicked.connect(lambda checked, b=color_btn: self._ov_pick_color(b))
            self.ov_table.setCellWidget(row, 3, color_btn)

        self.ov_table.resizeRowsToContents()

    def _ov_pick_color(self, btn):
        current_hex = btn.property("hex_color") or "#1f77b4"
        color = QColorDialog.getColor(QColor(current_hex), self, "Pick Color")
        if color.isValid():
            hex_color = color.name()
            btn.setProperty("hex_color", hex_color)
            btn.setText("")
            btn.setStyleSheet(
                f"background-color: {hex_color}; border: 1px solid #888; border-radius: 3px;")

    def _filter_overlay_table(self, text):
        text = text.strip().lower()
        for row in range(self.ov_table.rowCount()):
            path_item  = self.ov_table.item(row, 1)
            label_item = self.ov_table.item(row, 2)
            fname = (path_item.text()  if path_item  else "").lower()
            label = (label_item.text() if label_item else "").lower()
            self.ov_table.setRowHidden(row, bool(text) and text not in fname and text not in label)

    def _ov_set_all_checked(self, checked):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for row in range(self.ov_table.rowCount()):
            item = self.ov_table.item(row, 0)
            if item:
                item.setCheckState(state)

    def toggle_ui(self):
        if self.overlay_mode:
            self.stack.setCurrentIndex(2)
            self._sync_overlay_scan_folder()
            return
        current_selection = self.study_selector.currentText()
        if current_selection == "FKBP5 Heat":
            self.stack.setCurrentIndex(0)
        elif current_selection == "FKBP5 New":
            self.stack.setCurrentIndex(3)
        else:
            self.stack.setCurrentIndex(1)
            study_data = self.config.get(current_selection, {})
            self.ss_fields["raw_dir"].setText(
                study_data.get("raw_data_root", ""))
            self.ss_fields["meas_file"].setText(
                study_data.get("measurement_file", ""))
            self.ss_fields["csv_dir"].setText(
                study_data.get("output_folder", ""))
            if study_data.get("sex"):
                self.ss_sex.setCurrentText(study_data["sex"])
            if study_data.get("age"):
                self.ss_age.setText(study_data["age"])
            if study_data.get("bone"):
                self.ss_bone.setCurrentText(study_data["bone"])
            self._refresh_ss_sheet_selector()

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
            if self.overlay_mode:
                save_folder = self.ov_save_folder.text().strip()
                scan_folder = self.ov_scan_folder.text().strip()
                if not save_folder:
                    self.show_message_box("Validation Error",
                                          "Please set a Save Folder before running.", is_error=True)
                    return
                if not scan_folder:
                    self.show_message_box("Validation Error",
                                          "Please set a Scan Folder and scan for files first.", is_error=True)
                    return
                title = self.ov_title.text().strip() or "Overlay Comparison"
                files_data = []
                for row in range(self.ov_table.rowCount()):
                    chk = self.ov_table.item(row, 0)
                    if not chk or chk.checkState() != Qt.CheckState.Checked:
                        continue
                    path_item   = self.ov_table.item(row, 1)
                    label_item  = self.ov_table.item(row, 2)
                    color_btn   = self.ov_table.cellWidget(row, 3)
                    # Full path stored in UserRole; fall back to joining scan_folder + display text
                    full_path = (path_item.data(Qt.ItemDataRole.UserRole)
                                 if path_item else None)
                    if not full_path and path_item:
                        full_path = os.path.join(scan_folder, path_item.text())
                    label = (label_item.text().strip() if label_item else "") or os.path.basename(full_path or "")
                    color = (color_btn.property("hex_color") if color_btn else "") or \
                            _OVERLAY_COLORS[len(files_data) % len(_OVERLAY_COLORS)]
                    if full_path:
                        files_data.append((full_path, label, color))
                if not files_data:
                    self.show_message_box("No Files Selected",
                                          "Scan a folder and check at least one file.", is_error=True)
                    return
                output_path = OverlayGraphs_workflow.run_overlay(files_data, title, save_folder)
                self.show_message_box("Success", f"Overlay graph saved to:\n{output_path}")

            elif current_study == "FKBP5 Heat":
                inputs = {
                    "data_folder":       self.fkbp5_fields["raw_dir"].text(),
                    "tibia_master":      self.fkbp5_fields["tibia_master"].text(),
                    "femur_master":      self.fkbp5_fields["femur_master"].text(),
                    "measurement_path":  self.fkbp5_fields["meas_file"].text(),
                    "csv_out_dir":       self.fkbp5_fields["csv_dir"].text(),
                    "structure_type":    "Split Genders Male/Female",
                    "fallback_bone":     self.fkbp5_fallback.currentText(),
                }
                success = FKBP5Heat_workflow.run_workflow(inputs)
                if success:
                    msg = "FKBP5 Heat workflow completed and master files updated."
                    if self.fkbp5_prism.isChecked():
                        created = prism_export.workflow_output_to_pzfx(
                            inputs["csv_out_dir"], "FKBP5Heat")
                        if created:
                            msg += f"\n\nPrism files saved:\n" + "\n".join(created)
                    self.show_message_box("Success", msg)
                else:
                    self.show_message_box(
                        "Execution Incomplete",
                        "Check data directory contents.", is_error=True)

            elif current_study == "FKBP5 New":
                inputs = {
                    "data_folder":       self.fkbp5new_fields["raw_dir"].text(),
                    "tibia_master":      self.fkbp5new_fields["tibia_master"].text(),
                    "femur_master":      self.fkbp5new_fields["femur_master"].text(),
                    "measurement_path":  self.fkbp5new_fields["meas_file"].text(),
                    "csv_out_dir":       self.fkbp5new_fields["csv_dir"].text(),
                    "structure_type":    "Split Genders Male/Female",
                    "fallback_bone":     self.fkbp5new_fallback.currentText(),
                }
                success = FKBP5New_workflow.run_workflow(inputs)
                if success:
                    msg = "FKBP5 New workflow completed and master files updated."
                    if self.fkbp5new_prism.isChecked():
                        created = prism_export.workflow_output_to_pzfx(
                            inputs["csv_out_dir"], "FKBP5New")
                        if created:
                            msg += f"\n\nPrism files saved:\n" + "\n".join(created)
                    self.show_message_box("Success", msg)
                else:
                    self.show_message_box(
                        "Execution Incomplete",
                        "Check data directory contents.", is_error=True)

            else:
                study_data = self.config.get(current_study, {})
                meas_path = self.ss_fields["meas_file"].text()
                inputs = {
                    "study_name":           current_study,
                    "data_folder":          self.ss_fields["raw_dir"].text(),
                    "measurement_path":     meas_path,
                    "csv_out_dir":          self.ss_fields["csv_dir"].text(),
                    "sex":                  self.ss_sex.currentText(),
                    "age":                  self.ss_age.text(),
                    "bone":                 self.ss_bone.currentText(),
                    "group_map":            study_data.get("group_map", {}),
                    "anatomical_diameters": self.ss_anat_diam_checkbox.isChecked(),
                    "measurement_sheet":    self.ss_sheet_selector.currentText() or None,
                    "auto_csv_subfolder":   self.ss_auto_csvfolder.isChecked(),
                }
                success = SingleStudy_workflow.run_workflow(inputs)
                if success:
                    msg = f"Custom study pipeline executed successfully for: {current_study}"
                    if self.ss_prism.isChecked():
                        csv_dir = inputs["csv_out_dir"]
                        if self.ss_auto_csvfolder.isChecked():
                            csv_dir = os.path.join(
                                csv_dir, f"{current_study}_CSVFiles")
                        created = prism_export.workflow_output_to_pzfx(
                            csv_dir, current_study)
                        if created:
                            msg += f"\n\nPrism files saved:\n" + "\n".join(created)
                    self.show_message_box("Success", msg)
                else:
                    self.show_message_box(
                        "Execution Incomplete",
                        "Check data directory contents.", is_error=True)

        except Exception as e:
            self.show_message_box(
                "Execution Failure", "An unhandled error occurred:", str(e), is_error=True)
        finally:
            self.run_button.setEnabled(True)
            self.run_button.setText("Generate Overlay" if self.overlay_mode else "Execute Workflow")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainDashboard()
    window.show()
    sys.exit(app.exec())
