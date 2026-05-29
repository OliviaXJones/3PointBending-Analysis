import sys
import json
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QPushButton, QStackedWidget, QMessageBox, QLabel,
    QLineEdit, QFileDialog, QFormLayout, QGroupBox
)
from PyQt6.QtCore import Qt

# Import the backend modules
import FKBP5_workflow
import SingleStudy_workflow


class MainDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("3-Point Bending Workflow Hub")
        self.setMinimumWidth(650)
        self.config = self.load_config()

        # Main Layout
        self.central_widget = QWidget()
        self.layout = QVBoxLayout(self.central_widget)

        # Study Selector Row
        selector_layout = QHBoxLayout()
        selector_layout.addWidget(QLabel("Select Analysis Mode or Study:"))
        self.study_selector = QComboBox()
        self.study_selector.addItem("FKBP5 Genotyping")
        self.study_selector.addItems(self.config.keys())
        self.study_selector.currentIndexChanged.connect(self.toggle_ui)
        selector_layout.addWidget(self.study_selector, 1)
        self.layout.addLayout(selector_layout)

        # Stacked Widget for UI Switching
        self.stack = QStackedWidget()
        self.layout.addWidget(self.stack)

        # Build individual view layouts
        self.setup_fkbp5_ui()
        self.setup_single_study_ui()

        # Run Button
        self.run_button = QPushButton("Execute Workflow")
        self.run_button.setMinimumHeight(40)
        self.run_button.clicked.connect(self.run_pipeline)
        self.layout.addWidget(self.run_button)

        self.setCentralWidget(self.central_widget)
        self.toggle_ui()

    def load_config(self):
        try:
            with open("studies_config.json", "r") as f:
                return json.load(f)
        except:
            return {}

    def setup_fkbp5_ui(self):
        """Creates the form fields required for the FKBP5 Workflow."""
        self.fkbp5_widget = QWidget()
        layout = QVBoxLayout(self.fkbp5_widget)

        group = QGroupBox("FKBP5 Genotyping Pipeline Configurations")
        form = QFormLayout(group)

        self.fkbp5_fields = {}
        fields_meta = [
            ("raw_dir", "Raw Data Folder:",
             FKBP5_workflow.DEFAULT_RAW_DATA_ROOT, True),
            ("tibia_master", "Tibia Master Excel File:",
             FKBP5_workflow.DEFAULT_TIBIA_MASTER_FILE, False),
            ("femur_master", "Femur Master Excel File:",
             FKBP5_workflow.DEFAULT_FEMUR_MASTER_FILE, False),
            ("meas_file", "Measurements File:",
             FKBP5_workflow.DEFAULT_MEASUREMENT_FILE, False),
            ("csv_dir", "CSV Export Folder:",
             FKBP5_workflow.DEFAULT_CSV_OUTPUT_DIR, True),
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

        """ Structure and Fallback Dropdowns
        self.fkbp5_arch = QComboBox()
        self.fkbp5_arch.addItems(
            ["Single Table", "Split Genders (Male/Female)"])
        form.addRow(QLabel("Spreadsheet Architecture:"), self.fkbp5_arch)"""

        self.fkbp5_fallback = QComboBox()
        self.fkbp5_fallback.addItems(["Tibia", "Femur"])
        form.addRow(QLabel("Default Bone Fallback:"), self.fkbp5_fallback)

        layout.addWidget(group)
        self.stack.addWidget(self.fkbp5_widget)

    def setup_single_study_ui(self):
        """Creates the form fields required for custom single study configurations."""
        self.single_study_widget = QWidget()
        layout = QVBoxLayout(self.single_study_widget)

        group = QGroupBox("Custom Study Template Parameters")
        form = QFormLayout(group)

        self.ss_fields = {}

        # Meta Input Elements to build [StudyName]_[Sex]_[Age]_[Bone].xlsx
        self.ss_sex = QComboBox()
        self.ss_sex.addItems(["Male", "Female", "Mixed"])
        form.addRow(QLabel("Cohort Sex:"), self.ss_sex)

        self.ss_age = QLineEdit("16 Weeks")
        form.addRow(QLabel("Cohort Age:"), self.ss_age)

        self.ss_bone = QComboBox()
        self.ss_bone.addItems(["Femur", "Tibia"])
        form.addRow(QLabel("Target Bone Structural Type:"), self.ss_bone)

        # Standard file path selections
        paths_meta = [
            ("raw_dir", "Raw Data Folder:", True),
            ("meas_file", "Measurements Excel File:", False),
            ("csv_dir", "CSV/Master Output Directory:", True),
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
            # Pre-load JSON metadata values if selecting a specific study entry
            study_data = self.config.get(current_selection, {})
            self.ss_fields["raw_dir"].setText(
                study_data.get("raw_data_root", ""))
            self.ss_fields["meas_file"].setText(
                study_data.get("measurement_file", ""))

            # FIXED: Matches the exact key in studies_config.json
            self.ss_fields["csv_dir"].setText(
                study_data.get("output_folder", ""))

    def show_message_box(self, title, message, details="", is_error=False):
        """Helper ensuring safe UI text contrast colors on error and success messages."""
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

    def run_pipeline(self):
        current_study = self.study_selector.currentText()
        self.run_button.setEnabled(False)
        self.run_button.setText("Processing...")
        QApplication.processEvents()

        try:
            if current_study == "FKBP5 Genotyping":
                # Hardcode the architecture choice here so the backend receives it automatically
                inputs = {
                    "data_folder": self.fkbp5_fields["raw_dir"].text(),
                    "tibia_master": self.fkbp5_fields["tibia_master"].text(),
                    "femur_master": self.fkbp5_fields["femur_master"].text(),
                    "measurement_path": self.fkbp5_fields["meas_file"].text(),
                    "csv_out_dir": self.fkbp5_fields["csv_dir"].text(),
                    "structure_type": "Split Genders Male/Female",  # Hardcoded architecture
                    "fallback_bone": self.fkbp5_fallback.currentText()
                }

                success = FKBP5_workflow.run_workflow(inputs)
                if success:
                    self.show_message_box(
                        "Success", "FKBP5 calculation finalized and master files successfully updated.")
                else:
                    self.show_message_box(
                        "Execution Incomplete", "Check data directory contents.", is_error=True)
            else:
                # Compile dynamic properties for single study logic
                study_data = self.config.get(current_study, {})

                meas_path = self.ss_fields["meas_file"].text()
                # Force master to be in the same dir as measurement file
                master_dir = os.path.dirname(meas_path)
                master_filename = os.path.basename(
                    study_data.get("master_file", "Default_Master.xlsx"))
                inferred_master_path = os.path.join(
                    master_dir, master_filename)
                inputs = {
                    "study_name": current_study,
                    "data_folder": self.ss_fields["raw_dir"].text(),
                    "measurement_path": self.ss_fields["meas_file"].text(),
                    "csv_out_dir": self.ss_fields["csv_dir"].text(),
                    "sex": self.ss_sex.currentText(),
                    "age": self.ss_age.text(),
                    "bone": self.ss_bone.currentText(),
                    "group_map": study_data.get("group_map", {})
                }

                success = SingleStudy_workflow.run_workflow(inputs)
                if success:
                    self.show_message_box(
                        "Success", f"Custom study pipeline executed successfully for: {current_study}")
                else:
                    self.show_message_box(
                        "Execution Incomplete", "Check data directory contents.", is_error=True)

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
