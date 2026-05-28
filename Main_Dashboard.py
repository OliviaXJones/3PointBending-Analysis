import sys
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QComboBox, QPushButton, QStackedWidget, QMessageBox)

# Import the backend logic as modules
import fkbp5_pipeline
import single_study_pipeline


class MainDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("3-Point Bending Workflow Hub")
        self.config = self.load_config()

        # Main Layout
        self.central_widget = QWidget()
        self.layout = QVBoxLayout(self.central_widget)

        # Study Selector
        self.study_selector = QComboBox()
        self.study_selector.addItem("FKBP5 Genotyping")
        self.study_selector.addItems(self.config.keys())
        self.study_selector.currentIndexChanged.connect(self.toggle_ui)

        self.layout.addWidget(self.study_selector)

        # Stacked Widget for UI Switching
        self.stack = QStackedWidget()
        self.layout.addWidget(self.stack)

        self.run_button = QPushButton("Execute Workflow")
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

    def toggle_ui(self):
        # Logic to swap visibility based on FKBP5 vs Single Study selection
        is_fkbp5 = self.study_selector.currentText() == "FKBP5 Genotyping"
        # Hide/Show inputs here using setVisible(not is_fkbp5)

    def run_pipeline(self):
        if self.study_selector.currentText() == "FKBP5 Genotyping":
            fkbp5_pipeline.run_workflow(self.get_inputs())
        else:
            single_study_pipeline.run_workflow(self.get_inputs())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainDashboard()
    window.show()
    sys.exit(app.exec())
