import json
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QInputDialog, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QTabWidget, QComboBox
)
import os
from loading_overlay import LoadingOverlay

class SettingsTab(QWidget):
    def __init__(self, config_file_path):
        super().__init__()
        self.config_file_path = config_file_path
        self.columns = []  # Store extracted column names
        self.config_data = self.load_config()
        
        # Initialize dropdown variables
        self.id_col_dropdown = QComboBox()
        self.match_col_dropdown = QComboBox()
        self.amount_col_dropdown = QComboBox()
        self.settle_col_dropdown = QComboBox()
        self.date_col_dropdown = QComboBox()
        self.comment_col_dropdown = QComboBox()

        self.init_ui()
        
        # Add loading overlay
        self.loading_overlay = LoadingOverlay(self)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.loading_overlay.setFixedSize(self.size())

    def load_config(self):
        # Load configuration from config.json
        if os.path.exists(self.config_file_path):
            with open(self.config_file_path, "r") as file:
                return json.load(file)
        return {}

    def save_config(self):
        # Save the updated configuration back to config.json
        with open(self.config_file_path, "w") as file:
            json.dump(self.config_data, file, indent=4)

    def init_ui(self):
        layout = QVBoxLayout()

        self.tab_widget = QTabWidget()

        # Subtabs
        self.add_app_tab = self.create_add_app_tab()
        self.remove_app_tab = self.create_remove_app_tab()

        self.tab_widget.addTab(self.add_app_tab, "Add App")
        self.tab_widget.addTab(self.remove_app_tab, "Remove App")

        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

    def create_add_app_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(10)  # Add spacing between widgets
        layout.setContentsMargins(20, 20, 20, 20)  # Add margins around the layout

        # File upload section
        upload_section = QWidget()
        upload_layout = QVBoxLayout()
        upload_layout.setSpacing(5)
        
        title_label = QLabel("Upload Settlement Report")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        upload_layout.addWidget(title_label)
        
        self.upload_button = QPushButton("Choose Excel File")
        self.upload_button.setMinimumHeight(40)
        self.upload_button.clicked.connect(self.upload_file)
        upload_layout.addWidget(self.upload_button)
        
        upload_section.setLayout(upload_layout)
        layout.addWidget(upload_section)

        # Column selection section
        column_section = QWidget()
        column_layout = QVBoxLayout()
        column_layout.setSpacing(5)
        
        column_title = QLabel("Column Mapping")
        column_title.setStyleSheet("font-size: 14px; font-weight: bold; margin-top: 10px;")
        column_layout.addWidget(column_title)

        # Create dropdowns with labels
        dropdowns = [
            ("ID Column(in App settlment report)", self.id_col_dropdown),
            ("Match Column(in AFC-Triffi)", self.match_col_dropdown),
            ("Amount Column", self.amount_col_dropdown),
            ("Settle Column", self.settle_col_dropdown),
            ("Date Column", self.date_col_dropdown),
            ("Comment/Status Column", self.comment_col_dropdown)
        ]

        for label_text, dropdown in dropdowns:
            label = QLabel(label_text)
            label.setStyleSheet("margin-top: 5px;")
            dropdown.setMinimumHeight(30)
            column_layout.addWidget(label)
            column_layout.addWidget(dropdown)

        column_section.setLayout(column_layout)
        layout.addWidget(column_section)

        # Save button section
        self.save_button = QPushButton("Add App to Config")
        self.save_button.setMinimumHeight(40)
        self.save_button.setEnabled(False)
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.save_button.clicked.connect(self.add_app_to_config)
        layout.addWidget(self.save_button)

        layout.addStretch()  # Add stretch to push everything to the top
        tab.setLayout(layout)
        return tab

    def create_remove_app_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        title_label = QLabel("Remove App Configuration")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(title_label)

        # App selection section
        select_label = QLabel("Select App to Remove")
        layout.addWidget(select_label)
        
        self.app_dropdown = QComboBox()
        self.app_dropdown.setMinimumHeight(30)
        self.update_app_dropdown()
        layout.addWidget(self.app_dropdown)

        # Remove button
        self.remove_button = QPushButton("Remove Selected App")
        self.remove_button.setMinimumHeight(40)
        self.remove_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border-radius: 5px;
                font-weight: bold;
                margin-top: 10px;
            }
        """)
        self.remove_button.clicked.connect(self.remove_app)
        layout.addWidget(self.remove_button)

        layout.addStretch()  # Add stretch to push everything to the top
        tab.setLayout(layout)
        return tab

    def upload_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            # Show loading overlay
            self.loading_overlay.start_loading("Loading file...")
            
            try:
                # Load the Excel file
                df = pd.read_excel(file_path)
                self.loading_overlay.set_progress(50)
                
                self.columns = df.columns.tolist()

                # Populate dropdowns with column names
                for dropdown in [self.id_col_dropdown,
                                 self.amount_col_dropdown, self.settle_col_dropdown,
                                 self.date_col_dropdown, self.comment_col_dropdown]:
                    dropdown.clear()
                    dropdown.addItems(self.columns)
                
                # Set predefined values for match_col_dropdown
                self.match_col_dropdown.clear()
                self.match_col_dropdown.addItems(['TicketNUmber', 'order_id', 'transaction_ref_no'])

                self.loading_overlay.set_progress(80)

                # Enable the save button
                self.save_button.setEnabled(True)

                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Success", "File uploaded successfully. Select columns.")
                
            except Exception as e:
                self.loading_overlay.stop_loading()
                QMessageBox.critical(self, "Error", f"Error loading file: {e}")

    def add_app_to_config(self):
        app_name, ok = QInputDialog.getText(self, "App Name", "Enter the name of the app:")
        if not ok or not app_name.strip():
            QMessageBox.warning(self, "Error", "App name cannot be empty.")
            return

        # Show loading overlay
        self.loading_overlay.start_loading("Saving configuration...")

        try:
            app_config = {
                'id_col': self.id_col_dropdown.currentText(),
                'match_col': self.match_col_dropdown.currentText(),
                'amount_col': self.amount_col_dropdown.currentText(),
                'settle_col': self.settle_col_dropdown.currentText(),
                'date_col': self.date_col_dropdown.currentText(),
                'comment_col': self.comment_col_dropdown.currentText(),
            }

            self.loading_overlay.set_progress(50)
            
            self.config_data[app_name] = app_config
            self.save_config()
            self.update_app_dropdown()
            
            self.loading_overlay.set_progress(100)
            self.loading_overlay.stop_loading()
            QMessageBox.information(self, "Success", f"Configuration for '{app_name}' added successfully.")
            
        except Exception as e:
            self.loading_overlay.stop_loading()
            QMessageBox.critical(self, "Error", f"Error saving configuration: {e}")

    def update_app_dropdown(self):
        self.app_dropdown.clear()
        self.app_dropdown.addItems(self.config_data.keys())

    def remove_app(self):
        app_name = self.app_dropdown.currentText()
        if not app_name:
            QMessageBox.warning(self, "Error", "No app selected.")
            return

        confirm = QMessageBox.question(self, "Confirm Deletion",
                                        f"Are you sure you want to delete the configuration for '{app_name}'?",
                                        QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            # Show loading overlay
            self.loading_overlay.start_loading("Removing app configuration...")
            
            try:
                self.config_data.pop(app_name, None)
                self.loading_overlay.set_progress(50)
                
                self.save_config()
                self.update_app_dropdown()
                
                self.loading_overlay.set_progress(100)
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Success", f"Configuration for '{app_name}' removed successfully.")
                
            except Exception as e:
                self.loading_overlay.stop_loading()
                QMessageBox.critical(self, "Error", f"Error removing configuration: {e}")

