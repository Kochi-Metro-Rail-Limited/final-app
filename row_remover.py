import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QTableView, QCheckBox, QScrollArea
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from loading_overlay import LoadingOverlay


class ConsolidateUploader(QWidget):
    def __init__(self):
        super().__init__()

        # Styling
        drop_zone_style = """
            QLabel {
                background-color: white;
                border: 2px dashed #aaaaaa;
                border-radius: 8px;
                padding: 20px;
                color: #666666;
                font-size: 14px;
                min-height: 100px;
            }
            QLabel:hover {
                background-color: #f8f8f8;
                border-color: #666666;
            }
        """

        button_style = """
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
                font-size: 14px;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """

        options_style = """
            QLabel {
                color: #333333;
                font-size: 14px;
                font-weight: bold;
                margin-top: 15px;
            }
            QCheckBox {
                font-size: 13px;
                padding: 5px;
                spacing: 8px;
            }
            QCheckBox:hover {
                background-color: #f0f0f0;
                border-radius: 4px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator:unchecked {
                border: 2px solid #999999;
                border-radius: 3px;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                border: 2px solid #4CAF50;
                border-radius: 3px;
                background-color: #4CAF50;
            }
        """

        table_style = """
            QTableView {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 4px;
                gridline-color: #e0e0e0;
                selection-background-color: #e8f5e9;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 8px;
                border: none;
                border-right: 1px solid #cccccc;
                border-bottom: 1px solid #cccccc;
                font-weight: bold;
            }
            QTableView::item {
                padding: 5px;
            }
            QTableView::item:hover {
                background-color: #f8f8f8;
            }
        """

        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(15)
        self.main_layout.setContentsMargins(20, 20, 20, 20)

        # Header section
        header_label = QLabel("File Consolidation")
        header_label.setStyleSheet("""
            font-size: 18px;
            font-weight: bold;
            color: #333333;
            margin-bottom: 10px;
        """)
        self.main_layout.addWidget(header_label)

        # File Upload Area with icon and better instructions
        self.file_label = QLabel(" Excel file to consolidate\n click to browse")
        self.file_label.setStyleSheet(drop_zone_style)
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setCursor(Qt.PointingHandCursor)
        self.file_label.mousePressEvent = self.upload_file
        self.file_path = None
        self.main_layout.addWidget(self.file_label)

        # Options section with better organization
        self.options_container = QWidget()
        self.options_container.setStyleSheet(options_style)
        options_layout = QVBoxLayout(self.options_container)
        
        self.options_label = QLabel("Select options to remove:")
        options_layout.addWidget(self.options_label)

        # Checkbox container with scrolling support
        self.checkbox_container = QWidget()
        self.options_layout = QVBoxLayout(self.checkbox_container)
        self.options_layout.setSpacing(5)
        self.options_layout.setAlignment(Qt.AlignTop)
        
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.checkbox_container)
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
        """)
        
        options_layout.addWidget(scroll_area)
        self.options_container.hide()
        self.main_layout.addWidget(self.options_container)
        
        self.checkboxes = {}

        # Table view with improved styling
        self.file_table = QTableView()
        self.file_table.setStyleSheet(table_style)
        self.file_table.horizontalHeader().setStretchLastSection(True)
        self.file_table.setAlternatingRowColors(True)
        self.main_layout.addWidget(self.file_table)

        # Process Button with icon
        self.process_button = QPushButton("💾 Process and Download")
        self.process_button.setStyleSheet(button_style)
        self.process_button.clicked.connect(self.process_file)
        self.process_button.hide()
        
        # Center the button
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.addStretch()
        button_layout.addWidget(self.process_button)
        button_layout.addStretch()
        
        self.main_layout.addWidget(button_container)

        self.setLayout(self.main_layout)

        # Add loading overlay
        self.loading_overlay = LoadingOverlay(self)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.loading_overlay.setFixedSize(self.size())

    def upload_file(self, event):
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if self.file_path:
            # Show loading overlay
            self.loading_overlay.start_loading("Loading file...")
            
            try:
                # Read all sheets from the Excel file
                excel_file = pd.ExcelFile(self.file_path)
                self.loading_overlay.set_progress(30)
                
                combined_df = pd.DataFrame()  # Initialize empty DataFrame
                
                # Iterate through all sheets
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                    # Only include sheets that have an 'Action' column
                    if 'Action' in df.columns:
                        combined_df = pd.concat([combined_df, df], ignore_index=True)
                
                self.loading_overlay.set_progress(60)
                
                if combined_df.empty:
                    self.loading_overlay.stop_loading()
                    QMessageBox.warning(self, "Error", "No sheets found with an 'Action' column.")
                    return
                    
                # Update file label
                self.file_label.setText(f"{self.file_path.split('/')[-1]}")
                
                # Clear existing checkboxes
                for checkbox in self.checkboxes.values():
                    self.options_layout.removeWidget(checkbox)
                    checkbox.deleteLater()
                self.checkboxes.clear()

                # Get unique options from the Action column
                unique_options = combined_df['Action'].unique()
                unique_options = [opt for opt in unique_options if pd.notna(opt) and opt != '']

                if not unique_options:
                    self.loading_overlay.stop_loading()
                    QMessageBox.warning(self, "Error", "No valid options found in the Action column.")
                    return

                # Create new checkboxes for each unique option
                for option in unique_options:
                    checkbox = QCheckBox(str(option))
                    self.checkboxes[str(option)] = checkbox
                    self.options_layout.addWidget(checkbox)

                self.loading_overlay.set_progress(80)

                # Show the options section and process button
                self.options_label.show()
                self.options_container.show()
                self.process_button.show()

                # Load the combined table
                self.load_table(self.file_table, combined_df)
                
                self.loading_overlay.set_progress(100)
                self.loading_overlay.stop_loading()

            except Exception as e:
                self.loading_overlay.stop_loading()
                QMessageBox.critical(self, "Error", f"Error loading file:\n{str(e)}")

    def process_file(self):
        if not self.file_path:
            QMessageBox.warning(self, "Error", "Please upload a file first.")
            return

        # Get selected options to remove
        options_to_remove = [option for option, checkbox in self.checkboxes.items() if checkbox.isChecked()]
        
        if not options_to_remove:
            QMessageBox.warning(self, "Error", "Please select at least one option to remove.")
            return

        # Disable process button
        self.process_button.setEnabled(False)
        
        # Show loading overlay
        self.loading_overlay.start_loading("Processing file...")

        try:
            # Read all sheets and combine those with Action column
            excel_file = pd.ExcelFile(self.file_path)
            combined_df = pd.DataFrame()
            self.loading_overlay.set_progress(30)
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                if 'Action' in df.columns:
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
            
            # Remove rows where Action is in selected options
            df_filtered = combined_df[~combined_df['Action'].isin(options_to_remove)]
            self.loading_overlay.set_progress(60)
            
            # Get save location from user
            save_path, _ = QFileDialog.getSaveFileName(self, "Save Consolidated File", "", "Excel Files (*.xlsx)")
            if save_path:
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"
                    
                # Save the filtered DataFrame to a single sheet
                df_filtered.to_excel(save_path, index=False)
                self.loading_overlay.set_progress(90)
                
                # Update the table view with filtered data
                self.load_table(self.file_table, df_filtered)
                
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Success", f"File saved successfully to:\n{save_path}")
            else:
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Canceled", "Save operation was canceled.")
            
        except Exception as e:
            self.loading_overlay.stop_loading()
            QMessageBox.critical(self, "Error", f"Error processing file:\n{str(e)}")
        
        finally:
            # Re-enable process button
            self.process_button.setEnabled(True)

    def load_table(self, table_view, df):
        try:
            model = QStandardItemModel()

            # Set headers
            model.setHorizontalHeaderLabels(df.columns.tolist())

            # Populate data
            for row in df.itertuples(index=False):
                items = [QStandardItem(str(value)) for value in row]
                model.appendRow(items)

            table_view.setModel(model)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error loading table:\n{str(e)}")
