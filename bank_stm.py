import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, 
    QFileDialog, QMessageBox, QTableView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
import os
from loading_overlay import LoadingOverlay


class BankStatementProcessor(QWidget):
    def __init__(self):
        super().__init__()

        # Reuse styles from ExcelUploader
        drop_zone_style = """
            QLabel {
                background-color: white;
                border: 2px dashed #aaaaaa;
                border-radius: 8px;
                padding: 20px;
                color: #666666;
                font-size: 14px;
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
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """

        # Apply styles and update layout
        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(15)
        self.main_layout.setContentsMargins(20, 20, 20, 20)

        # Bank Statement Upload Area
        self.bank_statement_label = QLabel("🏦 Upload bank statement here\nclick to browse")
        self.bank_statement_label.setStyleSheet(drop_zone_style)
        self.bank_statement_label.setAlignment(Qt.AlignCenter)
        self.bank_statement_label.mousePressEvent = self.upload_bank_statement
        self.bank_statement_path = None
        self.main_layout.addWidget(self.bank_statement_label)

        # Table View for displaying data
        self.table_view = QTableView()
        self.main_layout.addWidget(self.table_view)

        # Process Button
        self.process_button = QPushButton("Process Bank Statement")
        self.process_button.setStyleSheet(button_style)
        self.process_button.clicked.connect(self.process_statement)
        self.main_layout.addWidget(self.process_button)

        self.setLayout(self.main_layout)

        # Add loading overlay
        self.loading_overlay = LoadingOverlay(self)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.loading_overlay.setFixedSize(self.size())

    def upload_bank_statement(self, event):
        self.bank_statement_path, _ = QFileDialog.getOpenFileName(
            self, "Select Bank Statement", "", "Excel Files (*.xlsx)"
        )
        if self.bank_statement_path:
            self.bank_statement_label.setText(
                f"Selected: {os.path.basename(self.bank_statement_path)}"
            )
            self.load_table(self.bank_statement_path)

    def load_table(self, file_path):
        # Disable the process button
        self.process_button.setEnabled(False)
        
        # Show loading overlay
        self.loading_overlay.start_loading("Loading bank statement...")

        try:
            df = pd.read_excel(file_path)
            self.loading_overlay.set_progress(50)
            
            model = QStandardItemModel()
            
            # Set headers
            model.setHorizontalHeaderLabels(df.columns.tolist())
            
            # Populate data
            for row in df.itertuples(index=False):
                items = [QStandardItem(str(value)) for value in row]
                model.appendRow(items)
            
            self.table_view.setModel(model)
            self.loading_overlay.set_progress(100)
            
            # Stop loading overlay before completion
            self.loading_overlay.stop_loading()
            
        except Exception as e:
            # Stop loading before showing error message
            self.loading_overlay.stop_loading()
            QMessageBox.critical(self, "Error", f"Error loading file:\n{str(e)}")
        
        finally:
            # Re-enable the process button
            self.process_button.setEnabled(True)

    def process_statement(self):
        if not self.bank_statement_path:
            QMessageBox.warning(self, "Error", "Please upload a bank statement first.")
            return

        # Disable the process button
        self.process_button.setEnabled(False)
        
        # Show loading overlay
        self.loading_overlay.start_loading("Processing bank statement...")

        try:
            # Read and process bank statement
            bank_statement = pd.read_excel(self.bank_statement_path)
            self.loading_overlay.set_progress(30)
            
            # Add app column based on transaction particulars
            bank_statement['app'] = bank_statement['Transaction Particulars'].apply(self._get_bank_statement_app)
            bank_statement['Tran Date'] = pd.to_datetime(bank_statement['Tran Date']).dt.strftime('%Y-%m-%d')
            self.loading_overlay.set_progress(60)
            
            # Group by app and date
            bank_amounts = bank_statement.groupby(['app', 'Tran Date'])['Amount(INR)'].sum().reset_index()
            self.loading_overlay.set_progress(80)

            # Save the processed data
            save_path, _ = QFileDialog.getSaveFileName(
                self, "Save Processed Bank Statement", "", "Excel Files (*.xlsx)"
            )
            if save_path:
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"
                
                bank_amounts.to_excel(save_path, sheet_name="Bank Amounts", index=False)
                
                # Stop loading before showing success message
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Success", f"Processed data saved to:\n{save_path}")
            else:
                # Stop loading before showing canceled message
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Canceled", "Save operation was canceled.")
            
        except Exception as e:
            # Stop loading before showing error message
            self.loading_overlay.stop_loading()
            QMessageBox.critical(self, "Error", f"Error processing file:\n{str(e)}")
        
        finally:
            # Re-enable the process button
            self.process_button.setEnabled(True)

    def _get_bank_statement_app(self, transaction_particulars):
        particulars = str(transaction_particulars).lower()
        if "moving tech innovations" in particulars or "ypp limit neft" in particulars:
            return "nammayathri"
        elif "redbus" in particulars:
            return "redbus"
        elif "roppen" in particulars:
            return "rapido"
        elif "paytm" in particulars or "pai platforms" in particulars:
            return "paytm"
        elif "easytrip" in particulars:
            return "easemytrip"
        elif "922020004688715" in particulars or "phonepe" in particulars:
            return "phonepe"
        return "unknown"
