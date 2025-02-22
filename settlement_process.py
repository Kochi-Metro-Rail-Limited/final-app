import traceback
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, 
    QFileDialog, QMessageBox, QTableView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import os
import json
from loading_overlay import LoadingOverlay
import numpy as np

class SingleFileUploader(QWidget):
    """
    GUI component that handles uploading and processing settlement files.
    - Allows uploading a main AFC-triffi file
    - Allows uploading settlement reports from different payment apps
    - Provides functionality to generate summaries and merged documents
    """
    def __init__(self, app_names):
        super().__init__()
        
        # Load configuration for supported payment apps (PayTm, PhonePe, etc.)
        appdata_dir = os.getenv("APPDATA") if os.name == "nt" else os.path.expanduser("~/.config")
        app_folder = os.path.join(appdata_dir, "kochimetro")
        config_path = os.path.join(app_folder, "config.json")
        
        with open(config_path, "r") as file:
            self.config = json.load(file)
        
        self.app_names = list(self.config.keys())

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

        # File Upload Area for the single Excel file
        self.file_label = QLabel("Upload a single Excel file")
        self.file_label.setStyleSheet(drop_zone_style)
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.mousePressEvent = self.upload_file
        self.file_path = None
        self.main_layout.addWidget(self.file_label)

        # Unified Upload Area for Settlement Reports
        self.settlement_label = QLabel("Upload all settlement reports")
        self.settlement_label.setStyleSheet(drop_zone_style)
        self.settlement_label.setAlignment(Qt.AlignCenter)
        self.settlement_label.mousePressEvent = self.upload_settlement_files
        self.settlement_files = {}
        self.main_layout.addWidget(self.settlement_label)

        # File Table for the single file
        self.file_table = QTableView()
        self.main_layout.addWidget(self.file_table)

        # Get Merged Doc Button
        self.merged_doc_button = QPushButton("Get Merged Doc")
        self.merged_doc_button.clicked.connect(self.get_merged_doc)
        self.merged_doc_button.setStyleSheet(button_style)
        self.main_layout.addWidget(self.merged_doc_button)

        # Get Summary Button
        self.summary_button = QPushButton("Get Summary")
        self.summary_button.clicked.connect(self.get_summary)
        self.summary_button.setStyleSheet(button_style)
        self.main_layout.addWidget(self.summary_button)

        # Apply styles to widgets
        self.file_label.setStyleSheet(drop_zone_style)
        self.settlement_label.setStyleSheet(drop_zone_style)
        self.summary_button.setStyleSheet(button_style)

        # Update labels with icons
        self.file_label.setText("📄 AFC-triffi file here\n click to browse")
        self.settlement_label.setText("📊 App settlement reports here\n click to browse")
        self.summary_button.setText("Get Summary")

        self.setLayout(self.main_layout)

        # Add loading overlay
        self.loading_overlay = LoadingOverlay(self)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.loading_overlay.setFixedSize(self.size())

    def upload_file(self, event):
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if self.file_path:
            self.file_label.setText(self.file_path.split('/')[-1])
            self.load_table(self.file_table, self.file_path)

    def upload_settlement_files(self, event):
        """
        Handles uploading of settlement report files from different payment apps.
        Validates that:
        1. Correct number of files are uploaded (one per payment app)
        2. Each file name contains the corresponding app name
        """
        files, _ = QFileDialog.getOpenFileNames(self, "Select Settlement Reports", "", "Excel Files (*.xlsx)")
        
        # Validate number of files matches number of payment apps
        if len(files) > len(self.app_names):
            QMessageBox.critical(self, "Error", f"More than the required number of files have been uploaded. Upload only {len(self.app_names)} files")
            return
        if len(files) < len(self.app_names):
            QMessageBox.critical(self, "Error", f"Less than the required number of files have been uploaded. Upload {len(self.app_names)} files")
            return

        # Map files to their corresponding apps based on filename
        if files:
            try:
                self.settlement_files = {self.get_app_name(os.path.basename(file)): file for file in files}
                self.settlement_label.setText(f"Uploaded {len(files)} settlement files")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Unrecognized file detected: {e.args[1]}. Make sure the name of the app exists in the filename")

    def get_app_name(self, file_name):
        app_keywords = self.app_names
        for keyword in app_keywords:
            if keyword.lower() in file_name.lower():
                return keyword
        raise Exception("Unrecognized file detected", file_name)
        

    def load_table(self, table_view, file_path):
        try:
            df = pd.read_excel(file_path)
            model = QStandardItemModel()

            # Set headers
            model.setHorizontalHeaderLabels(df.columns.tolist())

            # Populate data
            for row in df.itertuples(index=False):
                items = [QStandardItem(str(value)) for value in row]
                model.appendRow(items)

            table_view.setModel(model)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error loading file:\n{str(e)}")

    def get_summary(self):
        if not self.file_path:
            QMessageBox.warning(self, "Error", "No file uploaded for the main file.")
            return

        if not self.settlement_files:
            QMessageBox.warning(self, "Error", "No settlement files uploaded.")
            return

        # Disable buttons
        self.summary_button.setEnabled(False)
        self.merged_doc_button.setEnabled(False)
        
        # Show loading overlay
        self.loading_overlay.start_loading("Processing files...")

        try:
            # Load the main file
            df = pd.read_excel(self.file_path)
            self.loading_overlay.set_progress(20)

            # First try AppData directory
            # appdata_dir = os.getenv("APPDATA") if os.name == "nt" else os.path.expanduser("~/.config")
            # app_folder = os.path.join(appdata_dir, "kochimetro")
            # do_not_delete_path = os.path.join(app_folder, "DO_NOT_DELETE.csv")

            # # If AppData path doesn't exist, use local project directory
            # if not os.path.exists(do_not_delete_path):
            #     do_not_delete_path = os.path.join(os.path.dirname(__file__), "DO_NOT_DELETE.csv")

            # if not os.path.exists(do_not_delete_path):
            #     self.loading_overlay.stop_loading()
            #     QMessageBox.critical(self, "Error", "Please upload AFC and Triffy files again.")
            #     return

            # original_df = pd.read_csv(do_not_delete_path)

            original_df = df
            self.loading_overlay.set_progress(40)

            # Create an instance of Process with the original DataFrame
            process = Process(original_df, self.settlement_files)
            self.loading_overlay.set_progress(70)

            save_path, _ = QFileDialog.getSaveFileName(self, "Save Summary File", "", "Excel Files (*.xlsx)")
            if save_path:
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"

                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    if not process.sheet1.empty:
                        process.sheet1.to_excel(writer, sheet_name="Grouped Data", index=False)
                    else:
                        pd.DataFrame().to_excel(writer, sheet_name="No Data", index=False)

                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Success", f"Summary saved to:\n{save_path}")
            else:
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Canceled", "Save operation was canceled.")

        except Exception as e:
            self.loading_overlay.stop_loading()
            QMessageBox.critical(self, "Error", f"Error processing file:\n{str(e)}")
        
        finally:
            # Re-enable buttons
            self.summary_button.setEnabled(True)
            self.merged_doc_button.setEnabled(True)

    def get_merged_doc(self):
        if not self.file_path:
            QMessageBox.warning(self, "Error", "No file uploaded for the main file.")
            return

        if not self.settlement_files:
            QMessageBox.warning(self, "Error", "No settlement files uploaded.")
            return

        # Disable buttons
        self.summary_button.setEnabled(False)
        self.merged_doc_button.setEnabled(False)
        
        # Show loading overlay
        self.loading_overlay.start_loading("Processing files...")

        try:
            # Load the main file
            df = pd.read_excel(self.file_path)
            self.loading_overlay.set_progress(20)

            # First try AppData directory
            # appdata_dir = os.getenv("APPDATA") if os.name == "nt" else os.path.expanduser("~/.config")
            # app_folder = os.path.join(appdata_dir, "kochimetro")
            # do_not_delete_path = os.path.join(app_folder, "DO_NOT_DELETE.csv")

            # # If AppData path doesn't exist, use local project directory
            # if not os.path.exists(do_not_delete_path):
            #     do_not_delete_path = os.path.join(os.path.dirname(__file__), "DO_NOT_DELETE.csv")

            # if not os.path.exists(do_not_delete_path):
            #     self.loading_overlay.stop_loading()
            #     QMessageBox.critical(self, "Error", "Please upload AFC and Triffy files again.")
            #     return

            # original_df = pd.read_csv(do_not_delete_path)

            original_df = df
            self.loading_overlay.set_progress(40)

            # Create an instance of Process with the original DataFrame
            process = Process(original_df, self.settlement_files)
            self.loading_overlay.set_progress(70)

            save_path, _ = QFileDialog.getSaveFileName(self, "Save Merged Document", "", "Excel Files (*.xlsx)")
            if save_path:
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"
                
                # Pre-filter the data before writing to Excel
                if not process.sheet4.empty:
                    # Add Action column efficiently using numpy
                    process.sheet4['Action'] = ''
                    
                    # Use boolean indexing instead of multiple filtered copies
                    mask_settled = process.sheet4['result'] == 'settled'
                    mask_excess = process.sheet4['result'] == 'excess'
                    mask_shortage = process.sheet4['result'] == 'shortage'

                    # Write to Excel efficiently using a context manager
                    with pd.ExcelWriter(save_path, engine='openpyxl', mode='w') as writer:
                        # Write main sheet
                        process.sheet4.to_excel(writer, sheet_name="Merged Data", index=False)
                        
                        # Write filtered sheets without creating separate DataFrames
                        process.sheet4[mask_settled].to_excel(writer, sheet_name="Settled", index=False)
                        process.sheet4[mask_excess].to_excel(writer, sheet_name="Excess", index=False)
                        process.sheet4[mask_shortage].to_excel(writer, sheet_name="Shortage", index=False)
                else:
                    pd.DataFrame().to_excel(save_path, sheet_name="No Data", index=False)

                # Add data validation more efficiently
                workbook = load_workbook(save_path)
                sheet = workbook["Merged Data"]
                action_col = sheet.max_column
                
                # Create validation rule once
                validation = DataValidation(
                    type="list",
                    formula1='"Option1,Option2,Option3"',
                    allow_blank=True
                )
                sheet.add_data_validation(validation)
                
                # Add validation to entire column range at once
                validation.add(f"{chr(64 + action_col)}2:{chr(64 + action_col)}{sheet.max_row}")
                
                workbook.save(save_path)
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Success", f"Merged document saved to:\n{save_path}")
            else:
                self.loading_overlay.stop_loading()
                QMessageBox.information(self, "Canceled", "Save operation was canceled.")

        except Exception as e:
            self.loading_overlay.stop_loading()
            QMessageBox.critical(self, "Error", f"Error processing file:\n{str(e)}")
        
        finally:
            # Re-enable buttons
            self.summary_button.setEnabled(True)
            self.merged_doc_button.setEnabled(True)

class Process:
    """
    Core processing logic for settlement reconciliation.
    
    Flow:
    1. Loads and normalizes the main transaction data
    2. Loads and processes settlement files from different payment apps
    3. Merges settlement data with original transactions
    4. Identifies discrepancies (excess/shortage/settled)
    5. Generates summary reports
    """
    def __init__(self, original_df, settlement_files):
        self.original_df = original_df  # Main transaction data
        self.settlement_files = {}      # Settlement data from payment apps
        self.load_config()             # Load app-specific column mappings
        
        # Process each settlement file
        for app_name, file_path in settlement_files.items():
            app_name = app_name.lower()
            if app_name in self.app_mapping:
                self.settlement_files[app_name] = pd.read_excel(file_path)
        
        # Standardize formats and process data
        self._normalize_original_df()
        self._process_settlement_files()
        
        # Merge and analyze data
        self.merged_data = self._merge_settlement_data()
        self.grouped_data = self._summarize_transactions()
        
        # Prepare output sheets
        self.sheet1 = self.grouped_data    # Summary by app and date
        self.sheet4 = self.merged_data     # Detailed transaction matching

    def load_config(self):
        # Get AppData directory
        appdata_dir = os.getenv("APPDATA") if os.name == "nt" else os.path.expanduser("~/.config")
        app_folder = os.path.join(appdata_dir, "kochimetro")
        os.makedirs(app_folder, exist_ok=True)
        config_path = os.path.join(app_folder, "config.json")
        # Load or create config.json
        if os.path.exists(config_path):
            with open(config_path, "r") as file:
                self.app_mapping = json.load(file)
        else:
            # Default configuration
            default_config = {
                'easemytrip': {'id_col': 'TicketId', 'match_col': 'TicketNUmber', 'amount_col': 'TOTALAMOUNT', 'settle_col': 'SettlementAmount', 'date_col': 'Date'},
                'nammayathri': {'id_col': 'Ticket Id', 'match_col': 'TicketNUmber', 'amount_col': 'Total Amount', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'phonepe': {'id_col': 'Ticket Id', 'match_col': 'TicketNUmber', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'paytm': {'id_col': 'Operator Reference Number', 'match_col': 'order_id', 'amount_col': 'Total Price', 'settle_col': 'Payable Amount', 'date_col': 'Settlement Date'},
                'rapido': {'id_col': 'Network Order ID', 'match_col': 'transaction_ref_no', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'redbus': {'id_col': 'Network Order ID(From ondcTxnId)', 'match_col': 'transaction_ref_no', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'}
            }

            # Save the default configuration to the file
            with open(config_path, "w") as file:
                json.dump(default_config, file, indent=4)
            self.app_mapping = default_config

    def _normalize_original_df(self):
        self.original_df['ONDCapp'] = self.original_df['ONDCapp'].str.lower()
        self.original_df = self.original_df.replace('yathri', 'nammayathri')
        self.original_df['insertDT'] = self._standardize_date(self.original_df['insertDT'])

    def _process_settlement_files(self):
        for app_name, df in self.settlement_files.items():
            try:
                mapping = self.app_mapping[app_name]
                date_col = mapping['date_col']
                
                # Print debug info before processing
                print(f"\n[{app_name}] Date Processing:")
                print("Before normalization:")
                print(f"Column '{date_col}' - First 5 values:")
                print(df[date_col].head().to_string())
                print(f"Data type: {df[date_col].dtype}")
                
                # Handle UTC timestamps (like in nammayathri)
                if 'UTC' in str(df[date_col].iloc[0]):
                    df[date_col] = pd.to_datetime(df[date_col]).dt.tz_localize(None)
                else:
                    # Convert dates without timezone info
                    df[date_col] = pd.to_datetime(df[date_col], format='mixed')
                
                # Debug after datetime conversion
                print("\nAfter datetime conversion:")
                print(df[date_col].head().to_string())
                
                # Convert to string format (only date part)
                df[date_col] = df[date_col].dt.date.astype(str)
                
                # Debug after final formatting
                print("\nAfter final formatting:")
                print(df[date_col].head().to_string())
                print("----------------------------------------")
                
                self.settlement_files[app_name] = df
                
            except Exception as e:
                print(f"\n[{app_name}] Error processing dates:")
                print(f"  Error: {str(e)}")
                print(f"  Date column type: {type(df[date_col])}")
                print(f"  Raw date values: {df[date_col].head().tolist()}")
                continue

    def _merge_settlement_data(self):
        """
        Core reconciliation logic using outer join to identify matched, excess, and shortage records
        """
        # Create base DataFrame with required columns
        merged_data = self.original_df[['insertDT', 'TicketNUmber', 'order_id', 
                                      'transaction_ref_no', 'ONDCapp', 'total_amount', 
                                      'QRCodePrice', 'booking_status', 'descCode', 'Remark']].copy()
        
        # Add empty settlement columns
        merged_data['amount_col'] = None
        merged_data['settle_col'] = None
        merged_data['result'] = None
        merged_data['duplicate'] = None

        final_merged_data = pd.DataFrame()  # Create empty DataFrame to store results

        # Process each payment app
        for app_name, mapping in self.app_mapping.items():
            try:
                settlement_df = self.settlement_files.get(app_name)
                if settlement_df is None:
                    continue

                # Get data for current app
                app_data = merged_data[merged_data['ONDCapp'] == app_name].copy()

                # Get required columns from settlement file and remove duplicates
                settlement_data = settlement_df[[
                    mapping['id_col'], 
                    mapping['amount_col'], 
                    mapping['settle_col'],
                    mapping['date_col']
                ]].copy()

                # Remove duplicates from settlement data based on ID and amount
                subset_cols = [mapping['id_col'], mapping['amount_col'], mapping['settle_col']]
                settlement_data['is_duplicate'] = settlement_data.duplicated(subset=subset_cols, keep='first')

                # Clean Paytm IDs by removing trailing '...'
                if app_name == 'paytm':
                    settlement_data[mapping['id_col']] = settlement_data[mapping['id_col']].astype(str).str.replace('...', '').str.strip()

                if app_name == 'nammayathri':
                    # Set negative settlement values to 0
                    settlement_data[mapping['settle_col']] = settlement_data[mapping['settle_col']].apply(lambda x: max(0, x))

                # Format settlement date to match insertDT format (YYYY-MM-DD 00:00:00)
                if mapping['date_col'] in settlement_data.columns:
                    settlement_data[mapping['date_col']] = pd.to_datetime(settlement_data[mapping['date_col']]).dt.strftime('%Y-%m-%d 00:00:00')

                # Perform outer join 
                merged = pd.merge(
                    app_data,
                    settlement_data,
                    left_on=mapping['match_col'],
                    right_on=mapping['id_col'],
                    how='outer'
                )

                merged['duplicate'] = merged['is_duplicate'].fillna(False)

                # Fill ONDCapp for settlement-only records
                merged['ONDCapp'] = merged['ONDCapp'].fillna(app_name)

                # Handle dates
                merged['insertDT'] = merged['insertDT'].fillna(merged[mapping['date_col']])
                merged.loc[merged['insertDT'].isna(), 'insertDT'] = merged[mapping['date_col']]

                # Fill all ID-related columns based on the mapping
                if mapping['match_col'] == 'TicketNUmber':
                    merged['TicketNUmber'] = merged['TicketNUmber'].fillna(merged[mapping['id_col']])
                elif mapping['match_col'] == 'order_id':
                    merged['order_id'] = merged['order_id'].fillna(merged[mapping['id_col']])
                elif mapping['match_col'] == 'transaction_ref_no':
                    merged['transaction_ref_no'] = merged['transaction_ref_no'].fillna(merged[mapping['id_col']])

                # Set result based on conditions
                merged['amount_col'] = merged[mapping['amount_col']]
                merged['settle_col'] = merged[mapping['settle_col']]


                # Shortage: settlement amounts are empty
                shortage_mask = merged['settle_col'].isna() & merged['amount_col'].isna()
                merged.loc[shortage_mask, 'result'] = 'shortage'

                # Excess: settlement amounts exist but original amounts don't
                excess_mask = (~merged['settle_col'].isna() & ~merged['amount_col'].isna() & 
                              merged['total_amount'].isna() & merged['QRCodePrice'].isna())
                merged.loc[excess_mask, 'result'] = 'excess'

                # Settled: all amounts exist
                settled_mask = (~merged['settle_col'].isna() & ~merged['amount_col'].isna() & 
                              ~merged['total_amount'].isna() & ~merged['QRCodePrice'].isna())
                merged.loc[settled_mask, 'result'] = 'settled'

                # Add to final DataFrame
                final_merged_data = pd.concat([
                    final_merged_data,
                    merged[merged_data.columns]
                ])

                # Print summary
                print(f"\n[{app_name}] Processing Summary:")
                print(f"  Settled: {(merged['result'] == 'settled').sum()}")
                print(f"  Shortage: {(merged['result'] == 'shortage').sum()}")
                print(f"  Excess: {(merged['result'] == 'excess').sum()}")

            except Exception as e:
                print(f"\n[{app_name}] Error: {str(e)}")
                continue

        # Add unprocessed records from other apps
        processed_apps = set(self.settlement_files.keys())
        unprocessed_data = merged_data[~merged_data['ONDCapp'].isin(processed_apps)]
        final_merged_data = pd.concat([final_merged_data, unprocessed_data])

        # Ensure consistent column order
        columns = ['insertDT', 'TicketNUmber', 'order_id', 'transaction_ref_no', 'ONDCapp', 
            'total_amount', 'QRCodePrice', 'booking_status', 'descCode', 'Remark', 
            'amount_col', 'settle_col', 'result', 'duplicate']
        
        # Add debug logging to check final date coverage
        total_rows = len(final_merged_data)
        missing_dates = final_merged_data['insertDT'].isna().sum()
        print(f"\nDate Coverage Summary:")
        print(f"Total rows: {total_rows}")
        print(f"Rows with dates: {total_rows - missing_dates}")
        print(f"Rows missing dates: {missing_dates}")

        return final_merged_data[columns]

    def _summarize_transactions(self):
        grouped_data = self.merged_data.groupby(['ONDCapp', 'insertDT']).agg({
            'QRCodePrice': 'sum',
            'total_amount': 'sum',
            'amount_col': 'sum',
            'settle_col': 'sum'
        }).reset_index()
        grouped_data.rename(columns={
            'QRCodePrice': 'original_amount(afc)',
            'total_amount': 'original_amount(triffi)',
            'amount_col': 'total_amount',
            'settle_col': 'settlement_amount'
        }, inplace=True)
        return grouped_data

    def _standardize_date(self, date_series):
        """Standardize dates to YYYY-MM-DD format"""
        try:
            # Convert to datetime, handling multiple formats
            if 'UTC' in str(date_series.iloc[0]):
                date_series = pd.to_datetime(date_series).dt.tz_localize(None)
            else:
                date_series = pd.to_datetime(date_series, format='mixed')
            
            # Convert to date only (removes time component)
            return date_series.dt.date.astype(str)
            
        except Exception as e:
            print(f"Error standardizing dates: {e}")
            return date_series
