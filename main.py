import sys
import os
import json
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from excel_compare import ExcelUploader
from settlement_process import SingleFileUploader
from bank_stm import BankStatementProcessor
from row_remover import ConsolidateUploader
from settings import SettingsTab

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize config file path first
        appdata_dir = os.getenv("APPDATA") if os.name == "nt" else os.path.expanduser("~/.config")
        app_folder = os.path.join(appdata_dir, "kochimetro")
        os.makedirs(app_folder, exist_ok=True)
        self.config_file_path = os.path.join(app_folder, "config.json")
        
        # Now load the config
        self.config_file_path, self.config = self.load_config()
        
        # Initialize UI
        self.setWindowTitle("Data Consolidator")
        self.setGeometry(100, 100, 1200, 800)

        # Create tabs
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Create instances of each tab
        self.excel_compare_tab = ExcelUploader()
        self.single_file_tab = SingleFileUploader(list(self.config.keys()))
        self.settings_tab = SettingsTab(self.config_file_path)
        self.row_remover_tab = ConsolidateUploader()
        self.bank_statement_tab = BankStatementProcessor()

        # Add tabs to the widget
        self.tabs.addTab(self.excel_compare_tab, "Compare")
        self.tabs.addTab(self.single_file_tab, "Settlement")
        self.tabs.addTab(self.bank_statement_tab, "Bank Statement")
        self.tabs.addTab(self.settings_tab, "Settings")
        self.tabs.addTab(self.row_remover_tab, "Row Remover")

    def load_config(self):
        # Get AppData directory
        appdata_dir = os.getenv("APPDATA") if os.name == "nt" else os.path.expanduser("~/.config")
        app_folder = os.path.join(appdata_dir, "kochimetro")
        os.makedirs(app_folder, exist_ok=True)
        config_path = os.path.join(app_folder, "config.json")

        # Load or create config.json
        if os.path.exists(config_path):
            with open(config_path, "r") as file:
                config = json.load(file)
        else:
            # Default configuration
            config = {
                'easemytrip': {'id_col': 'Ticket Id', 'match_col': 'TicketNUmber', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'nammayathri': {'id_col': 'Ticket Id', 'match_col': 'TicketNUmber', 'amount_col': 'Total Amount', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'phonepe': {'id_col': 'Ticket Id', 'match_col': 'TicketNUmber', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'paytm': {'id_col': 'Operator Reference Number', 'match_col': 'order_id', 'amount_col': 'Total Price', 'settle_col': 'Payable Amount', 'date_col': 'Settlement Date'},
                'rapido': {'id_col': 'Network Order ID', 'match_col': 'transaction_ref_no', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'},
                'redbus': {'id_col': 'Network Order ID(From ondcTxnId)', 'match_col': 'transaction_ref_no', 'amount_col': 'TOTAL AMOUNT', 'settle_col': 'Settlement Amount', 'date_col': 'Date'}
            }

            # Save the default configuration to the file
            with open(config_path, "w") as file:
                json.dump(config, file, indent=4)

        return config_path, config

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())