"""
FRC Ticket GUI - QtPy Version (Compatible with PyQt5/PyQt6/PySide2/PySide6)
A professional desktop application with native file dialogs and forms
"""
import sys
import numpy as np   # PyInstaller runtime anchor
np.__config__

from dialogs.add_ticket_dialog import AddTicketDialog

from qtpy.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QFileDialog,
    QTabWidget, QMessageBox, QHeaderView, QDialog
)
from qtpy.QtGui import QFont


class FRCTicketGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.manager = None
        self.ticket_data_service = None
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("FRC Ticket GUI")
        self.setGeometry(100, 100, 1200, 800)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Header
        header_label = QLabel("FRC Ticket GUI")
        header_font = QFont()
        header_font.setPointSize(20)
        header_font.setBold(True)
        header_label.setFont(header_font)
        main_layout.addWidget(header_label)
        
        subtitle = QLabel("Select Your FRC Ticket Listing Excel")
        main_layout.addWidget(subtitle)
        
        # File selection area
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected. Please use the 'Change File' button to select your Excel file.")
        self.file_label.setWordWrap(True)
        file_layout.addWidget(self.file_label, stretch=1)
        
        self.change_file_btn = QPushButton("Change File")
        self.change_file_btn.clicked.connect(self.select_file)
        self.change_file_btn.setMinimumHeight(40)
        file_layout.addWidget(self.change_file_btn)
        
        main_layout.addLayout(file_layout)
        
        # Tab widget for different views
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Create tabs
        self.ticket_listing_tab = QWidget()
        self.labor_summary_tab = QWidget()
        self.labor_details_tab = QWidget()
        self.material_summary_tab = QWidget()
        self.material_details_tab = QWidget()
        
        self.tabs.addTab(self.ticket_listing_tab, "Ticket Listings")
        self.tabs.addTab(self.labor_summary_tab, "Labor Summary")
        self.tabs.addTab(self.labor_details_tab, "Labor Details")
        self.tabs.addTab(self.material_summary_tab, "Material Summary")
        self.tabs.addTab(self.material_details_tab, "Material Details")
        
        # Setup each tab
        self.setup_ticket_listing_tab()
        self.setup_labor_summary_tab()
        self.setup_labor_details_tab()
        self.setup_material_summary_tab()
        self.setup_material_details_tab()
        
        # Initially disable tabs
        self.tabs.setEnabled(False)
    
    def setup_ticket_listing_tab(self):
        layout = QVBoxLayout(self.ticket_listing_tab)
        
        # Add row button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.add_row_btn = QPushButton("➕ Add Row")
        self.add_row_btn.clicked.connect(self.add_ticket_row)
        self.add_row_btn.setMinimumHeight(35)
        btn_layout.addWidget(self.add_row_btn)
        layout.addLayout(btn_layout)
        
        # Table
        self.ticket_listing_table = QTableWidget()
        layout.addWidget(self.ticket_listing_table)
    
    def setup_labor_summary_tab(self):
        layout = QVBoxLayout(self.labor_summary_tab)
        self.labor_summary_table = QTableWidget()
        layout.addWidget(self.labor_summary_table)
    
    def setup_labor_details_tab(self):
        layout = QVBoxLayout(self.labor_details_tab)
        self.labor_details_table = QTableWidget()
        layout.addWidget(self.labor_details_table)
    
    def setup_material_summary_tab(self):
        layout = QVBoxLayout(self.material_summary_tab)
        self.material_summary_table = QTableWidget()
        layout.addWidget(self.material_summary_table)
    
    def setup_material_details_tab(self):
        layout = QVBoxLayout(self.material_details_tab)
        self.material_details_table = QTableWidget()
        layout.addWidget(self.material_details_table)
    
    def select_file(self):
        """Open native file dialog to select Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(f"Selected file: {file_path}")
            self.load_data()
    
    def load_data(self):
        """Load data from selected Excel file"""
        from data_manager.excel_manager import ExcelManager
        from data_manager.ticket_data_service import TicketDataService

        try:
            # Load with ExcelManager
            self.manager = ExcelManager(self.file_path)
            self.manager.load()
            
            # Create ticket data service
            self.ticket_data_service = TicketDataService(self.file_path)
            
            # Enable tabs and add row button
            self.tabs.setEnabled(True)
            self.add_row_btn.setEnabled(True)
            
            # Populate all tables
            self.populate_ticket_listing()
            self.populate_labor_summary()
            self.populate_labor_details()
            self.populate_material_summary()
            self.populate_material_details()
            
            QMessageBox.information(self, "Success", "Load Successful!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error loading file: {str(e)}")
    
    def populate_table(self, table, df, currency_cols=[], currency_rows=[]):
        """Generic method to populate a QTableWidget with DataFrame data"""
        # Clear existing data
        table.clear()
        
        # Set dimensions
        table.setRowCount(len(df))
        table.setColumnCount(len(df.columns))
        
        # Set headers
        table.setHorizontalHeaderLabels([str(col) for col in df.columns])
        table.setVerticalHeaderLabels([str(idx) for idx in df.index])
        
        # Populate data
        for i in range(len(df)):
            row_idx = df.index[i]
            row_values = df.iloc[i]

            for j, col in enumerate(df.columns):
                value = row_values[col]

                if col in currency_cols or row_idx in currency_rows:
                    value = self.format_currency(value)

                item = QTableWidgetItem(str(value))
                table.setItem(i, j, item)
        
        # Auto-resize columns to contents
        try:
            # PyQt6/PySide6 style
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        except AttributeError:
            # PyQt5/PySide2 style
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    
    def populate_ticket_listing(self):
        df = self.ticket_data_service.ticket_listing.copy()
        df.index = df.index.astype(str).str.replace(",", "")
        self.populate_table(
            self.ticket_listing_table, 
            df, 
            currency_cols=[
                "Labor Sell", 
                "Labor Cost", 
                "Material Sell", 
                "Material Cost", 
                "Total Sell", 
                "Total Cost"
            ]
        )
    
    def populate_labor_summary(self):
        self.populate_table(self.labor_summary_table, self.ticket_data_service.labor_summary, 
        currency_rows=["Sell to Date", "Cost to Date", "Sell Per Unit", "Cost Per Unit (w/Tax)"]
        )
    
    def populate_labor_details(self):
        df = self.ticket_data_service.labor_ticket_summary.copy()
        df.index = df.index.astype(str).str.replace(",", "")
        self.populate_table(self.labor_details_table, df, currency_cols=[
                "Labor Sell", 
                "Labor Cost", 
            ])
    
    def populate_material_summary(self):
        self.populate_table(self.material_summary_table, self.ticket_data_service.material_summary,
        currency_rows=["Sell to Date", "Cost to Date", "Sell Per Unit", "Cost Per Unit (w/Tax)"])
    
    def populate_material_details(self):
        df = self.ticket_data_service.material_ticket_summary.copy()
        df.index = df.index.astype(str).str.replace(",", "")
        self.populate_table(self.material_details_table, df, currency_cols=[
                "Material Sell", 
                "Material Cost", 
            ])

    def format_currency(self, value):
        try:
            return f"${float(value):,.2f}"
        except:
            return str(value)
    
    def add_ticket_row(self):
        """Open dialog to add a new ticket row"""        
        dialog = AddTicketDialog(self.manager, self)

        # PyQt6/PySide6 style
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.load_data()



def main():
    app = QApplication(sys.argv)
    
    # Set application style
    try:
        app.setStyle('Fusion')
    except:
        pass  # If Fusion style not available, use default
    
    window = FRCTicketGUI()
    window.show()
    sys.exit(app.exec() if hasattr(app, 'exec') else app.exec_())


if __name__ == "__main__":
    main()