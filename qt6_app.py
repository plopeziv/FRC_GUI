"""
FRC Ticket GUI - QtPy Version (Compatible with PyQt5/PyQt6/PySide2/PySide6)
A professional desktop application with native file dialogs and forms
"""
import sys
import os
import re
import traceback

from data_manager.e_ticket_creator import ETicketCreator
from data_manager.pdf_creator import excel_to_pdf, process_ticket

from qtpy.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QFileDialog,
    QTabWidget, QMessageBox, QHeaderView, QDialog, QLineEdit, QComboBox,
    QTextEdit, QFormLayout, QListWidget, QListWidgetItem, QSpinBox,
    QDoubleSpinBox, QScrollArea, QCompleter, QCheckBox
)
from qtpy.QtCore import Qt
from qtpy.QtGui import QFont

import numpy as np

# Config needed for distribution
np.__config__


class InlineCompleterLineEdit(QLineEdit):
    def __init__(self, items, parent=None):
        super().__init__(parent)

        # Create a completer with the list of items
        self.completer = QCompleter(items, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)  # match anywhere
        self.completer.setCompletionMode(QCompleter.InlineCompletion)  # <— key line for inline autocomplete

        # Connect the completer to this QLineEdit
        self.setCompleter(self.completer)


class MaterialListWidget(QWidget):
    def __init__(self):
        super().__init__()

        self.materials_list = QListWidget()
        self.remove_btn = QPushButton("➖ Remove Selected")
        self.remove_btn.clicked.connect(self.remove_selected_materials)

        layout = QVBoxLayout()
        layout.addWidget(self.materials_list)
        layout.addWidget(self.remove_btn)
        self.setLayout(layout)

    def remove_selected_materials(self):
        selected_items = self.materials_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a material to remove")
            return

        for item in selected_items:
            material_data = item.data(Qt.UserRole)
            # Remove from your internal tracking list too
            if material_data in getattr(self, "materials_to_add", []):
                self.materials_to_add.remove(material_data)
            self.materials_list.takeItem(self.materials_list.row(item))


class AddTicketDialog(QDialog):
    """Dialog for adding a new ticket"""
    
    def __init__(self, excel_manager, parent=None):
        super().__init__(parent)
        self.excel_manager = excel_manager
        self.available_materials = list(self.excel_manager.material_map.keys())
        self.materials_to_add = []
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Add New Ticket")
        self.setMinimumWidth(800)
        self.setMinimumHeight(700)
        
        # Create scroll area for form
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        # Prevent scroll area from stealing focus from widgets like QComboBox/QCompleter
        scroll.setFocusPolicy(Qt.NoFocus)

        scroll_widget = QWidget()
        main_layout = QVBoxLayout(scroll_widget)
        
        # Title
        title = QLabel("Add New Row")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)
        
        # Ticket Info Section
        ticket_info_label = QLabel("Ticket Info")
        ticket_info_font = QFont()
        ticket_info_font.setPointSize(12)
        ticket_info_font.setBold(True)
        ticket_info_label.setFont(ticket_info_font)
        main_layout.addWidget(ticket_info_label)
        
        # Ticket info form
        ticket_form = QFormLayout()
        
        self.ticket_number = QLineEdit()
        self.ticket_number.setPlaceholderText("0")
        ticket_form.addRow("Ticket Number:", self.ticket_number)
        
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("MM/DD/YY")
        ticket_form.addRow("Date (MM/DD/YY):", self.date_input)
        
        self.signature = QComboBox()
        self.signature.addItems(["YES", "NO"])
        ticket_form.addRow("Signature:", self.signature)
        
        self.ticket_type = QComboBox()
        self.ticket_type.addItems(["REGULAR", "EXTRA", "MISC INSTALL"])
        ticket_form.addRow("Type:", self.ticket_type)
        
        self.installer_input = QLineEdit()
        self.installer_input.setPlaceholderText("Insert Installers' Names")
        ticket_form.addRow("Installers:", self.installer_input)
        
        self.work_location = QLineEdit()
        self.work_location.setPlaceholderText("Insert Work Location")
        ticket_form.addRow("Work Location:", self.work_location)
        
        self.description = QTextEdit()
        self.description.setPlaceholderText("Enter ticket details here...")
        self.description.setMaximumHeight(100)
        ticket_form.addRow("Description:", self.description)
        
        main_layout.addLayout(ticket_form)
        
        # Labor Section
        labor_label = QLabel("Add Labor")
        labor_label.setFont(ticket_info_font)
        main_layout.addWidget(labor_label)
        
        labor_layout = QHBoxLayout()
        
        self.rt_input = QLineEdit()
        self.rt_input.setPlaceholderText("0")
        labor_layout.addWidget(QLabel("RT:"))
        labor_layout.addWidget(self.rt_input)
        
        self.ot_input = QLineEdit()
        self.ot_input.setPlaceholderText("0")
        labor_layout.addWidget(QLabel("OT:"))
        labor_layout.addWidget(self.ot_input)
        
        self.dt_input = QLineEdit()
        self.dt_input.setPlaceholderText("0")
        labor_layout.addWidget(QLabel("DT:"))
        labor_layout.addWidget(self.dt_input)
        
        self.ot_diff_input = QLineEdit()
        self.ot_diff_input.setPlaceholderText("0")
        labor_layout.addWidget(QLabel("OT DIFF:"))
        labor_layout.addWidget(self.ot_diff_input)
        
        self.dt_diff_input = QLineEdit()
        self.dt_diff_input.setPlaceholderText("0")
        labor_layout.addWidget(QLabel("DT DIFF:"))
        labor_layout.addWidget(self.dt_diff_input)
        
        main_layout.addLayout(labor_layout)
        
        # Material Section
        material_label = QLabel("Add Material")
        material_label.setFont(ticket_info_font)
        main_layout.addWidget(material_label)
        
        material_input_layout = QHBoxLayout()
        
        material_input_layout.addWidget(QLabel("Quantity:"))
        self.material_qt = QLineEdit()
        self.material_qt.setPlaceholderText("0")
        self.material_qt.setMaximumWidth(100)
        material_input_layout.addWidget(self.material_qt)
        
        material_input_layout.addWidget(QLabel("Material:"))

        # Inline Completion
        self.material_input = InlineCompleterLineEdit(self.available_materials)
        self.material_input.setPlaceholderText("Start typing material...")
        self.material_input.setMinimumWidth(300)

        # Add autocomplete/search functionality
        completer = QCompleter(self.available_materials)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.material_input.setCompleter(completer)

        material_input_layout.addWidget(self.material_input)
        
        add_material_btn = QPushButton("➕ Add Material")
        add_material_btn.clicked.connect(self.add_material)
        material_input_layout.addWidget(add_material_btn)

        remove_material_btn = QPushButton("➖ Remove Material")
        remove_material_btn.clicked.connect(self.remove_material)
        material_input_layout.addWidget(remove_material_btn)
        
        material_input_layout.addStretch()
        
        main_layout.addLayout(material_input_layout)
        
        # Materials list
        self.materials_list = QListWidget()
        self.materials_list.setMaximumHeight(150)

        self.materials_list.setSelectionMode(QListWidget.SingleSelection)
        main_layout.addWidget(self.materials_list)

        # Output Folder Picker
        folder_row = QHBoxLayout()

        self.folder_label = QLabel("No folder selected")
        self.folder_label.setStyleSheet("color: gray;")

        folder_row.addWidget(self.folder_label)
        folder_row.addStretch()

        self.use_eticket_checkbox = QCheckBox("Generate E-Ticket")
        self.use_eticket_checkbox.setChecked(True)

        folder_row.addWidget(self.use_eticket_checkbox)

        main_layout.addLayout(folder_row)

        self.folder_btn = QPushButton("Select Output Folder")
        self.folder_btn.clicked.connect(self.select_output_folder)
        main_layout.addWidget(self.folder_btn)

        self.selected_folder_path = None
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        submit_btn = QPushButton("✅ Add Row")
        submit_btn.clicked.connect(self.submit_form)
        submit_btn.setMinimumHeight(40)
        submit_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        button_layout.addWidget(submit_btn)
        
        cancel_btn = QPushButton("❌ Cancel")
        cancel_btn.clicked.connect(self.reject)
        cancel_btn.setMinimumHeight(40)
        cancel_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        button_layout.addWidget(cancel_btn)
        
        main_layout.addLayout(button_layout)
        main_layout.addStretch()
        
        scroll.setWidget(scroll_widget)
        
        # Set scroll area as main layout
        dialog_layout = QVBoxLayout(self)
        dialog_layout.addWidget(scroll)
    
    def add_material(self):
        """Add material to the list"""
        try:
            quantity = self.material_qt.text().strip()
            material = self.material_input.text().strip()
            if not quantity or not material:
                QMessageBox.warning(self, "Input Error", "Please enter both quantity and material")
                return

            if material not in self.available_materials:
                QMessageBox.warning(
                    self,
                    "Invalid Material",
                    f"'{material}' is not a current material in the spreadsheet.\n"
                    f"Please select a material from the dropdown list."
                )
                return
            
            quantity_float = float(quantity)
            if quantity_float <= 0:
                QMessageBox.warning(self, "Input Error", "Quantity must be greater than zero")
                return
            
            sell_price_float = float(self.excel_manager.material_map[material]["Sell Per Unit"])
            sell_price = round(sell_price_float, 2)

            units = self.excel_manager.material_map[material]["Units"]
            
            # Check if material already exists and update
            for i in range(self.materials_list.count()):
                item = self.materials_list.item(i)
                item_data = item.data(Qt.UserRole)
                if item_data['material'] == material:
                    # Update existing
                    item_data['quantity'] = quantity
                    item.setText(f"{quantity} × {material} @ ${sell_price:.2f}")
                    QMessageBox.information(self, "Updated", f"Updated {material} quantity to {quantity}")
                    return
            
            # Add new material
            material_data = {"material": material, "quantity": quantity, "units": units, "sell price": sell_price}
            self.materials_to_add.append(material_data)
            
            item = QListWidgetItem(f"{quantity} × {material} @ ${sell_price:.2f}")
            item.setData(Qt.UserRole, material_data)
            self.materials_list.addItem(item)
            
            # Clear inputs
            self.material_qt.clear()
            self.material_qt.setFocus()
            
            
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Quantity must be a valid number")

    def remove_material(self):
        """Remove selected material from the list"""
        selected_items = self.material_input.text().strip()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a material to remove")
            return

        material_name = selected_items

        # Check if material exists in the underlying list
        if not any(m['material'] == material_name for m in self.materials_to_add):
            QMessageBox.warning(self, "Not Found", f"'{material_name}' is not in the list")
            return

        # Remove from underlying list
        self.materials_to_add = [
            m for m in self.materials_to_add if m['material'] != material_name
        ]

        # Remove from QListWidget
        for i in range(self.materials_list.count() - 1, -1, -1):
            item = self.materials_list.item(i)  # QListWidgetItem
            item_data = item.data(Qt.UserRole)   # stored dictionary
            if item_data['material'] == material_name:
                self.materials_list.takeItem(i)

        QMessageBox.information(self, "Removed", f"Removed '{material_name}'")
        
        self.material_input.clear()

    def select_output_folder(self):
        """Open a folder picker and store the selected directory"""
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")

        if folder:
            self.selected_folder_path = folder
            self.folder_label.setText(f"Folder: {folder}")
            self.folder_label.setStyleSheet("color: black; font-weight: bold;")
    
    def validate_form(self, ticket_data):
        """Validate form data"""
        errors = []
        
        # Validate date
        date_str = ticket_data["Date"]
        if not re.match(r"^\d{2}/\d{2}/\d{2}$", date_str):
            errors.append("Work Date: Must be in MM/DD/YY format")
        
        # Validate labor fields
        labor = ticket_data["Labor"]
        for key, labor_object in labor.items():
            if labor_object["hours"] == "":
                continue
            try:
                val = float(labor_object["hours"])
                if val < 0:
                    errors.append(f"{key}: Must be positive")
            except ValueError:
                errors.append(f"{key}: Must be numeric")

        # Validate that the material exists in headers before proceeding
        materials = ticket_data.get("Materials", [])
        excel_materials = self.available_materials

        missing_materials = [material for material in materials if material["material"] not in excel_materials]        
        
        for material in missing_materials:
            errors.append(
                f"Invalid Material: Material '{material}' does not exist in the material spreadsheet headers.\n"
                "Please add this material as a column first."
            )
                
        return errors
    
    def submit_form(self):
        # Check if Output Folder has been selected
        if self.use_eticket_checkbox.isChecked() and not self.selected_folder_path:
            QMessageBox.warning(
                self,
                "Missing Output Folder",
                "Please select an output folder before submitting the ticket."
            )
            return

        ticket_data = {
            'Job Number': self.excel_manager.job_number,
            'Job Name': self.excel_manager.job_name,
            "Ticket Number": self.ticket_number.text().strip(),
            'Job Address': self.excel_manager.job_address,
            "Date": self.date_input.text().strip(),
            "Signature": self.signature.currentText(),
            "Type": self.ticket_type.currentText(),
            'Installers': self.installer_input.text().strip(),
            'Work Location': self.work_location.text().strip(),
            "Description": self.description.toPlainText().strip(),
            "Labor": {
                "RT": {
                    "hours": self.rt_input.text().strip(), 
                    "rate": self.excel_manager.labor_map["RT"]["rate"]
                    },
                "OT": {
                    "hours": self.ot_input.text().strip(), 
                    "rate": self.excel_manager.labor_map["OT"]["rate"]
                    },
                "DT": {
                    "hours": self.dt_input.text().strip(), 
                    "rate": self.excel_manager.labor_map["DT"]["rate"]
                    },
                "OT DIFF": {
                    "hours": self.ot_diff_input.text().strip(), 
                    "rate": self.excel_manager.labor_map["OT DIFF"]["rate"]
                    },
                "DT DIFF": {
                    "hours": self.dt_diff_input.text().strip(), 
                    "rate": self.excel_manager.labor_map["DT DIFF"]["rate"]
                    },
            },
            "Materials": self.materials_to_add
        }
        
        # Validate
        errors = self.validate_form(ticket_data)
        
        if errors:
            error_msg = "\n".join(errors)
            QMessageBox.warning(self, "Validation Errors", error_msg)
            return
        
        # Try to save
        try:
            # create and save e-ticket files
            if self.use_eticket_checkbox.isChecked():
                e_ticket_creator = ETicketCreator(self.selected_folder_path, ticket_data)
                e_ticket_creator.load_ticket()
                
                # Save PDF
                excel_file_path = os.path.join(
                    self.selected_folder_path,
                    f"{ticket_data['Job Number']} - {ticket_data['Ticket Number']}.xlsx"
                )
                process_ticket(excel_file_path, ticket_data["Date"], ticket_data["Signature"])

            # Insert row to ticket listing
            self.excel_manager.insert_ticket(ticket_data)
            self.excel_manager.load()

            QMessageBox.information(self, "Success", "✅ Ticket created and saved successfully!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"❌ Failed to save ticket: {str(e)}")
            traceback.print_exc()


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