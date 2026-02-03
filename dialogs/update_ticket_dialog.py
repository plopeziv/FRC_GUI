import os, re, traceback

from dialogs.add_ticket_dialog import AddTicketDialog

from qtpy.QtWidgets import *
from qtpy.QtCore import Qt
from qtpy.QtGui import QFont

class UpdateTicketDialog(QDialog):
    """Dialog for adding a new ticket"""
    
    def __init__(self, ticket_data_service, excel_manager, parent=None):
        super().__init__(parent)
        self.ticket_data_service = ticket_data_service
        self.excel_manager = excel_manager
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("Update Existing Ticket")
        self.setMinimumWidth(400)
        self.setMinimumHeight(125)

        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)

        #Title
        title = QLabel("Update Ticket")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignTop)
        layout.addWidget(title)
        layout.setSpacing(15)
        

        # Ticket number input
        input_layout = QHBoxLayout()
        ticket_label = QLabel("Ticket Number: ")

        input_layout.addWidget(ticket_label)

        self.ticket_input = QLineEdit()
        self.ticket_input.setPlaceholderText("Please Enter Ticket Number")
        self.ticket_input.setMinimumHeight(30)

        input_layout.addWidget(self.ticket_input)
        
        layout.addLayout(input_layout)

        #Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.setSpacing(8)

        self.submit_btn = QPushButton("Submit")
        self.cancel_btn = QPushButton("Close")

        self.submit_btn.clicked.connect(self.onSubmit)
        self.cancel_btn.clicked.connect(self.reject)
        

        button_layout.addWidget(self.submit_btn)
        button_layout.addWidget(self.cancel_btn)

        layout.addStretch()
        layout.addLayout(button_layout)

    def onSubmit(self):
        ticket_number = self.ticket_input.text().strip()

        if not ticket_number:
            QMessageBox.warning(self, "Input Required", "Please enter a ticket number.")
            return

        df = self.ticket_data_service.ticket_listing
        # make sure all ticket numbers are in string format
        df.index = df.index.map(lambda x: str(x).strip())


        if ticket_number in df.index:
            ticket = self.prepTicket(ticket_number)
            dialog=AddTicketDialog(self.excel_manager, ticket, parent=self)

            # Make sure ticket UpdateDialog closes with AddTicketDialog
            result = dialog.exec()
            if result == QDialog.Accepted:
                self.accept()
            else:
                self.reject()

        else:
            QMessageBox.warning(self, "Not Found", f"Ticket {ticket_number} Does Not Exist In Listing.")
            print(False)
    
    def prepTicket(self, ticket_number):
        blank_ticket = {
            "Ticket Number": ticket_number,
            "Date": self.ticket_data_service.ticket_listing.loc[ticket_number]["Date"],
            "Signature": self.ticket_data_service.ticket_listing.loc[ticket_number]["Signed"],
            "Type": self.ticket_data_service.ticket_listing.loc[ticket_number]["Type\n(Regular, Extra)"],
            "Description": self.ticket_data_service.ticket_listing.loc[ticket_number]["Description"],
            "Labor": {
                "RT": {
                    "hours": self.ticket_data_service.labor_ticket_summary.loc[ticket_number]["RT"], 
                    },
                "OT": {
                    "hours": self.ticket_data_service.labor_ticket_summary.loc[ticket_number]["OT"],  
                    },
                "DT": {
                    "hours": self.ticket_data_service.labor_ticket_summary.loc[ticket_number]["DT"],  
                    },
                "OT DIFF": {
                    "hours": self.ticket_data_service.labor_ticket_summary.loc[ticket_number]["OT DIFF "],  
                    },
                "DT DIFF": {
                    "hours": self.ticket_data_service.labor_ticket_summary.loc[ticket_number]["DT DIFF "],  
                    },
            },
            "Materials": []
        }

        return blank_ticket

        