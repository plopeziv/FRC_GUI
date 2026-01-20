import os, re, traceback
from qtpy.QtWidgets import *
from qtpy.QtCore import Qt
from qtpy.QtGui import QFont

class UpdateTicketDialog(QDialog):
    """Dialog for adding a new ticket"""
    
    def __init__(self, excel_manager, parent=None):
        super().__init__(parent)
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

        self.cancel_btn.clicked.connect(self.reject)

        button_layout.addWidget(self.submit_btn)
        button_layout.addWidget(self.cancel_btn)

        layout.addStretch()
        layout.addLayout(button_layout)
        