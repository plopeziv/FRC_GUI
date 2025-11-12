import pandas as pd

from excel_manager import ExcelManager

class TicketDataService:
    def __init__(self, file_path=None):
        self.file_path = file_path;
        self.material_ticket_details = None

        self.excel_manager = ExcelManager(self.file_path);

        self.excel_manager.load()



if __name__ =="__main__":
    
    test_path = "027386 - DETAILED TICKET LISTING - PYTHON Copy.xlsx"
    
    manager = TicketDataService(test_path).excel_manager
