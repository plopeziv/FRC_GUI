# -*- coding: utf-8 -*-
"""
Created on Sat Nov 22 12:04:44 2025

@author: plopez
"""
from pathlib import Path

class ETicketCreator:
    def __init__(self, file_path=None, incoming_ticket=None):
        self.file_path = file_path;
        self.workbook = None
        self.incoming_ticket = incoming_ticket;
        
    def load_ticket(self):
        import pandas as pd
        from openpyxl import load_workbook
        
        if not self.file_path:
            raise ValueError("No file provided!")
            
        path = Path(self.file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        
        path = self._ensure_xlsx_copy()
        wb = load_workbook(path)
        
        # raw_df = pd.read_excel(path, engine=self._get_excel_engine(), header=None)
        raw_df = self._insert_job_info(wb)
        
        return raw_df
    
    def _ensure_xlsx_copy(self):
        """
        If the working copy path is .xls, convert it to .xlsx for openpyxl.
        Returns the new path (string) compatible with openpyxl.
        """
        import pandas as pd

        path = Path(self.file_path)
        if path.suffix.lower() == ".xls":
            new_path = path.with_suffix(".xlsx")
            if not new_path.exists():
                print(f"⚙️ Converting {path.name} → {new_path.name}")
                df = pd.read_excel(path, engine="xlrd", header=None)
                df.to_excel(new_path, index=False, header=False)
            return str(new_path)
        return str(path)
    
    def _get_excel_engine(self):
        extension = Path(self.file_path).suffix.lower()
        
        if extension == ".xlsx":
            engine = "openpyxl"
            
        elif extension == ".xls":
            engine = "xlrd"
            
        else:
            raise ValueError(f"Unsupported Excel file extension: {extension}")
            
        return engine
    
    def _insert_job_info(self, workbook):
        from datetime import datetime
        import pandas as pd
        ws = workbook.active
        
        ws['B4'] = self.incoming_ticket["Job Number"]
        ws['B5'] = self.incoming_ticket["Job Name"]
        ws['B6'] = self.incoming_ticket["Job Address"]
        ws['B7'] = self.incoming_ticket["Installers"]
        ws['B8'] = self.incoming_ticket["Work Location"]
        
        ws['H4']= self.incoming_ticket["Ticket Number"]
        ws['H5']= pd.to_datetime(datetime.today(), format="%m/%d/%y", errors="coerce")
            
        ws['H6']= pd.to_datetime(self.incoming_ticket["Date"], format="%m/%d/%y", errors="coerce")        
        
        ws['B10']= self.incoming_ticket["Description"]
        
        data = ws.values
        df = pd.DataFrame(data)
        
        workbook.save(f"{self.incoming_ticket['Job Number']} - {self.incoming_ticket['Ticket Number']}.xlsx")
        return df
    
    def write_merged_safe(ws, row, col, value):
        # Check merged ranges
        for merged_range in ws.merged_cells.ranges:
            if ws.cell(row=row, column=col).coordinate in merged_range:
                # Redirect to the top-left of the merged range
                row, col = merged_range.min_row, merged_range.min_col
                break
        ws.cell(row=row, column=col).value = value
        
        
        
        
    
    
        
    

if __name__ =="__main__":
    test_path = "E-ticket Replacement EDITABLE.xlsx"
    
    
    incoming_ticket = {
      'Job Number': '123456',
      'Job Name': 'Test Job',
      'Ticket Number': '00001',
      'Job Address': '1234 Fake Street',
      'Date': '11/10/25',
      'Signature': 'Yes',
      'Type': 'REGULAR',
      'Installers': 'Juan',
      'Work Location': '35TH FLR',
      'Description': 'Fixed pump housing leak',
      'Labor': {'RT': '8', 'OT': '2', 'DT': '0', 'OT DIFF': '0.5', 'DT DIFF': '0'},
      'Materials': [
          {
              'material': 'MAPEI PLANIPREP SC 10LB BAG', 
              'quantity': '3'
          }, 
          {
              'material': 'MAPEI QUICK PATCH 25LB', 
              'quantity': '10'
          }
     ]
    }
    
    wb = ETicketCreator(test_path, incoming_ticket).load_ticket()
    