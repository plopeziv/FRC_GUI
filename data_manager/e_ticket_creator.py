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
        from openpyxl import load_workbook
        
        if not self.file_path:
            raise ValueError("No file provided!")
            
        path = Path(self.file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        
        path = self._ensure_xlsx_copy()
        wb = load_workbook(path)
        ws = wb.active
        
        self._insert_job_info(ws)
        self._insert_labor(ws)
        
        wb.save(f"{self.incoming_ticket['Job Number']} - {self.incoming_ticket['Ticket Number']}.xlsx")
        
    
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
    
    def _insert_job_info(self, ws):
        from datetime import datetime
        import pandas as pd
        
        ws['B4'] = self.incoming_ticket["Job Number"]
        ws['B5'] = self.incoming_ticket["Job Name"]
        ws['B6'] = self.incoming_ticket["Job Address"]
        ws['B7'] = self.incoming_ticket["Installers"]
        ws['B8'] = self.incoming_ticket["Work Location"]
        
        ws['H4']= self.incoming_ticket["Ticket Number"]
        ws['H5']= pd.to_datetime(datetime.today(), format="%m/%d/%y", errors="coerce")
            
        ws['H6']= pd.to_datetime(self.incoming_ticket["Date"], format="%m/%d/%y", errors="coerce")        
        
        ws['B10']= self.incoming_ticket["Description"]
        
    def _insert_labor(self, ws):
        labor_object = self.incoming_ticket["Labor"]
        
        category_map = {
        "RT": "REGULAR TIME",
        "OT": "OVERTIME",
        "DT": "DOUBLE TIME",
        "OT DIFF": "OVERTIME DIFF",
        "DT DIFF": "DOUBLE TIME DIFF",
        }
        
        # Keep only catagories with hours > 0 and translate their keys
        active_labor = {
            category_map[k]: v for k, v 
            in labor_object.items()
            if float(v["hours"]) > 0}
        
        iterator = iter(active_labor)
        first_key = next(iterator)
        
        # Write first labor entry into row 17
        ws['B17'] = round(float(active_labor[first_key]["hours"]), 1)
        ws['C17'] = first_key
        ws['F17'] = round(float(active_labor[first_key]["rate"]),2)
        ws['I17'] = round(
            float(active_labor[first_key]["hours"]) * float(active_labor[first_key]["rate"]), 2
        )
        
        # Insert Additional Labor Rows
        current_row = 17
        for key in iterator:
           current_row += 1
           ws.insert_rows(current_row)
           
           self._copy_and_insert_row(ws, 17, current_row)
           
           ws.merge_cells(start_row=current_row, end_row=current_row, 
                          start_column=3, end_column=4)
           
           ws[f'B{current_row}'] = round(float(active_labor[key]["hours"]), 1)
           ws[f'C{current_row}'] = key
           ws[f'F{current_row}'] = round(float(active_labor[key]["rate"]),2)
           ws[f'I{current_row}'] = round(
               float(active_labor[key]["hours"]) * float(active_labor[key]["rate"]), 2
           )
        
        
    def _copy_and_insert_row(self, worksheet, src_row, dest_row, max_col=9):
        from copy import copy
        
        for col in range(1, max_col+1):
            src_cell = worksheet.cell(row=src_row, column=col)
            dest_cell = worksheet.cell(row=dest_row, column=col)
            
            dest_cell.value = src_cell.value
            
            if src_cell.has_style:
                dest_cell.font = copy(src_cell.font)
                dest_cell.border = copy(src_cell.border)
                dest_cell.fill = copy(src_cell.fill)
                dest_cell.number_format = src_cell.number_format
                dest_cell.protection = copy(src_cell.protection)
                dest_cell.alignment = copy(src_cell.alignment)
    
        
        
        
    
    
        
    

if __name__ =="__main__":
    test_path = "E-ticket Replacement EDITABLE.xlsx"
    
    
    incoming_ticket = {
      'Job Number': '123456',
      'Job Name': 'Test Job',
      'Ticket Number': '00003',
      'Job Address': '1234 Fake Street',
      'Date': '11/10/25',
      'Signature': 'Yes',
      'Type': 'REGULAR',
      'Installers': 'Juan',
      'Work Location': '35TH FLR',
      'Description': 'Fixed pump housing leak',
      'Labor': {
          'RT': {'hours':'8', 'rate': '157.15'}, 
          'OT': {'hours':'2', 'rate': '197.70'}, 
          'DT': {'hours':'0', 'rate': '237.72'}, 
          'OT DIFF': {'hours':'0.5', 'rate':'40.55'}, 
          'DT DIFF': {'hours': '0', 'rate': '80.57'}
          },
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
    