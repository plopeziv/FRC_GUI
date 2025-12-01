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
        
        self._insert_materials(ws)
        
        self._calculate_ticket_total(ws)
        
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
        
        #Safety Check
        if not active_labor:
            rt_object = labor_object["RT"]
            active_labor = {
                    "REGULAR TIME": {
                        "hours": "0",
                        "rate": rt_object["rate"]
                        }
                }
        
        iterator = iter(active_labor)
        first_key = next(iterator)
        
        # Write first labor entry into row 17
        ws['B17'] = round(float(active_labor[first_key]["hours"]), 1)
        ws['C17'] = first_key
        ws['F17'] = round(float(active_labor[first_key]["rate"]),2)
        ws['I17'] = "=B17*F17"
        
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
           ws[f'I{current_row}'] = f'=B{current_row}*F{current_row}'
           
        ws[f'I{current_row+2}'] = f'=SUM(I17:I{current_row})'
           
        
        
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
                
    def _insert_materials(self, ws):
        material_object = self.incoming_ticket["Materials"]
        
        if not material_object:
           total_material_row = self._find_material_row(ws, "Total Material", column="G")
           ws[f'I{total_material_row}'] = 0
           return
       
        start_row = self._find_material_row(ws) + 1
        
        if start_row is None:
            raise ValueError("Material starting row not found")
            
        # First Material Row
        ws[f'B{start_row}'] = material_object[0]["quantity"]
        ws[f'C{start_row}'] = material_object[0]["material"]
        ws[f'F{start_row}'] = round(float(material_object[0]["sell price"]))
        ws[f'I{start_row}'] = f'=B{start_row} * F{start_row}'
        
        ws.row_dimensions[start_row].height = 25
        
        ws.merge_cells(start_row=start_row, end_row=start_row, 
                       start_column=3, end_column=4)
      
        # Any Material After 1
        current_row = start_row
        
        for material in material_object[1:]:
            current_row += 1
            ws.insert_rows(current_row)
            
            self._copy_and_insert_row(ws, start_row, current_row)
            
            ws[f'B{current_row}'] = material["quantity"]
            ws[f'C{current_row}'] = material["material"]
            ws[f'F{current_row}'] = round(float(material["sell price"]))
            ws[f'I{current_row}'] = f'=B{current_row} * F{current_row}'
            
            ws.row_dimensions[current_row].height = 25
            
            ws.merge_cells(start_row=current_row, end_row=current_row, 
                           start_column=3, end_column=4)
        
        
        # Create the material total summary
        total_material_row = self._find_material_row(ws, "Total Material", column="G")
        ws[f'I{total_material_row}'] = f'=SUM(I{start_row}:I{current_row})'
        
        
        
    def _calculate_ticket_total(self, ws):
        labor_total_row = self._find_material_row(ws, "Total Hours", column='G')
        material_total_row = self._find_material_row(ws, "Total Material", column='G')
        ticket_total_row = self._find_material_row(ws, "Total Ticket", column='G')
        
        formula = f'=I{labor_total_row}+I{material_total_row}'
        ws[f'I{ticket_total_row}'] = formula
        ws.row_dimensions[ticket_total_row].height = 15
        

        
    def _find_material_row(self, ws, search_term="Material Used:", column="A"):
        for row in range(1,50):
            check_value = ws[f'{column}{row}'].value
            if check_value and search_term in str(check_value):
                return row
        
        return None    
    
        
        
        
    
    
        
    

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
          'RT': {'hours':'12', 'rate': '153.15'}, 
          'OT': {'hours':'4', 'rate': '197.70'}, 
          'DT': {'hours':'0', 'rate': '237.72'}, 
          'OT DIFF': {'hours':'0', 'rate':'40.55'}, 
          'DT DIFF': {'hours': '0', 'rate': '80.57'}
          },
      'Materials': [
          {
              'material': 'MAPEI PLANIPREP SC 10LB BAG', 
              'quantity': '3',
              'sell price': '32.50'
          }, 
          {
              'material': 'MAPEI QUICK PATCH 25LB', 
              'quantity': '10',
              'sell price': '36.20'
          },
          {
              'material': 'HEPA SANDER/VAC', 
              'quantity': '1',
              'sell price': '150'
          }, 

     ]
    }
    
    wb = ETicketCreator(test_path, incoming_ticket).load_ticket()
    