import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
 

class ExcelManager: 
    def __init__(self, file_path=None):
        self.file_path = file_path;
        self.dataframe = None;
        self.materials = [];
        self.headers = []
        self.data_rows = [];
        self.header_row = 13
        
    def load(self):
        if not self.file_path:
            raise ValueError("No file provided!")
            
        path = Path(self.file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
            
        raw_df = pd.read_excel(path, engine=self.get_excel_engine(), header=None)
        
        self.dataframe = raw_df.iloc[:, 6:]

        self.get_headers()
        
        self.get_materials()
        
        self.get_data_rows()
        
        return self.dataframe
    
    def get_excel_engine(self):
        extension = Path(self.file_path).suffix.lower()
        
        if extension == ".xlsx":
            engine = "openpyxl"
            
        elif extension == ".xls":
            engine = "xlrd"
            
        else:
            raise ValueError(f"Unsupported Excel file extension: {extension}")
            
        return engine
    
    def get_headers(self):
        """Extract all headers, combining two different rows for different column ranges."""
        if self.dataframe is None:
            raise ValueError("Excel file not loaded yet. Call load() first.")
            
        header_values = self.dataframe.iloc[self.header_row]
        switch_col = None
        for idx, val in enumerate(header_values):
            if str(val).strip().upper() == "DT DIFF":
                switch_col = idx + 1  # include 'DT DIFF' itself
                break
        
        if switch_col is None:
            raise ValueError("'DT DIFF' not found in main header row.")
        
        # First part: columns 0–16 from the main header row
        main_headers = self.dataframe.iloc[self.header_row, :switch_col]
        
        # Second part: columns 17 onward from row 5
        additional_headers = self.dataframe.iloc[6, switch_col:]
        
        self.headers = [str(h).strip() for h in list(main_headers) + list(additional_headers)]
    
        return self.headers

    
    def get_materials(self, start_col=74-6, header_row = 6):
        if self.dataframe is None:
            raise ValueError("Excel file not loaded yet. Call load() first.")
            
        row_values = self.dataframe.iloc[header_row, start_col:]
        self.materials = [material for material in row_values if pd.notna(material)]
        
        return self.materials
    
    def get_data_rows (self):
        
        for index in range(self.header_row+1, len(self.dataframe)):
            data_row = self.dataframe.iloc[index]
            
            if self.is_row_empty(data_row):
                break
            
            self.data_rows.append(data_row)
            
    def is_row_empty(self, row):
        return not any(bool(cell) for cell in row.fillna(0))
    
    def insert_ticket (self, frc_ticket):
        path = self.ensure_xlsx_copy(self.file_path)
        ticket_data = frc_ticket
        
        wb=load_workbook(path)
        ws = wb.active
        
        last_data_row_index = self.header_row + len(self.data_rows) + 1
        new_row = last_data_row_index + 1
        
        self._insert_ticket_info(ws, new_row, ticket_data)
        
        
        self._insert_labor(ws, new_row, ticket_data["Labor"])
        
        self._insert_materials(ws, new_row, ticket_data["Materials"])
        
        wb.save(self.file_path)
        wb.close()
        
    def _safe_float(self, incoming_value):
        try:
            incoming_value = str(incoming_value).strip().lower()
            return float(incoming_value) if incoming_value not in ("none", "", "nan") else None
        except(ValueError, TypeError):
            return None
        
    def _insert_ticket_info(self, worksheet, ticket_row, ticket_object):
        worksheet.cell(row=ticket_row, column=7, value=pd.to_datetime(ticket_object["Date"], format="%m/%d/%y", errors="coerce"))
        worksheet.cell(row=ticket_row, column=8, value=ticket_object["Signature"])
        worksheet.cell(row=ticket_row, column=9, value=int(ticket_object["Ticket Number"]))
        worksheet.cell(row=ticket_row, column=10, value=ticket_object["Type"])
        worksheet.cell(row=ticket_row, column=11, value=ticket_object["Description"])
        
    def _insert_labor(self, worksheet, ticket_row, labor_object):
        worksheet.cell(row=ticket_row, column=19, value = self._safe_float(labor_object["RT"]))
        worksheet.cell(row=ticket_row, column=20, value = self._safe_float(labor_object["OT"]))
        worksheet.cell(row=ticket_row, column=21, value = self._safe_float(labor_object["DT"]))
        worksheet.cell(row=ticket_row, column=22, value = self._safe_float(labor_object["OT DIFF"]))
        worksheet.cell(row=ticket_row, column=23, value = self._safe_float(labor_object["DT DIFF"]))
        
    def _insert_materials(self, worksheet, material_row, materials):
        for material_object in materials:
            material_name = material_object["material"].strip()
            material_quantity = self._safe_float(material_object["quantity"])
                
            try: 
                column_index = self.headers.index(material_name) + 7
                
                worksheet.cell(row=material_row, column=column_index, value=material_quantity)
                
            except ValueError:
                print(F"{material_name} not found in headers.... Skipped!")

        
        
    def ensure_xlsx_copy(self, path):
        """
        If the working copy path is .xls, convert it to .xlsx for openpyxl.
        Returns the new path (string) compatible with openpyxl.
        """
        path = Path(path)
        if path.suffix.lower() == ".xls":
            new_path = path.with_suffix(".xlsx")  # keep _test in the name already
            if not new_path.exists():
                print(f"⚙️ Converting {path.name} → {new_path.name}")
                df = pd.read_excel(path, engine="xlrd", header=None)
                df.to_excel(new_path, index=False, header=False)
            return str(new_path)
        return str(path)
        
        
    
if __name__ =="__main__":
    test_path = "027386 - DETAILED TICKET LISTING - PYTHON Copy.xlsx"
    
    manager = ExcelManager(test_path)
    
    df = manager.load()
    
    # incoming_ticket = {
    #   'Ticket Number': '00001',
    #   'Date': '11/10/25',
    #   'Signature': 'Yes',
    #   'Type': 'REGULAR',
    #   'Description': 'Fixed pump housing leak',
    #   'Labor': {'RT': '8', 'OT': '2', 'DT': '0', 'OT DIFF': '0.5', 'DT DIFF': '0'},
    #   'Materials': [{'material': 'MAPEI PLANIPREP SC 10LB BAG', 'quantity': '3'}, {'material': 'MAPEI QUICK PATCH 25LB', 'quantity': '10'}]
    # }
    
    # manager.insert_ticket(incoming_ticket)
        