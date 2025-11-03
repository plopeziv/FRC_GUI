import pandas as pd
from pathlib import Path
 

class ExcelManager: 
    def __init__(self, file_path=None):
        self.file_path = file_path;
        self.dataframe = None;
        self.materials = [];
        self.data_rows = [];
        self.header_row = 12
        
    def load(self):
        if not self.file_path:
            raise ValueError("No file provided!")
            
        path = Path(self.file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
            
        self.dataframe = pd.read_excel(path, engine=self.get_excel_engine(), header=None)
        
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
    
    def get_materials(self, start_col=74, header_row = 12):
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
    
if __name__ =="__main__":
    test_path = "/Users/plopeziv4/Desktop/GUI Project/027386 - DETAILED TICKET LISTING.xls"
    
    manager = ExcelManager(test_path)
    
    df = manager.load()
    
    rows = manager.data_rows
    print(len(manager.data_rows))
        