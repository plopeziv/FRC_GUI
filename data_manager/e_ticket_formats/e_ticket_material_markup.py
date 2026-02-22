import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[2]))

from data_manager.e_ticket_creator import ETicketCreator
from config.equipment_lists import EXCLUDED_EQUIPMENT

# from data_manager.e_ticket_creator import ETicketCreator
# from config.equipment_lists import EXCLUDED_EQUIPMENT

class ETicketMaterialMarkup(ETicketCreator):
    def __init__(
        self, 
        file_path=None, 
        incoming_ticket=None,
        template_name="E-ticket Material Markup - PYTHON.xlsx",
        markup= 10
    ):
        super().__init__(file_path, incoming_ticket, template_name)
        self.markup = markup

    def _insert_materials(self, ws):
        
        grouped_materials = self.incoming_ticket["Materials"]
        
        excluded_equipment = EXCLUDED_EQUIPMENT
      
        # ORGANIZE MATERIALS INTO EQUIPMENT AND MATERIALS
        equipment_object = []
        material_object = []
        
        for material in grouped_materials:
            if material["material"] in excluded_equipment:
                equipment_object.append(material)
            else:
                material_object.append(material)
       
        # INSERT MATERIALS
        start_row = self._find_material_row(ws, "Material Used", column="A")
        
        if start_row is None:
            raise ValueError("Material starting row not found")
        
        start_row += 1

        last_material_row = self._insert_line_items(ws, material_object, start_row, "unit cost")            
        
        # Create the material total summary
        total_material_row = self._find_material_row(ws, "Subtotal Material", column="G")
        if total_material_row is None:
            raise ValueError("Subtotal Material row not found")

        multiplier = self.markup / 100
        
        ws[f'I{total_material_row}'] = f'=SUM(I{start_row}:I{last_material_row})'
        
        ws[f'G{total_material_row + 1}'] = f'={multiplier}'
        ws[f'I{total_material_row + 1}'] =f'=I{total_material_row}*{multiplier}'
        
        ws[f'I{total_material_row + 2}'] = f'=I{total_material_row+1}+I{total_material_row}'
        
        ws.merge_cells(start_row=total_material_row+2, end_row=total_material_row+2,
                       start_column=6, end_column=7)

        # INSERT EQUIPMENT
        
        self._insert_equipment(ws, equipment_object)
        
    def _insert_equipment(self, ws, equipment_object):
        start_row = self._find_material_row(ws, "Equipment", column="A")
        
        if start_row is None:
            raise ValueError("Equipment starting row not found")
        
        start_row += 1
        
        last_row = self._insert_line_items(ws, equipment_object, start_row, "sell price")
        
        # Create Equipment Totals
        total_equipment_row = self._find_material_row(ws, "Equipment Total", column="G")
        
        if total_equipment_row is None:
            raise ValueError("Equipment Total row not found")
            
        ws[f'I{total_equipment_row}'] = f'=SUM(I{start_row}:I{last_row})'
        
        
        ws.merge_cells(start_row=total_equipment_row, end_row=total_equipment_row,
                       start_column=6, end_column=7)
        
        ws.cell(row=total_equipment_row, column=6, value="Equipment Total")
        
    def _calculate_ticket_total(self, ws):
        labor_total_row = self._find_material_row(ws, "Total Hours", column='G')
        material_total_row = self._find_material_row(ws, "Material Markup Total", column='F')
        equipment_total_row = self._find_material_row(ws, "Equipment Total", column='F')
        
        ticket_total_row = self._find_material_row(ws, "Total Ticket", column='G')
        
        formula = f'=I{labor_total_row}+I{material_total_row}+I{equipment_total_row}'
        ws[f'I{ticket_total_row}'] = formula
        ws.row_dimensions[ticket_total_row].height = 15

if __name__ =="__main__":
    import os
    
    test_path = os.getcwd()
    
    
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
              'units': 'BG',
              'unit cost': '23.55',
              'sell price': '34.35'
          }, 
          {
              'material': 'MAPEI QUICK PATCH 25LB', 
              'quantity': '10',
              'units': 'BG',
              'unit cost': '23.97',
              'sell price': '34.87'
          },
          {
              'material': 'HEPA SANDER#302 & VAC #701', 
              'quantity': '2',
              'units': 'EA',
              'unit cost': '150',
              'sell price': '150'
          },
          {
              'material': 'TURBO STRIPPER # 203', 
              'quantity': '1',
              'units': 'EA',
              'unit cost': '315',
              'sell price': '315'
          },
          {
              'material': 'HEPA SANDER#302 & VAC #701', 
              'quantity': '2',
              'units': 'EA',
              'unit cost': '150',
              'sell price': '150'
          },
          
     ]
    }
    
    wb = ETicketMaterialMarkup(test_path, incoming_ticket).load_ticket()