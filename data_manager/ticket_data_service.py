import pandas as pd
from data_manager.excel_manager import ExcelManager
# from excel_manager import ExcelManager

class TicketDataService:
    def __init__(self, file_path=None):
        self.file_path = file_path
        
        self.labor_summary = None
        self.labor_ticket_summary = None
        
        self.material_summary = None
        self.material_ticket_summary = None
        
        self.ticket_listing = None

        self.excel_manager = ExcelManager(self.file_path)
        
        self.excel_manager.load()
        self.job_number = self.excel_manager.job_number
        self.job_name = self.excel_manager.job_name
        self.job_address = self.excel_manager.job_address
        
        self.build_labor_data()
        self.build_material_data()
        self.build_ticket_listing()
        
    def build_labor_data(self):
        self._build_labor_summary()
        self._build_labor_ticket_summary()
        
        # Complete Labor Summary
        sum_of_labor_categories = self.labor_ticket_summary.loc[:, "RT":].sum().to_numpy()
        self.labor_summary.loc["Hours to Date"] = sum_of_labor_categories
        
        cost_to_date_row = self.labor_summary.loc["Hours to Date"] * self.labor_summary.loc["Cost Per Unit (w/Tax)"]
        self.labor_summary.loc["Cost to Date"] = cost_to_date_row
        
        sell_to_date_row = self.labor_summary.loc["Hours to Date"] * self.labor_summary.loc["Sell Per Unit"]
        self.labor_summary.loc["Sell to Date"] = sell_to_date_row
        
        # Complete Labor Ticket Summary
        cost_per_unit = self.labor_summary.loc["Cost Per Unit (w/Tax)", "RT":]
        self.labor_ticket_summary["Labor Cost"] = (
            self.labor_ticket_summary.loc[:, "RT":].multiply(cost_per_unit, axis=1).sum(axis=1)
        )
        
        sell_per_unit = self.labor_summary.loc["Sell Per Unit", "RT":]
        self.labor_ticket_summary["Labor Sell"] = (
            self.labor_ticket_summary.loc[:, "RT":].multiply(sell_per_unit, axis=1).sum(axis=1)
        )
        
    def _build_labor_summary(self):
        df = self.excel_manager.dataframe.iloc[6:12, 11:17]
        df.columns = df.iloc[0]
        df = df[1:]
        df.reset_index(drop=True, inplace=True)
        df.set_index(df.columns[0], inplace=True)
        
        # Calculate Cost Per Unit differences
        rt_value = df.loc["Cost Per Unit (w/Tax)", "RT"]
        ot_value = df.loc["Cost Per Unit (w/Tax)", "OT"]
        dt_value = df.loc["Cost Per Unit (w/Tax)", "DT"]
        df.loc["Cost Per Unit (w/Tax)", "OT DIFF "] = ot_value - rt_value
        df.loc["Cost Per Unit (w/Tax)", "DT DIFF "] = dt_value - rt_value
        
        # Calculate Sell Per Unit differences
        rt_value = df.loc["Sell Per Unit", "RT"]
        ot_value = df.loc["Sell Per Unit", "OT"]
        dt_value = df.loc["Sell Per Unit", "DT"]
        df.loc["Sell Per Unit", "OT DIFF "] = ot_value - rt_value
        df.loc["Sell Per Unit", "DT DIFF "] = dt_value - rt_value
        
        self.labor_summary = df
    
    def _build_labor_ticket_summary(self):
        df = self.excel_manager.dataframe.iloc[13:, :17]
        df = df.drop(df.columns[-6], axis=1)
        
        df.columns = df.iloc[0]
        df = df[1:]
        df.reset_index(drop=True, inplace=True)
        df.set_index("Ticket #", inplace=True)
        
        df = df.drop(columns=df.loc[:, "Material Sell":"Total Cost"].columns)

        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%m/%d/%Y")

        df = df[df.index.notnull()]
        
        self.labor_ticket_summary = df
        
    def build_material_data(self):        
        row5 = self.excel_manager.dataframe.iloc[5]
        
        col_index = row5[row5=="Structure Material #"].index[0]
        col_pos = self.excel_manager.dataframe.columns.get_loc(col_index)
        
        self.material_summary = self._build_material_summary(col_pos)
        
        self.material_ticket_summary = self._build_material_ticket_summary(col_pos)
        
        # Update summary up to date columns
        first_item = '1/4 UNDERLAYMENT 4 X 5"'
        sliced_ticket_summary = self.material_ticket_summary.loc[:, first_item:]
        
        numeric_slice = sliced_ticket_summary.apply(pd.to_numeric, errors='coerce')
        
        numeric_totals = numeric_slice.sum()
        
        self.material_summary.loc["Material Counts to Date", numeric_totals.index] = numeric_totals.values
        
        cost_to_date_row = self.material_summary.loc["Material Counts to Date"] * self.material_summary.loc["Cost Per Unit (w/Tax)"]
        self.material_summary.loc["Cost to Date"] = cost_to_date_row
        
        sell_to_date_row = self.material_summary.loc["Material Counts to Date"] * self.material_summary.loc["Sell Per Unit"]
        self.material_summary.loc["Sell to Date"] = sell_to_date_row
        
        #Drop zero columns
        
        zero_columns = self.material_summary.loc["Material Counts to Date"] == 0

        columns_to_drop = zero_columns[zero_columns].index.tolist()
        
        self.material_summary = self.material_summary.drop(columns=columns_to_drop)
        self.material_ticket_summary = self.material_ticket_summary.drop(columns = columns_to_drop + ["Labor Sell", "Labor Cost"])
        
        # ================================================================
        # SAFETY FIX: If all columns were dropped, restore standard headers
        # ================================================================
        STANDARD_MATERIAL_HEADERS = [
            '1/4 UNDERLAYMENT 4 X 5"',
            "MAPEI PLANIPREP SC 10LB BAG",
            "MAPEI QUICK PATCH 25LB",
            "PRIMER X",
            "SANDING DISC GRIT 36 MEDIUM EA",
        ]
    
        if len(self.material_summary.columns) == 0:
            # Rebuild a safe blank summary with required rows + headers
            self.material_summary = (
                pd.DataFrame(
                    index=[
                        "Cost Per Unit (w/Tax)",
                        "Sell Per Unit",
                        "Material Counts to Date",
                        "Cost to Date",
                        "Sell to Date"
                    ],
                    columns=STANDARD_MATERIAL_HEADERS
                ).fillna(0)
            )
    
        if len(self.material_ticket_summary.columns) == 0:
            # Rebuild blank ticket summary
            self.material_ticket_summary = (
                pd.DataFrame(
                    columns=STANDARD_MATERIAL_HEADERS + ["Material Cost", "Material Sell"]
                ).astype(float).fillna(0)
            )
        # ================================================================
        
        # Complete Material Ticket Summary            
        first_filtered_item = self.material_summary.columns[0]
        
        if first_filtered_item not in self.material_ticket_summary.columns:
            first_filtered_item = self.material_ticket_summary.columns[0]
            
        cost_per_unit = self.material_summary.loc["Cost Per Unit (w/Tax)", first_filtered_item:]
        self.material_ticket_summary["Material Cost"] = (
            self.material_ticket_summary.loc[:, first_filtered_item:].multiply(cost_per_unit, axis=1).sum(axis=1)
        )
        
        sell_per_unit = self.material_summary.loc["Sell Per Unit", first_filtered_item:]
        self.material_ticket_summary["Material Sell"] = (
            self.material_ticket_summary.loc[:, first_filtered_item:].multiply(sell_per_unit, axis=1).sum(axis=1)
        )
        
        return self.material_summary, self.material_ticket_summary
        
        
    def _build_material_summary(self, col_pos):
        df = self.excel_manager.dataframe.iloc[6:12, col_pos:]
        df.columns = df.iloc[0]
        df = df[1:]
        df.reset_index(drop=True, inplace=True)
        df.set_index(df.columns[0], inplace=True)
        
        return df
    
    def _build_material_ticket_summary(self, col_pos):
        df = self.excel_manager.dataframe
        
        list_indices = [*range(0,9), *range(col_pos +1, len(df.columns))]
        
        material_ticket_df = df.iloc[13:, list_indices]
        
        start_pos = list_indices.index(col_pos + 1) 
        
        material_ticket_df = self._create_material_ticket_headers(material_ticket_df, col_pos, start_pos)
        
        material_ticket_df = material_ticket_df[material_ticket_df.index.notnull()]
        
        material_ticket_df["Date"] = pd.to_datetime(material_ticket_df["Date"], errors="coerce").dt.strftime("%m/%d/%Y")
        
        return material_ticket_df
        
    def _create_material_ticket_headers(self, df, col_pos, start_pos):
        material_headers = self.excel_manager.dataframe.iloc[6, col_pos + 1 :].values
        
        df.iloc[0, start_pos:] = material_headers
        
        df.columns = df.iloc[0]
        df = df[1:]
        df.reset_index(drop=True, inplace=True)
        df.set_index("Ticket #", inplace=True)
        
        return df
    
    def build_ticket_listing(self):
        self.ticket_listing = self.labor_ticket_summary.loc[:, : "Labor Cost"]
        
        self.ticket_listing[["Material Sell", "Material Cost"]] = self.material_ticket_summary[["Material Sell", "Material Cost"]]
        
        self.ticket_listing["Total Sell"] = self.ticket_listing["Labor Sell"] + self.ticket_listing["Material Sell"]
        self.ticket_listing["Total Cost"] = self.ticket_listing["Labor Cost"] + self.ticket_listing["Material Cost"]
        
        


if __name__ == "__main__":
    test_path = "test_spread.xlsx"
    
    manager = TicketDataService(test_path)
    
    df1 = manager.labor_ticket_summary
    df2 = manager.material_ticket_summary
    df3 = manager.ticket_listing
    
    