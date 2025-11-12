from excel_manager import ExcelManager

class TicketDataService:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.labor_summary = None
        self.labor_ticket_summary = None

        self.excel_manager = ExcelManager(self.file_path)
        self.excel_manager.load()
        
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
        
        self.labor_ticket_summary = df


if __name__ == "__main__":
    test_path = "027386 - DETAILED TICKET LISTING - PYTHON Copy.xlsx"
    
    manager = TicketDataService(test_path)
    manager.build_labor_data()
    
    df = manager.labor_summary
    df2 = manager.labor_ticket_summary