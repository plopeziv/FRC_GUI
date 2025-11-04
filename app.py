# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 11:15:08 2025

@author: plopez
"""

import streamlit as st
import pandas as pd
from data_manager.excel_manager import ExcelManager

from pathlib import Path

st.title("FRC Ticket GUI")
st.write("Select Your FRC Ticket Listing Excel")

excel_file = st.file_uploader("Choose Excel File", type=["xlsx", "xls"])

if excel_file is not None:
    cwd = Path.cwd()
    try:
        manager = ExcelManager(f"{cwd}/{excel_file.name}")
        manager.load()

        st.success("Load Successful!")


        data_list = [row.tolist()[:11] for row in manager.data_rows]
        
        new_headers = []
        nan_count=1
        headers = manager.headers 
        for h in headers:
            if h == 'nan':
                new_headers.append(f"nan_{nan_count}")
                nan_count += 1
            else:
                new_headers.append(h)
        
        df = pd.DataFrame(data_list, columns=new_headers[:11])
        df = df.loc[:, df.columns.notna()]
        df.set_index("Ticket #", inplace=True)
        df = df.drop(columns=["Type\n(Regular, Extra)"], errors="ignore")

        df.index = df.index.astype(str).str.replace(",", "")

        df["Date"] = df["Date"].dt.strftime("%m/%d/%Y")

        st.markdown("### Ticket Listings")
        st.dataframe(df)

#   LABOR SUMMARY CODE
        st.markdown("### Labor Summary")

        labor_tab1, labor_tab2 = st.tabs(["Labor Summary", "Labor Ticket Details"])

        list_indices = [*range(0,7), *range(12,17)]

        labor_headers = [headers[i] for i in list_indices]

        labor_data = [
            [row.tolist()[i] for i in list_indices]
            for row in manager.data_rows
        ]

        with labor_tab1: 
            data_slice = manager.dataframe.iloc[6:12, 11:17]

            labor_summary_df = pd.DataFrame(
                data_slice
            )

            new_labor_summary_header = labor_summary_df.iloc[0] #grab the first row for the header
            labor_summary_df = labor_summary_df[1:] #take the data less the header row
            labor_summary_df.columns = new_labor_summary_header

            labor_summary_df.set_index("Labor", inplace=True)


            st.dataframe(labor_summary_df)

        with labor_tab2:
            labor_df = pd.DataFrame(labor_data, columns=labor_headers)
            labor_df.set_index("Ticket #", inplace=True)

            labor_df = labor_df.drop(columns=["Type\n(Regular, Extra)", "Signed"], errors="ignore")

            labor_df.index = labor_df.index.astype(str).str.replace(",", "")

            labor_df["Date"] = labor_df["Date"].dt.strftime("%m/%d/%Y")
            
            st.dataframe(labor_df)

#     MATERIAL SUMMARY CODE
        st.markdown("### Material Summary")

    except Exception as e:
        st.error(f"Error loading file: {e}")

else: st.info("Please select a file to get started")
