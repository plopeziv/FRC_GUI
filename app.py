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

        material_tab1, material_tab2 = st.tabs(["Material Summary", "Material Ticket Details"])

        labor_data = [
            [row.tolist()[i] for i in list_indices]
            for row in manager.data_rows
        ]
        
        with material_tab1:
            row5 = manager.dataframe.iloc[5]

            try:
                col_index = row5[row5 == "Structure Material #"].index[0]  # column label
                col_pos = manager.dataframe.columns.get_loc(col_index)     # numeric index

            except IndexError:
                st.write("'Structure Material #' not found in row 5")

            material_summary_df = manager.dataframe.iloc[6:11, col_pos:]

            material_summary_df.set_index(material_summary_df.columns[0], inplace=True)

            row=material_summary_df.loc["Material Counts to Date"]

            zero_columns = row[row == 0].index

            df_filtered = material_summary_df.drop(columns=zero_columns)

            filtered_headers = df_filtered.iloc[0].fillna("").values  # keep as 1D array

            # Remove the header row from the data
            df_filtered = df_filtered[1:]

            # Assign headers
            df_filtered.columns = filtered_headers

            # Display
            st.dataframe(df_filtered)

        with material_tab2:
            list_indices = [*range(0,9)]

            material_headers = [headers[index] for index in list_indices]

            material_data = [
                [row.tolist()[index] for index in list_indices]
                for row in manager.data_rows
            ]

            material_df = pd.DataFrame(material_data, columns=material_headers)

            material_df.set_index("Ticket #", inplace=True)

            material_df = material_df.drop(columns=["Type\n(Regular, Extra)", "Signed", "Labor Sell", "Labor Cost"], errors="ignore")
            
            material_df.index = material_df.index.astype(str).str.replace(",", "")

            material_df["Date"] = material_df["Date"].dt.strftime("%m/%d/%Y")

            st.dataframe(material_df)


    except Exception as e:
        st.error(f"Error loading file: {e}")

else: st.info("Please select a file to get started")
