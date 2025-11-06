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

        col1, col2 = st.columns([5, 1.5])
        with col2:
            add_row_clicked = st.button("➕ Add Row", use_container_width=True)

        # --- Handle form visibility ---
        if "show_popup" not in st.session_state:
            st.session_state.show_popup = False

        if add_row_clicked:
            st.session_state.show_popup = not st.session_state.show_popup

        # --- Popup-like expander ---
        if st.session_state.show_popup:
            with st.form("add_row_form"):
                st.markdown("### Add New Row")

                st.markdown("#### Ticket Info")

                row1_column1, row1_column2, row1_column3, row1_column4 = st.columns([2, 2, 2, 2])

                with row1_column1:
                    ticket_number = st.number_input("Ticket Number", min_value=0, step=1)

                with row1_column2: 
                    date_str = st.text_input("Date (MM/DD/YY)", placeholder="Select a date")

                with row1_column3: 
                    signature = st.selectbox("Signature", ["Yes", "No"])
                
                with row1_column4: 
                    signature = st.selectbox("Type", ["REGULAR", "EXTRA", "MISC INSTALL"])

                description = st.text_area("Description", placeholder="Enter ticket details here...")

                st.markdown("#### Add Labor")
                
                labor_col1, labor_col2, labor_col3, labor_col4, labor_col5 = st.columns(5)
                with labor_col1:
                    regular_time = st.number_input("RT", min_value=0, step=1)
                
                with labor_col2:
                    overtime = st.number_input("OT", min_value=0, step=1)

                with labor_col3:
                    double_time = st.number_input("DT", min_value=0, step=1)

                with labor_col4:
                    ot_dif = st.number_input("OT DIFF", min_value=0, step=1)

                with labor_col5:
                    ot_dif = st.number_input("DT DIFF", min_value=0, step=1)

                st.markdown("#### Add Material")

                form_column1, form_column2, form_column3 = st.columns([5.75,1.5,1.5])
                with form_column2:
                    submitted = st.form_submit_button("✅ Add Row")
                with form_column3:
                    canceled = st.form_submit_button("❌ Cancel")

                if canceled:
                    st.session_state.show_popup = False

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

        row5 = manager.dataframe.iloc[5]

        try:
            col_index = row5[row5 == "Structure Material #"].index[0]  # column label
            col_pos = manager.dataframe.columns.get_loc(col_index)     # numeric index

        except IndexError:
            st.write("'Structure Material #' not found in row 5")

        material_summary_df = manager.dataframe.iloc[6:12, col_pos:]

        material_summary_df.set_index(material_summary_df.columns[0], inplace=True)

        row=material_summary_df.loc["Material Counts to Date"]

        zero_columns = row[row == 0].index
        
        with material_tab1:

            df_filtered = material_summary_df.drop(columns=zero_columns)

            filtered_headers = df_filtered.iloc[0].fillna("").values  # keep as 1D array

            # Remove the header row from the data
            df_filtered = df_filtered[1:]

            # Assign headers
            df_filtered.columns = filtered_headers

            # Display
            st.dataframe(df_filtered)

        with material_tab2:
            list_indices = [*range(0, 9), *range(col_pos + 1, len(manager.dataframe.columns))]
            test_df = manager.dataframe.iloc[13:(14 + len(manager.data_rows)), list_indices]

            filtered_test = test_df.drop(columns=zero_columns)
            filtered_test_headers = filtered_test.iloc[0].fillna("").values
            filtered_test = filtered_test[1:]
            filtered_test.columns = filtered_test_headers

            filtered_test = filtered_test.drop(columns=["Type\n(Regular, Extra)", "Signed", "Labor Sell", "Labor Cost"], errors="ignore") 

            filtered_test.set_index("Ticket #", inplace=True)

            filtered_test.index = filtered_test.index.astype(str).str.replace(",", "")

            filtered_test["Date"] = pd.to_datetime(filtered_test["Date"], errors="coerce")
            filtered_test["Date"] = filtered_test["Date"].dt.strftime("%m/%d/%Y")

            st.dataframe(filtered_test)

    except Exception as e:
        st.error(f"Error loading file: {e}")

else: st.info("Please select a file to get started")
