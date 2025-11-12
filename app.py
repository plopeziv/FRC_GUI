# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 11:15:08 2025

@author: plopez
"""

import streamlit as st
import pandas as pd
from pathlib import Path

from data_manager.excel_manager import ExcelManager
from data_manager.ticket_data_service import TicketDataService
from frontend_components.add_ticket_form import show_add_ticket_form

st.title("FRC Ticket GUI")
st.write("Select Your FRC Ticket Listing Excel")

excel_file = st.file_uploader("Choose Excel File", type=["xlsx", "xls"])

if excel_file is not None:
    cwd = Path.cwd()
    file_path = f"{cwd}/{excel_file.name}"
    try:
        manager = ExcelManager(file_path)
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

        if "show_popup" not in st.session_state:
            st.session_state.show_popup = False

        if add_row_clicked:
            st.session_state.show_popup = not st.session_state.show_popup

        if st.session_state.show_popup:
            show_add_ticket_form(manager)

#   LABOR SUMMARY CODE
        st.markdown("### Labor Summary")

        ticket_data_service = TicketDataService(file_path)

        labor_tab1, labor_tab2 = st.tabs(["Labor Summary", "Labor Ticket Details"])

        with labor_tab1:
            st.dataframe(ticket_data_service.labor_summary) 

        with labor_tab2:
            labor_df = ticket_data_service.labor_ticket_summary

            labor_df.index = labor_df.index.astype(str).str.replace(",", "")
            
            st.dataframe(labor_df)

#     MATERIAL SUMMARY CODE
        st.markdown("### Material Summary")

        material_tab1, material_tab2 = st.tabs(["Material Summary", "Material Ticket Details"])
        
        with material_tab1:
            st.dataframe(ticket_data_service.material_summary)


        with material_tab2:
            material_df = ticket_data_service.material_ticket_summary

            material_df.index = material_df.index.astype(str).str.replace(",", "")

            st.dataframe(material_df)

    except Exception as e:
        st.error(f"Error loading file: {e}")

else: st.info("Please select a file to get started")
