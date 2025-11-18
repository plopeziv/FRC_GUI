# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 11:15:08 2025

@author: plopez
"""

import streamlit as st
import pandas as pd
from pathlib import Path

import webbrowser
import threading

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

        ticket_data_service = TicketDataService(file_path)


        st.markdown("### Ticket Listings")
        ticket_listing_df = ticket_data_service.ticket_listing

        ticket_listing_df.index = ticket_listing_df.index.astype(str).str.replace(",", "")
        st.dataframe(ticket_listing_df)

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


def open_browser():
    webbrowser.open("http://localhost:8501")

if __name__ == "__main__":
    import streamlit as st
    st.set_page_config(page_title="FRC Ticket GUI")