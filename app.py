# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 11:15:08 2025

@author: plopez
"""

import streamlit as st
import pandas as pd

st.title("FRC Ticket GUI")
st.write("Select Your FRC Ticket Listing Excel")

excel_file = st.file_uploader("Choose Excel File", type=["xlsx", "xls"])

if excel_file is not None:
    try:
        df = pd.read_excel(excel_file)

        st.success("Load Successful!")

    except Exception as e:
        st.error(f"Error loading file: {e}")

else: st.info("Please select a file to get started")
