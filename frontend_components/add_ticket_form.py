import streamlit as st

def show_add_ticket_form(excel_manager):
    
    with st.form("add_row_form"):
        if "materials_to_add" not in st.session_state:
            st.session_state.materials_to_add = []

        st.markdown("### Add New Row")

        st.markdown("#### Ticket Info")

        (
            ticket_info1_column1, 
            ticket_info1_column2, 
            ticket_info1_column3, 
            ticket_info1_column4
        ) = st.columns([2, 2, 2, 2])

        with ticket_info1_column1:
            ticket_number = st.number_input("Ticket Number", min_value=0, step=1)

        with ticket_info1_column2: 
            date_str = st.text_input("Date (MM/DD/YY)", placeholder="Select a date")

        with ticket_info1_column3: 
            signature = st.selectbox("Signature", ["Yes", "No"])
        
        with ticket_info1_column4: 
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
            dt_dif = st.number_input("DT DIFF", min_value=0, step=1)

        st.markdown("#### Add Material")

        material_qt_col, material_col, material_btn_col= st.columns([1,4, 1.5])

        with material_qt_col:
            material_qt = st.number_input("QT", min_value=0, step=1)

        with material_col:
            selected_material = st.selectbox(
                "Select Material",
                options= excel_manager.materials,
                help="Start typing to search..."
            )
        
        with material_btn_col:
            st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
    
            add_material = st.form_submit_button("➕ Add Material")

        if add_material and material_qt > 0: 
            st.session_state.materials_to_add.append({ "quantity": material_qt, "material": selected_material }) 
            st.success(f"Added {material_qt} × {selected_material}") 
            
            if (len(st.session_state.materials_to_add) > 0): 
                for material_index, material in enumerate(st.session_state.materials_to_add): 
                    new_material_cols = st.columns([1,4,1.5]) 
                    new_material_cols[0].write(material["quantity"]) 
                    new_material_cols[1].write(material["material"])
                    
                    with new_material_cols[2]:
                        canceled = st.form_submit_button("❌ Cancel", key=material_idex)


        st.write("") 
        st.write("") 



        form_column1, form_column2, form_column3 = st.columns([5.75,1.5,1.5])
        with form_column2:
            submitted = st.form_submit_button("✅ Add Row")
        with form_column3:
            canceled = st.form_submit_button("❌ Cancel")

        if canceled:
            st.session_state.show_popup = False
            st.session_state.materials_to_add = []