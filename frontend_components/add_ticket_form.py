import streamlit as st

def show_add_ticket_form(excel_manager):
    
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

        add_material = st.button("➕ Add Material")

    if add_material: 
        if material_qt > 0: 
            _add_or_update_material(selected_material, material_qt)

        else:
            st.error("Please select a quantity before adding a material.") 
        
    if (len(st.session_state.materials_to_add) > 0):
        header_cols = st.columns([0.5, 1, 3, 1])
        header_cols[1].markdown("**Quantity**")
        header_cols[2].markdown("**Material**")

        for material_index, material in enumerate(st.session_state.materials_to_add): 
            new_material_cols = st.columns([0.5,1,3,1]) 
            new_material_cols[1].write(material["quantity"]) 
            new_material_cols[2].write(material["material"])
            
            with new_material_cols[3]:
                remove_material = st.button(f"❌", key={material_index}, on_click=remove_material_callback,
                args=[material["material"]],)
                    



    st.write("") 
    st.write("") 



    form_column1, form_column2, form_column3 = st.columns([5.75,1.5,1.5])
    with form_column2:
        submitted = st.button("✅ Add Row")
    with form_column3:
        canceled = st.button("❌ Cancel")
        if canceled:
            st.session_state.show_popup = False
            st.session_state.materials_to_add = []
            st.rerun()

def _add_or_update_material(material_name, quantity):
    if "materials_to_add" not in st.session_state:
        st.session_state.materials_to_add = []

    for material in st.session_state.materials_to_add:
        if material["material"] == material_name:
            material["quantity"] = quantity
            st.success(f"Updated {material_name} quantity to {material['quantity']}")
            return

    st.session_state.materials_to_add.append({
        "material": material_name,
        "quantity": quantity
    })
    st.success(f"Added {quantity} × {material_name}")

def remove_material_callback(material_name):
    st.session_state.materials_to_add = [
        m for m in st.session_state.materials_to_add
        if m["material"] != material_name
    ]