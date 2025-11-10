import streamlit as st
import re

def show_add_ticket_form(excel_manager):
    
    if "materials_to_add" not in st.session_state:
        st.session_state.materials_to_add = []

    if "form_errors" not in st.session_state:
        st.session_state.form_errors = {}

    st.markdown("### Add New Row")

    st.markdown("#### Ticket Info")

    ticket_info_columns= st.columns([2, 2, 2, 2])

    with ticket_info_columns[0]:
        ticket_number = st.text_input("Ticket Number", placeholder="0")
        if "Ticket Number" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["Ticket Number"]) 

    with ticket_info_columns[1]: 
        date_str = st.text_input("Date (MM/DD/YY)", placeholder="Select a date")
        if "Work Date" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["Work Date"]) 

    with ticket_info_columns[2]: 
        signature = st.selectbox("Signature", ["Yes", "No"])
    
    with ticket_info_columns[3]: 
        ticket_type = st.selectbox("Type", ["REGULAR", "EXTRA", "MISC INSTALL"])

    description = st.text_area("Description", placeholder="Enter ticket details here...")

    st.markdown("#### Add Labor")
    
    labor_columns = st.columns(5)
    with labor_columns[0]:
        regular_time = st.text_input("RT", placeholder="0")
        if "RT" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["RT"]) 
    
    with labor_columns[1]:
        overtime = st.text_input("OT", placeholder="0")
        if "OT" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["OT"]) 

    with labor_columns[2]:
        double_time = st.text_input("DT", placeholder="0")
        if "DT" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["DT"]) 

    with labor_columns[3]:
        ot_dif = st.text_input("OT DIFF", placeholder="0")
        if "OT DIFF" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["OT DIFF"]) 

    with labor_columns[4]:
        dt_dif = st.text_input("DT DIFF", placeholder="0")
        if "DT DIFF" in st.session_state.form_errors:
            st.error(st.session_state.form_errors["DT DIFF"]) 

    st.markdown("#### Add Material")

    material_qt_col, material_col, material_btn_col= st.columns([1,4, 1.5])

    with material_qt_col:
        material_qt = st.text_input("QT", placeholder="0")

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
        try: 
            material_qt_float = float(material_qt)
            if material_qt_float > 0: 
                _add_or_update_material(selected_material, material_qt)

            else:
                st.error("Please select a numeric quantity greater than zero.") 
        except:
            st.error("Quantity must be a valid positive number")
        
    if (len(st.session_state.materials_to_add) > 0):
        header_cols = st.columns([0.5, 1, 3, 1])
        header_cols[1].markdown("**Quantity**")
        header_cols[2].markdown("**Material**")

        for material_index, material in enumerate(st.session_state.materials_to_add): 
            new_material_cols = st.columns([0.5,1,3,1]) 
            new_material_cols[1].write(material["quantity"]) 
            new_material_cols[2].write(material["material"])
            
            with new_material_cols[3]:
                remove_material = st.button(
                    f"❌", 
                    key=f"remove_{material_index}", 
                    on_click=remove_material_item,
                    args=[material["material"]]
                )
                    



    st.write("") 
    st.write("") 


    ticket_data = {
        "Ticket Number": ticket_number.strip(),
        "Date": date_str.strip(),
        "Signature": signature,
        "Type": ticket_type,
        "Description": description.strip(),
        "Labor": {
            "RT": regular_time.strip(),
            "OT": overtime.strip(),
            "DT": double_time.strip(),
            "OT DIFF": ot_dif.strip(),
            "DT DIFF": dt_dif.strip(),
        },
        "Materials": st.session_state.materials_to_add
    }

    st.session_state.form_errors = validate_ticket_form(ticket_data)

    submit_columns = st.columns([5.75,1.5,1.5])
    with submit_columns[1]:
        submitted = st.button("✅ Add Row")

        if submitted:
            if not st.session_state.form_errors:
                st.success("Ticket data is valid!")

    with submit_columns[2]:
        canceled = st.button("❌ Cancel")
        if canceled:
            st.session_state.show_popup = False
            st.session_state.materials_to_add = []
            st.rerun()

    st.write(st.session_state.form_errors)

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

def remove_material_item(material_name):
    st.session_state.materials_to_add = [
        m for m in st.session_state.materials_to_add
        if m["material"] != material_name
    ]

def validate_ticket_form(ticket_data):
    form_errors = {}

    if not ticket_data["Ticket Number"].isdigit() or len(ticket_data["Ticket Number"]) != 5:
        form_errors["Ticket Number"] = "Ticket must be a 5 digit Integer"

    if not re.match(r"^\d{2}/\d{2}/\d{2}$", ticket_data["Date"]):
        form_errors["Work Date"] = "Date must be in MM/DD/YY format"

    form_errors = validate_labor_forms(ticket_data["Labor"], form_errors)

    return form_errors

def validate_labor_forms(labor_object, form_errors):
    for key, value in labor_object.items():
        if value == "":
            continue

        try:
            val = float(value)
            if val < 0:
                form_errors[key] = f"{key} must be positive"
        except ValueError:
            form_errors[key] = f"{key} must be a numeric"

    return form_errors