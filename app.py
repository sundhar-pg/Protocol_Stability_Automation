import streamlit as st

import pandas as pd

import os

import requests

import msal




GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

def generate_access_token(client_id, scopes, cache_file='token_cache.bin'):

    cache = msal.SerializableTokenCache()

    if os.path.exists(cache_file):

        cache.deserialize(open(cache_file, 'r').read())

    app = msal.PublicClientApplication(

        client_id,

        authority="https://login.microsoftonline.com/common",

        token_cache=cache

    )

    accounts = app.get_accounts()

    if accounts:

        result = app.acquire_token_silent(scopes, account=accounts[0])

    else:

        flow = app.initiate_device_flow(scopes=scopes)

        if "user_code" not in flow:

            raise Exception("Failed to start device flow")

        print("Go to", flow["verification_uri"], "and enter code:", flow["user_code"])

        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:

        with open(cache_file, 'w') as f:

            f.write(cache.serialize())

    if "access_token" in result:

        return result

    else:

        raise Exception("Failed to acquire token: " + str(result))

def download_excel_from_onedrive():

    APP_ID = '33e576cd-e2db-4a05-8778-71c7f799375f'

    SCOPES = ['Files.Read']

    FILE_PATH = "Protocol Automation EXCEL Grid.xlsx"

    save_location = os.path.join(os.getcwd(), "tmp")

    os.makedirs(save_location, exist_ok=True)

    local_path = os.path.join(save_location, "downloaded.xlsx")

    access_token = generate_access_token(APP_ID, SCOPES)

    headers = {

        'Authorization': 'Bearer ' + access_token['access_token']

    }

    download_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{FILE_PATH}:/content"

    response = requests.get(download_url, headers=headers, verify=False)

    if response.status_code == 200:

        with open(local_path, "wb") as f:

            f.write(response.content)

        print(f"✅ File downloaded and saved to: {local_path}")

    else:

        print(f"❌ Download failed: {response.status_code} {response.text}")

        raise Exception("Download failed!")

    return local_path

def load_dependencies():

    return {

        'Background': {

            'BJIC': 'BJIC Background Text',

            'MBIC': 'MBIC Background Text',

            'RIC': ''

        },

    }

def load_bjic_case_dropdowns():

    excel_path = download_excel_from_onedrive()

    df_bjic_case = pd.read_excel(excel_path, sheet_name="BJIC Case", header=None)

    def read_dropdown_list(row_index):

        raw_value = df_bjic_case.iloc[row_index, 3]

        if isinstance(raw_value, str):

            return [x.strip() for part in raw_value.split("\n") for x in part.split(",") if x.strip()]

        else:

            return []

    row_map = {

        "Franchise": 4,

        "Study_Purpose": 5,

        "Packaging_Configuration": 11,

        "Active_Ingredients": 13,

        "Regulatory_Classification": 15,

        "Intended_Market": 16,

        "Manufacturing_Site": 18,

        "Packing_Site": 19,

        "Testing_Site": 21

    }

    return {

        key: read_dropdown_list(row) for key, row in row_map.items()

    }

def multiselect_with_free_text(label, options):

    col1, col2 = st.columns([1, 1])

    with col1:

        selected = st.multiselect(label, options)

    with col2:

        custom_input = st.text_input(f"Custom {label}", key=label)

    custom_values = [x.strip() for x in custom_input.split(",") if x.strip()]

    return selected + custom_values

# Streamlit App

st.title("Stability Protocol Document Automation")

dependencies = load_dependencies()

with st.form("protocol_form"):

    protocol_dev_site = st.selectbox("Protocol Development Site", ["BJIC", "MBIC", "RIC"])

    if protocol_dev_site == "BJIC":

        bjic_dropdowns = load_bjic_case_dropdowns()

        business_unit_options = ['Business Unit 1', 'Business Unit 2', 'Business Unit 3']  # Replace with actual options
        business_unit = st.selectbox("Business Unit", business_unit_options)

        franchise = st.selectbox("Franchise", bjic_dropdowns["Franchise"])

        study_purpose = st.selectbox("Study Purpose", bjic_dropdowns["Study_Purpose"])

        protocol_number = st.text_input("Stability Protocol Number (Nexus)")

        protocol_number_enovia = st.text_input("Stability Protocol Number (Enovia)")

        product_name_formula = st.text_input("Product Name & Formula #")

        packaging_combined = multiselect_with_free_text("Packaging Configuration", bjic_dropdowns["Packaging_Configuration"])

        project_name = st.text_input("Project Name")

        active_combined = multiselect_with_free_text("Active Ingredients", bjic_dropdowns["Active_Ingredients"])

        dose_combined = multiselect_with_free_text("Product Dose Form", ["Dentifrice"])

        reg_combined = multiselect_with_free_text("Regulatory Classification", bjic_dropdowns["Regulatory_Classification"])

        market_combined = multiselect_with_free_text("Intended Market", bjic_dropdowns["Intended_Market"])

        manuf_combined = multiselect_with_free_text("Manufacturing Site", bjic_dropdowns["Manufacturing_Site"])

        packing_combined = multiselect_with_free_text("Packing Site", bjic_dropdowns["Packing_Site"])

        placement_combined = multiselect_with_free_text("Placement Site", bjic_dropdowns["Manufacturing_Site"])

        testing_combined = multiselect_with_free_text("Testing Site", bjic_dropdowns["Testing_Site"])

    else:

        st.write("### Fields for MBIC / RIC")

        business_unit = st.selectbox("Business Unit", ["OC", "PHC", "Other"])

        franchise_options = {

            'OC': ['Crest', 'Oral-B', 'ProHealth', 'Other'],

            'PHC': ['Vicks', 'Pepto', 'Metamucil', 'Nervive', 'Other'],

            'Other': ['Other']

        }

        franchise = st.selectbox("Franchise", franchise_options[business_unit])

        study_purpose = st.selectbox("Study Purpose", ["Pre-market", "Pre-market for GC", "Pre-market for ROW", "Confirmatory"])

        protocol_number = st.text_input("Stability Protocol Number (Nexus)")

        protocol_number_enovia = st.text_input("Stability Protocol Number (Enovia)")

        product_name_formula = st.text_input("Product Name & Formula #")

        packaging_combined = multiselect_with_free_text("Packaging Configuration", [

            "0.85 oz PBL", "4.1 oz PBL", "15ml HDPE", "75ml HDPE", "170ml HDPE", "20g ABL", "40g ABL", "90g ABL",

            "120g ABL", "750 ml Bottles", "1000 ml Bottles", "4 oz Bottles", "8 oz Bottles", "12 oz Bottles", "Other"

        ])

        project_name = st.text_input("Project Name")

        active_combined = multiselect_with_free_text("Active Ingredients", [

            "NaF", "SnF2", "NaF/SnF2", "MFP", "BSS", "APAP", "DEX", "DOX", "CPM", "DPH", "Other"

        ])

        dose_combined = multiselect_with_free_text("Product Dose Form", [

            "Dentifrice", "Rinse", "Strips", "Liquid", "Tablet", "Caplet", "Liquicap", "Lozenge",

            "Spray", "Cream", "Ointment", "Gummy", "Other"

        ])

        reg_combined = multiselect_with_free_text("Regulatory Classification", [

            "Household Product", "Cosmetic", "Drug", "Dietary Supplement", "Food", "Medical Device", "Other"

        ])

        market_combined = multiselect_with_free_text("Intended Market", [

            "Greater China", "US", "Canada", "EU", "EMEA", "AMA", "LA", "Other"

        ])

        background = dependencies['Background'].get(protocol_dev_site, '')

        st.text_area("Background", background, height=150)

        manuf_combined = multiselect_with_free_text("Manufacturing Site", [

            "P&G Beijing Innovation Center (BJIC), China", "P&G XQ plant", "P&G HP plant",

            "P&G Reading Innovation Centre (RIC), UK", "P&G Gross-Gerau, Germany", "P&G Mason Business Innovation Center, Mason, OH.",

            "P&G GBO-BS, Iowa City", "P&G GBO-Swing Road", "P&G Naucalpan", "Phoenix", "BestCo", "Trillium", "Other"

        ])

        packing_combined = multiselect_with_free_text("Packing Site", manuf_combined)

        placement_combined = multiselect_with_free_text("Placement Site", [

            "P&G Beijing Innovation Center (BJIC), China", "P&G Mason Business Innovation Center, Mason, OH.",

            "P&G Reading Innovation Centre (RIC), UK", "Others"

        ])

        testing_combined = multiselect_with_free_text("Testing Site", [

            "P&G BJIC (Analytical lab, MCO lab, Sensory lab and HOPE lab)", "P&G Mason Business Innovation Center, Mason, OH.",

            "P&G Reading Innovation Centre (RIC), UK", "Others"

        ])

    submitted = st.form_submit_button("Submit Form")

    if submitted:

        replacements = {

            "Protocol_Development_Site": protocol_dev_site,

            "Business_Unit": business_unit,

            "Franchise": franchise,

            "Study_Purpose": study_purpose,

            "Stability_Protocol_Number_Nexus": protocol_number,

            "Stability_Protocol_Number_Enovia": protocol_number_enovia,

            "Product_Name_Formula": product_name_formula,

            "Packaging_Configuration": ", ".join(packaging_combined),

            "Project_Name": project_name,

            "Active_Ingredients": ", ".join(active_combined),

            "Product_Dose_Form": ", ".join(dose_combined),

            "Regulatory_Classification": ", ".join(reg_combined),

            "Intended_Market": ", ".join(market_combined),

            "Background": background,

            "Manufacturing_Site": ", ".join(manuf_combined),

            "Packing_Site": ", ".join(packing_combined),

            "Placement_Site": ", ".join(placement_combined),

            "Testing_Site": ", ".join(testing_combined)

        }

        st.json(replacements)

        st.success("✅ Form submitted successfully! Replacements dictionary generated.")
 

