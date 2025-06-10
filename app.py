import msal
import os
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
def generate_access_token(client_id, scopes, cache_file='token_cache.bin'):
   # Setup token cache
   cache = msal.SerializableTokenCache()
   if os.path.exists(cache_file):
       cache.deserialize(open(cache_file, 'r').read())
   app = msal.PublicClientApplication(
       client_id,
       authority="https://login.microsoftonline.com/common",
       token_cache=cache
   )
   # Try to get token silently (from cache)
   accounts = app.get_accounts()
   if accounts:
       result = app.acquire_token_silent(scopes, account=accounts[0])
   else:
       # Use device code flow if no cached token
       flow = app.initiate_device_flow(scopes=scopes)
       if "user_code" not in flow:
           raise Exception("Failed to start device flow")
       print("Go to", flow["verification_uri"], "and enter code:", flow["user_code"])
       result = app.acquire_token_by_device_flow(flow)
   # Save updated cache
   if cache.has_state_changed:
       with open(cache_file, 'w') as f:
           f.write(cache.serialize())
   if "access_token" in result:
       return result
   else:
       raise Exception("Failed to acquire token: " + str(result))


import os
import requests
# === Configuration ===
APP_ID = '33e576cd-e2db-4a05-8778-71c7f799375f'
SCOPES = ['Files.Read']
FILE_PATH = "Protocol Automation EXCEL Grid.xlsx"  # path in OneDrive
# Local file location
save_location = os.path.join(os.getcwd() , "tmp")
os.makedirs(save_location,exist_ok=True)
LOCAL_EXCEL_FILENAME = "downloaded.xlsx"
# Get token from cache or login
access_token = generate_access_token(APP_ID, SCOPES)
headers = {
   'Authorization': 'Bearer ' + access_token['access_token']
}
# Download URL (OneDrive - self)
download_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{FILE_PATH}:/content"
# Download request
response = requests.get(download_url, headers=headers, verify=False)
# Save the file
if response.status_code == 200:
   local_path = os.path.join(save_location , LOCAL_EXCEL_FILENAME)
   with open(local_path, "wb") as f:
       f.write(response.content)
   print(f"✅ File downloaded and saved to: {local_path}")
else:
   print(f"❌ Download failed: {response.status_code} — {response.text}")


def download_excel_from_onedrive():
   save_location = os.path.join(os.getcwd(), "tmp")
   os.makedirs(save_location, exist_ok=True)  # Ensure tmp/ folder exists
   local_path = os.path.join(save_location, "Protocol_Automation_EXCEL_Grid.xlsx")
   download_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{FILE_PATH}:/content"
   response = requests.get(download_url, headers=headers, verify=False)
   if response.status_code == 200:
       with open(local_path, "wb") as f:
           f.write(response.content)
       print(f"✅ File downloaded and saved to: {local_path}")
   else:
       print(f"❌ Download failed: {response.status_code} {response.text}")
       raise Exception("Download failed!")
   return local_path  # Always return the path



# ===========================
# 2️⃣ Rest of your app imports and code starts here
# ===========================
import streamlit as st
import pandas as pd
import io
# You can now do:
excel_path = download_excel_from_onedrive()
df_bjic_case = pd.read_excel(excel_path, sheet_name="BJIC Case", header=None)
# The rest of your app continues exactly as you already wrote it



import streamlit as st
import pandas as pd
import io

# Load Dependencies Level 1 sheet (mocked for now)
def load_dependencies():
    return {
        'Background': {
            'BJIC': 'BJIC Background Text',
            'MBIC': 'MBIC Background Text',
            'RIC': ''
        },
        'DESIGN': {
            'BJIC': 'BJIC DESIGN Text',
            'MBIC': 'MBIC DESIGN Text',
            'RIC': ''
        },
    }

st.title("Stability Protocol Document Automation")

# Load BJIC Case sheet from backend Excel
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
        "Franchise": read_dropdown_list(row_map["Franchise"]),
        "Study_Purpose": read_dropdown_list(row_map["Study_Purpose"]),
        "Packaging_Configuration": read_dropdown_list(row_map["Packaging_Configuration"]),
        "Active_Ingredients": read_dropdown_list(row_map["Active_Ingredients"]),
        "Regulatory_Classification": read_dropdown_list(row_map["Regulatory_Classification"]),
        "Intended_Market": read_dropdown_list(row_map["Intended_Market"]),
        "Manufacturing_Site": read_dropdown_list(row_map["Manufacturing_Site"]),
        "Packing_Site": read_dropdown_list(row_map["Packing_Site"]),
        "Testing_Site": read_dropdown_list(row_map["Testing_Site"])
    }

dependencies = load_dependencies()

# Start of Form
with st.form("protocol_form"):

    protocol_dev_site = st.selectbox("Protocol Development Site", ["BJIC", "MBIC", "RIC"])

    def multiselect_with_free_text(label, options):
        selected_options = st.multiselect(label, options)
        custom_values = st.text_input(f"Add custom values for {label} (comma-separated)")
        combined_values = selected_options + [x.strip() for x in custom_values.split(",") if x.strip()]
        return combined_values

    if protocol_dev_site == "BJIC":

        st.write("### Fields for BJIC")

        bjic_dropdowns = load_bjic_case_dropdowns()

        business_unit = "Oral Care"
        st.write(f"Business Unit: {business_unit}")

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

        background = dependencies['Background'].get(protocol_dev_site, '')
        st.text_area("Background", background, height=150)

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

        packaging_combined = multiselect_with_free_text("Packaging Configuration", ["0.85 oz PBL", "4.1 oz PBL", "15ml HDPE", "75ml HDPE", "170ml HDPE", "20g ABL", "40g ABL", "90g ABL", "120g ABL", "750 ml Bottles", "1000 ml Bottles", "4 oz Bottles", "8 oz Bottles", "12 oz Bottles", "Other"])

        project_name = st.text_input("Project Name")

        active_combined = multiselect_with_free_text("Active Ingredients", ["NaF", "SnF2", "NaF/SnF2", "MFP", "BSS", "APAP", "DEX", "DOX", "CPM", "DPH", "Other"])

        dose_combined = multiselect_with_free_text("Product Dose Form", ["Dentifrice", "Rinse", "Strips", "Liquid", "Tablet", "Caplet", "Liquicap", "Lozenge", "Spray", "Cream", "Ointment", "Gummy", "Other"])

        reg_combined = multiselect_with_free_text("Regulatory Classification", ["Household Product", "Cosmetic", "Drug", "Dietary Supplement", "Food", "Medical Device", "Other"])

        market_combined = multiselect_with_free_text("Intended Market", ["Greater China", "US", "Canada", "EU", "EMEA", "AMA", "LA", "Other"])

        background = dependencies['Background'].get(protocol_dev_site, '')
        st.text_area("Background", background, height=150)

        manuf_combined = multiselect_with_free_text("Manufacturing Site", ["P&G Beijing Innovation Center (BJIC), China", "P&G XQ plant", "P&G HP plant", "P&G Reading Innovation Centre (RIC), UK", "P&G Gross-Gerau, Germany", "P&G Mason Business Innovation Center, Mason, OH.", "P&G GBO-BS, Iowa City", "P&G GBO-Swing Road", "P&G Naucalpan", "Phoenix", "BestCo", "Trillium", "Other"])

        packing_combined = multiselect_with_free_text("Packing Site", ["P&G Beijing Innovation Center (BJIC), China", "P&G XQ plant", "P&G HP plant", "P&G Reading Innovation Centre (RIC), UK", "P&G Gross-Gerau, Germany", "P&G Mason Business Innovation Center, Mason, OH.", "P&G GBO-BS, Iowa City", "P&G GBO-Swing Road", "P&G Naucalpan", "Phoenix", "BestCo", "Trillium", "Other"])

        placement_combined = multiselect_with_free_text("Placement Site", ["P&G Beijing Innovation Center (BJIC), China", "P&G Mason Business Innovation Center, Mason, OH.", "P&G Reading Innovation Centre (RIC), UK", "Others"])

        testing_combined = multiselect_with_free_text("Testing Site", ["P&G BJIC (Analytical lab, MCO lab, Sensory lab and HOPE lab)", "P&G Mason Business Innovation Center, Mason, OH.", "P&G Reading Innovation Centre (RIC), UK", "Others"])

    # Shared fields B to H

    design = dependencies['DESIGN'].get(protocol_dev_site, '')
    st.text_area("A. DESIGN", design, height=150)

    b_product_manuf_info = st.text_area("B. PRODUCT MANUFACTURING INFORMATION", "[Auto populated text here]")
    c_container_closure = st.text_area("C. CONTAINER / CLOSURE SYSTEM", "[Auto populated text here]")
    d_excursions_other = st.text_area("D. EXCURSIONS and OTHER STUDIES", "[Auto populated text here]")
    e_acceptance_criteria = st.text_area("E. ACCEPTANCE CRITERIA", "[Auto populated text here]")
    f_evaluation_of_data = st.text_area("F. EVALUATION OF DATA", "[Auto populated text here]")
    g_anticipated_reports = st.text_area("G. ANTICIPATED REPORTS", "[Auto populated text here]")
    h_test_methods_specs = st.text_area("H. TEST Methods And Specifications", "[Auto populated text here]")

    submitted = st.form_submit_button("Submit Form")

    if submitted:
        # Build replacements dictionary
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
            "Testing_Site": ", ".join(testing_combined),
            "A_DESIGN": design,
            "B_PRODUCT_MANUFACTURING_INFORMATION": b_product_manuf_info,
            "C_CONTAINER_CLOSURE_SYSTEM": c_container_closure,
            "D_EXCURSIONS_AND_OTHER_STUDIES": d_excursions_other,
            "E_ACCEPTANCE_CRITERIA": e_acceptance_criteria,
            "F_EVALUATION_OF_DATA": f_evaluation_of_data,
            "G_ANTICIPATED_REPORTS": g_anticipated_reports,
            "H_TEST_METHODS_AND_SPECIFICATIONS": h_test_methods_specs,
        }

        # Show replacements dict (for verification)
        st.json(replacements)

        st.success("✅ Form submitted successfully! Replacements dictionary generated.")
