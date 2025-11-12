import pandas as pd 
import requests
import time
from geopy.distance import geodesic
import streamlit as st

# ------------------- CORE LOGIC ------------------- #
def run_risk_scoring(leads_file, branches_file):
    leads_df = pd.read_excel(leads_file)
    branches_df = pd.read_excel(branches_file)

    leads_df["Latitude"] = None
    leads_df["Longtitude"] = None
    leads_df["Nearest Branch"] = None
    leads_df["Distance to Nearest Branch (miles)"] = None

    def get_coords(postcode):
        if not isinstance(postcode, str) or not postcode:
            return None, None
        url = f"https://api.postcodes.io/postcodes/{postcode.replace(' ', '')}"
        try:
            response = requests.get(url)
            if response.status_code == 200:
                result = response.json()["result"]
                return result["latitude"], result["longitude"]
        except:
            return None, None
        return None, None

    progress_text = "Running risk scoring..."
    progress_bar = st.progress(0)
    total_steps = len(leads_df) * 2
    step = 0

    # Get coords
    for i, row in leads_df.iterrows():
        postcode = row["Post Code"]
        if isinstance(postcode, str) and postcode:
            lat, lon = get_coords(postcode)
            leads_df.at[i, "Latitude"] = lat
            leads_df.at[i, "Longitude"] = lon
            time.sleep(0.1)
        step += 1
        progress_bar.progress(int((step / total_steps) * 100))

    # Nearest branch
    for i, lead in leads_df.iterrows():
        lead_lat = lead["Latitude"]
        lead_lon = lead["Longitude"]

        if pd.isna(lead_lat) or pd.isna(lead_lon):
            leads_df.at[i, "Nearest Branch"] = "N/A"
            leads_df.at[i, "Distance to Nearest Branch (miles)"] = "N/A"
            continue

        lead_location = (lead_lat, lead_lon)
        min_distance = float("inf")
        nearest_branch = None

        for _, branch in branches_df.iterrows():
            branch_location = (branch["Latitude"], branch["Longitude"])
            if pd.isna(branch_location[0]) or pd.isna(branch_location[1]):
                continue
            distance = geodesic(lead_location, branch_location).miles
            if distance < min_distance:
                min_distance = distance
                nearest_branch = branch["Branch Name"]

        leads_df.at[i, "Nearest Branch"] = nearest_branch
        leads_df.at[i, "Distance to Nearest Branch (miles)"] = round(min_distance, 2)

        step += 1
        progress_bar.progress(int((step / total_steps) * 100))

    # Risk calculations (same as before)
    def calculate_ncd_score(years):
        if years == "N/A" or pd.isna(years): return False
        try:
            years = float(years)
            if 0 <= years <= 2: return 10 * 0.25
            elif 2 < years <= 4: return 9 * 0.25
            elif 4 < years <= 6: return 8 * 0.25
            elif 6 < years <= 8: return 7 * 0.25
            elif 8 < years <= 10: return 6 * 0.25
            elif 10 < years <= 12: return 5 * 0.25
            elif 12 < years <= 14: return 4 * 0.25
            elif 14 < years <= 16: return 3 * 0.25
            elif 16 < years <= 18: return 2 * 0.25
            elif years > 18: return 1 * 0.25
        except: return False
        return False

    def calculate_age_score(age):
        if age == "N/A" or pd.isna(age): return False
        try:
            age = float(age)
            if 14 <= age <= 25.2: return 10 * 0.35
            elif 25.2 < age <= 32.4: return 7 * 0.35
            elif 32.4 < age <= 39.6: return 5 * 0.35
            elif 39.6 < age <= 46.8: return 3 * 0.35
            elif 46.8 < age <= 54: return 1 * 0.35
            elif 54 < age <= 61.2: return 2 * 0.35
            elif 61.2 < age <= 68.4: return 4 * 0.35
            elif 68.4 < age <= 75.6: return 6 * 0.35
            elif 75.6 < age <= 82.8: return 8 * 0.35
            elif age > 82.81: return 9 * 0.35
        except: return False
        return False

    def calculate_vehicle_value_score(val):
        if val == "N/A" or pd.isna(val): return False
        try:
            val = float(val)
            if 0 <= val <= 5000: return 1 * 0.15
            elif 5000.2 <= val <= 10000: return 2 * 0.15
            elif 10000 < val <= 15000: return 3 * 0.15
            elif 15000 < val <= 20000: return 4 * 0.15
            elif 20000 < val <= 25000: return 5 * 0.15
            elif 25000 < val <= 30000: return 6 * 0.15
            elif 30000 < val <= 35000: return 7 * 0.15
            elif 35000 < val <= 40000: return 8 * 0.15
            elif 40000 < val <= 45000: return 9 * 0.15
            elif val > 45000.01: return 10 * 0.15
        except: return False
        return False

    def calculate_address_score(distance):
        if distance == "N/A" or pd.isna(distance): return False
        try:
            distance = float(distance)
            if 0 <= distance <= 10: return 1 * 0.25
            elif 10 < distance <= 20: return 2 * 0.25
            elif 20 < distance <= 30: return 3 * 0.25
            elif 30 < distance <= 40: return 4 * 0.25
            elif 40 < distance <= 50: return 5 * 0.25
            elif 50 < distance <= 60: return 6 * 0.25
            elif 60 < distance <= 70: return 7 * 0.25
            elif 70 < distance <= 80: return 8 * 0.25
            elif 80 < distance <= 90: return 9 * 0.25
            elif distance > 90: return 10 * 0.25
        except: return False
        return False

    leads_df["NCDRiskScore"] = leads_df["NCD Years"].apply(calculate_ncd_score)
    leads_df["AgeRiskScore"] = leads_df["Age"].apply(calculate_age_score)
    leads_df["VehicleValueRiskScore"] = leads_df["MotorVehicleValue"].apply(calculate_vehicle_value_score)
    leads_df["AddressBranchScore"] = leads_df["Distance to Nearest Branch (miles)"].apply(calculate_address_score)

    def calculate_total_risk(m, o, q, t):
        if any(score is False for score in [m, o, q, t]): return "FALSE"
        try: return round(m + o + q + t, 2)
        except: return "FALSE"

    leads_df["RiskScore"] = leads_df.apply(
        lambda row: calculate_total_risk(
            row["NCDRiskScore"],
            row["AgeRiskScore"],
            row["VehicleValueRiskScore"],
            row["AddressBranchScore"]
        ),
        axis=1
    )

    progress_bar.empty()

    st.success("✅ Risk Scoring Complete!")
    st.metric(
        "Risk Score Average",
    round(leads_df["RiskScore"].replace("FALSE", pd.NA).dropna().astype(float).mean(), 2)
)

    from io import BytesIO
    buffer = BytesIO()
    leads_df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    st.download_button(
    label="Download Results Excel",
    data=buffer,
    file_name="Risk_Scoring_Results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# ------------------- STREAMLIT UI ------------------- #
st.set_page_config(page_title="PPC Risk Score Data Tool", layout="centered")

st.markdown(
    "<h2 style='color:#cefdc9; background-color:#244c5a; padding:10px; text-align:center;'>PPC Risk Score Data Tool</h2>",
    unsafe_allow_html=True
)

leads_file = st.file_uploader("Upload Client Data (.xlsx)", type=["xlsx"])
branches_file = st.file_uploader("Upload Branch Data (.xlsx)", type=["xlsx"])

if leads_file and branches_file:
    if st.button("Run Risk Scoring"):
        run_risk_scoring(leads_file, branches_file)
else:
    st.info("Please upload both Client Data and Branch Data to begin.")
