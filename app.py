import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="PMO Leitstand", layout="wide")

EXCEL_FILENAME = "PMO_Leitstand_Zielstruktur_Template.xlsx"

@st.cache_data
def load_data_from_excel(file):
    xls = pd.ExcelFile(file)
    goals = pd.read_excel(xls, "Goals")
    persons = pd.read_excel(xls, "Persons")
    partners = pd.read_excel(xls, "Partners")
    return goals, persons, partners

st.title("PMO Projekt-Leitstand")

# --- Excel source handling ---
excel_file = None

if os.path.exists(EXCEL_FILENAME):
    excel_file = EXCEL_FILENAME
    st.success("Excel-Datei aus Repository geladen")
else:
    uploaded = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    if uploaded:
        excel_file = uploaded

if excel_file is None:
    st.warning("Bitte Excel-Datei bereitstellen, um den Leitstand anzuzeigen.")
    st.stop()

# --- Load data ---
goals, persons, partners = load_data_from_excel(excel_file)

# --- Status calculation ---
def calculate_status(df):
    df = df.copy()
    df.loc[df["Goal_Level"] == 4, "Calculated_Status"] = df["Manual_Status"]

    for level in [3, 2, 1]:
        for _, row in df[df["Goal_Level"] == level].iterrows():
            children = df[df["Parent_Goal_ID"] == row["Goal_ID"]]
            statuses = children["Calculated_Status"].dropna()
            if statuses.empty:
                continue
            if all(statuses == "Done"):
                status = "Done"
            elif any(statuses.isin(["At Risk", "Not Started"])):
                status = "At Risk"
            else:
                status = "On Track"
            df.loc[df["Goal_ID"] == row["Goal_ID"], "Calculated_Status"] = status
    return df

goals = calculate_status(goals)

# --- Sidebar ---
levels = {1: "Ebene 1", 2: "Ebene 2", 3: "Ebene 3", 4: "Ebene 4"}
selected_levels = [lvl for lvl in levels if st.sidebar.checkbox(levels[lvl], lvl == 1)]
partner_only = st.sidebar.checkbox("Nur Partnerziele anzeigen")

# --- Filter ---
view = goals[goals["Goal_Level"].isin(selected_levels)]
if partner_only:
    view = view[view["Partner_Involved"] == True]

view = view.merge(
    persons[["Person_ID", "Name"]],
    left_on="Responsible_ID",
    right_on="Person_ID",
    how="left"
)

view = view.merge(
    partners[["Partner_ID", "Partner_Name", "Criticality"]],
    how="left",
    on="Partner_ID"
)

# --- Display ---
for lvl in selected_levels:
    st.subheader(levels[lvl])
    st.dataframe(
        view[view["Goal_Level"] == lvl][[
            "Goal_ID",
            "Goal_Name",
            "Calculated_Status",
            "Planned_Start_Date",
            "Planned_End_Date",
            "Name",
            "Partner_Name",
            "Criticality",
        ]],
        use_container_width=True,
    )
