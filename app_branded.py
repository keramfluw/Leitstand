import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="PMO Leitstand", layout="wide")

# ======================
# Branding (Farben & Logos)
# ======================
SAGEMCOM_TURQUOISE = "#00B2A9"
SAGEMCOM_BLUE = "#0077C8"

st.markdown(f"""
<style>
.stApp {{
    background: linear-gradient(180deg, {SAGEMCOM_TURQUOISE}22 0%, #ffffff 40%);
}}
.header-logos {{
    display: flex;
    justify-content: space-between;
    align-items: center;
}}
</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1,1])
with col1:
    if os.path.exists("lackmann.png"):
        st.image("lackmann.png", height=60)
with col2:
    if os.path.exists("sagemcom.png"):
        st.image("sagemcom.png", height=60)

# ======================
# Daten laden
# ======================
EXCEL_FILENAME = "PMO_Leitstand_Zielstruktur_Template.xlsx"

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    return (
        pd.read_excel(xls, "Goals"),
        pd.read_excel(xls, "Persons"),
        pd.read_excel(xls, "Partners"),
    )

excel = EXCEL_FILENAME if os.path.exists(EXCEL_FILENAME) else st.file_uploader("Excel hochladen", type=["xlsx"])
if excel is None:
    st.stop()

goals, persons, partners = load_data(excel)

# ======================
# Statusberechnung
# ======================
def calculate_status(df):
    df = df.copy()
    df.loc[df.Goal_Level == 4, 'Calculated_Status'] = df.Manual_Status
    for lvl in [3, 2, 1]:
        for gid in df[df.Goal_Level == lvl].Goal_ID:
            children = df[df.Parent_Goal_ID == gid]
            if children.empty:
                continue
            s = children.Calculated_Status
            if all(s == 'Done'):
                r = 'Done'
            elif any(s.isin(['At Risk', 'Not Started'])):
                r = 'At Risk'
            else:
                r = 'On Track'
            df.loc[df.Goal_ID == gid, 'Calculated_Status'] = r
    return df

goals = calculate_status(goals)

# ======================
# Ampellogik
# ======================
def ampel(status):
    if status in ['Done', 'On Track']:
        return '🟢'
    if status == 'At Risk':
        return '🔴'
    return '🟡'

st.title("PMO Projekt‑Leitstand")

# ======================
# Sidebar
# ======================
levels = {1: 'Ebene 1', 2: 'Ebene 2', 3: 'Ebene 3', 4: 'Ebene 4'}
selected = [lvl for lvl in levels if st.sidebar.checkbox(levels[lvl], lvl == 1)]

view = goals[goals.Goal_Level.isin(selected)]

# ======================
# Anzeige je Ebene mit Ampel
# ======================
for lvl in selected:
    df_lvl = view[view.Goal_Level == lvl]
    if df_lvl.empty:
        continue
    overall = (
        'At Risk' if any(df_lvl.Calculated_Status == 'At Risk') else
        'On Track' if any(df_lvl.Calculated_Status == 'On Track') else
        'Done'
    )
    st.subheader(f"{levels[lvl]} {ampel(overall)}")
    st.dataframe(
        df_lvl[['Goal_ID', 'Goal_Name', 'Calculated_Status', 'Planned_End_Date']],
        use_container_width=True,
    )
