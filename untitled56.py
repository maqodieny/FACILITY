from io import BytesIO
from urllib.request import urlopen

import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st


# -------------------------------------------------
# 1. Configure the page
# -------------------------------------------------
st.set_page_config(
    page_title="FP Uptake Trend",
    layout="wide"
)


# -------------------------------------------------
# 2. GitHub Excel file location
# -------------------------------------------------
# Replace this with the raw URL of your Excel file.
GITHUB_FILE_URL = (
    "https://raw.githubusercontent.com/"
    "YOUR_USERNAME/YOUR_REPOSITORY/main/Data_for_use.xlsx"
)


# -------------------------------------------------
# 3. Define FP uptake indicators
# -------------------------------------------------
fp_cols = [
    "MOH 711 Client receiving Male condoms",
    "MOH 711 Clients Counselled Natural Family Planning",
    "MOH 711 Clients receiving Female Condoms",
    "MOH 711 Emergency contraceptive pill",
    "MOH 711 Pills Combined oral contraceptive",
    "MOH 711 Pills progestin only",
    "MOH 711 Rev 2020_FP Injections DMPA- IM",
    "MOH 711 Rev 2020_FP Injections DMPA- SC",
    "MOH 711 Rev 2020_IUCD Insertion Hormonal",
    "MOH 711 Rev 2020_IUCD Insertion Non Hormonal",
    "MOH 711 Rev 2020_Implants insertion 1 Rod",
    "MOH 711 Rev 2020_Implants insertion 2 Rod",
    "MOH 711 Rev 2020_Voluntary Surgical Contraception Vasectomy",
    "MOH 711 Rev 2020_Voluntary surgical contraception BTL",
]


# -------------------------------------------------
# 4. Define target facilities
# -------------------------------------------------
target_facilities = [
    "Shafshafey Health Centre",
    "Bp1 Dispensary",
    "Neboi Dispensary",
    "Khadija Dispensary",
    "Iftin Sub-District Hospital",
    "Medina Sub County Hospital",
    "Police Line Health Centre (Garissa)",
    "Utawala Dispensary",
    "Logologo AIC Dispensary",
    "Logologo Model Health Centre (Marsabit South)",
    "Kirimun Health Centre",
    "Shabaa Healthcare (Maralal)",
    "Samburu County Referral Hospital",
    "Kapenguria County Referral Hospital",
    "Marigat Sub County Referral Hospital",
    "Muskut Health Centre",
    "Makongeni Level 4 (Homabay)",
    "Homa Bay County Teaching and Referral Hospital",
    "Got Kojowi Health Centre",
    "Lwanda Kobita Health Center",
    "Enoosaen Health Centre",
    "Transmara West Sub County Hospital",
    "Osupuko Dispensary",
    "Olchorro Health Centre",
    "Siyiapei (AIC) Dispensary",
    "Enabelbel Health Centre",
    "FLC Fountain of life hospital",
    "Narok County Referral Hospital",
    "Kajiado County Referral Hospital",
    "Githurai 44 Health Centre",
    "Kahawa West Health Centre",
    "Muyeye Health Centre (Municipal)",
    "Vishakani Dispensary",
    "Gotani Model Health Centre",
]


# -------------------------------------------------
# 5. Load and prepare the GitHub dataset
# -------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def load_data(file_url):
    # Download the Excel file with a connection timeout.
    with urlopen(file_url, timeout=30) as response:
        excel_bytes = response.read()

    dt = pd.read_excel(
        BytesIO(excel_bytes),
        engine="openpyxl"
    )

    # Clean column names.
    dt.columns = dt.columns.astype(str).str.strip()

    if dt.columns.duplicated().any():
        duplicates = dt.columns[
            dt.columns.duplicated()
        ].tolist()

        raise ValueError(
            "Duplicate column names found: "
            + ", ".join(duplicates)
        )

    # Validate required columns.
    required_cols = ["Facility", "County", "periodid"]

    missing_required = [
        col for col in required_cols
        if col not in dt.columns
    ]

    if missing_required:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing_required)
        )

    # Clean facility and county names.
    dt["Facility"] = (
        dt["Facility"]
        .astype("string")
        .str.strip()
    )

    dt["County"] = (
        dt["County"]
        .astype("string")
        .str.strip()
    )

    # Retain only the target facilities.
    dt = dt.loc[
        dt["Facility"].isin(target_facilities)
    ].copy()

    if dt.empty:
        raise ValueError(
            "No matching target facilities were found in the dataset."
        )

    # Preserve the original calculation:
    # missing FP columns and nonnumeric values count as zero.
    for col in fp_cols:
        if col not in dt.columns:
            dt[col] = 0

    dt[fp_cols] = (
        dt[fp_cols]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
    )

    dt["FP_Uptake"] = dt[fp_cols].sum(axis=1)

    # Convert YYYYMM periods to dates.
    period_strings = (
        dt["periodid"]
        .astype("string")
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
    )

    valid_period_format = period_strings.str.fullmatch(
        r"\d{6}",
        na=False
    )

    dt["Period_Date"] = pd.to_datetime(
        period_strings.where(valid_period_format),
        format="%Y%m",
        errors="coerce"
    )

    dt = dt.dropna(
        subset=["Period_Date"]
    ).copy()

    if dt.empty:
        raise ValueError(
            "No valid reporting periods were found. "
            "Expected periodid format: YYYYMM, for example 202601."
        )

    return dt


try:
    filtered_df = load_data(GITHUB_FILE_URL)

except ImportError:
    st.error(
        "The openpyxl package is missing or needs updating. "
        "Add openpyxl to requirements.txt or install it using: "
        "python -m pip install --upgrade openpyxl"
    )
    st.stop()

except Exception as error:
    st.error(f"Unable to load the GitHub Excel file: {error}")
    st.stop()


# -------------------------------------------------
# 6. Sidebar filters
# -------------------------------------------------
st.sidebar.header("Filters")

county_options = ["All"] + sorted(
    filtered_df["County"]
    .dropna()
    .unique()
    .tolist()
)

if st.session_state.get("selected_county") not in county_options:
    st.session_state["selected_county"] = "All"

selected_county = st.sidebar.selectbox(
    "County",
    options=county_options,
    key="selected_county"
)

county_df = filtered_df.copy()

if selected_county != "All":
    county_df = county_df.loc[
        county_df["County"] == selected_county
    ].copy()

facility_options = ["All"] + sorted(
    county_df["Facility"]
    .dropna()
    .unique()
    .tolist()
)

if st.session_state.get("selected_facility") not in facility_options:
    st.session_state["selected_facility"] = "All"

selected_facility = st.sidebar.selectbox(
    "Facility",
    options=facility_options,
    key="selected_facility"
)

plot_df = county_df.copy()

if selected_facility != "All":
    plot_df = plot_df.loc[
        plot_df["Facility"] == selected_facility
    ].copy()

if plot_df.empty:
    st.info("No data available for the selected filters.")
    st.stop()


# -------------------------------------------------
# 7. Aggregate monthly FP uptake
# -------------------------------------------------
grouped = (
    plot_df.groupby(
        "Period_Date",
        as_index=False
    )["FP_Uptake"]
    .sum()
    .sort_values("Period_Date")
)

grouped["Period_Month"] = (
    grouped["Period_Date"]
    .dt.strftime("%b %Y")
)


# -------------------------------------------------
# 8. Display only the chart
# -------------------------------------------------
fig, ax = plt.subplots(figsize=(16, 7))

ax.plot(
    grouped["Period_Date"],
    grouped["FP_Uptake"],
    marker="o",
    color="#1f77b4",
    linewidth=2
)

ax.set_title(
    f"FP Uptake Trend\n"
    f"County: {selected_county} | Facility: {selected_facility}",
    fontsize=14,
    fontweight="bold"
)

ax.set_xlabel("Reporting Month", fontsize=12)
ax.set_ylabel("FP Uptake (Sum)", fontsize=12)

ax.set_xticks(grouped["Period_Date"])

ax.set_xticklabels(
    grouped["Period_Month"],
    rotation=45,
    ha="right"
)

ax.grid(
    True,
    linestyle="--",
    alpha=0.5
)

fig.tight_layout()

st.pyplot(fig)
plt.close(fig)
