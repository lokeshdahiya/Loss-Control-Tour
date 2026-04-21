import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Tour Analysis Tool")

# === Default locations (fixed list) ===
default_locations = [
    "Baghewala EPS",
    "John Rig#14",
    "DR-39 (Deep)",
    "Materials Godown Hamira, Jaisalmer",
    "DANGRI (54 MW)",
    "Dandewala GPC",
    "WOR-DR-39 (DIL)",
    "Hamira Stores",
    "RF Well Logging Workshop",
    "RAMGHAR SOLAR (9 MW)",
    "Explosive Magazine Hamira",
    "CHANDGARH (38 MW)",
    "KOTIYA (27.3 MW)",
    "LUDERVA (13.6 MW)",
    "PATAN (16 MW)",
    "RAMGHAR SOLAR (5 MW)",
    "UCHAWAS (25.2 MW)"
]

# Upload file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Show default locations
st.subheader("Default Locations (Auto Applied)")
st.write(", ".join(default_locations))

# Extra input
extra_locations_input = st.text_input(
    "Add more locations (use semicolon ; to separate, optional)"
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # Clean column names
        df.columns = df.columns.str.strip()

        # Normalize text columns
        df["Location"] = df["Location"].astype(str).str.strip()
        df["Status"] = df["Status"].astype(str).str.strip()
        df["Executive"] = df["Executive"].astype(str).str.strip()

        # Extra locations using semicolon
        extra_locations = [
            loc.strip() for loc in extra_locations_input.split(";") if loc.strip()
        ]

        # Combine all locations
        all_locations = list(set(default_locations + extra_locations))

        # Filter data
        filtered = df[df["Location"].isin(all_locations)].copy()

        if filtered.empty:
            st.warning("No data found for selected locations.")
        else:
            # Status flags
            filtered["Completed"] = filtered["Status"].str.lower() == "completed"
            filtered["Non_Completed"] = filtered["Status"].str.lower() != "completed"

            # Group by Executive
            result = filtered.groupby("Executive").agg(
                Total_Tours=("Executive", "count"),
                Completed_Tours=("Completed", "sum"),
                Non_Completed_Tours=("Non_Completed", "sum")
            ).reset_index()

            # Grand total
            total_row = pd.DataFrame({
                "Executive": ["GRAND TOTAL"],
                "Total_Tours": [result["Total_Tours"].sum()],
                "Completed_Tours": [result["Completed_Tours"].sum()],
                "Non_Completed_Tours": [result["Non_Completed_Tours"].sum()]
            })

            final_result = pd.concat([result, total_row], ignore_index=True)

            # Show results
            st.subheader("Results")
            st.dataframe(final_result)

            # Convert to Excel
            def convert_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Summary')
                return output.getvalue()

            excel_data = convert_to_excel(final_result)

            # Download button
            st.download_button(
                label="Download Excel",
                data=excel_data,
                file_name="tour_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
