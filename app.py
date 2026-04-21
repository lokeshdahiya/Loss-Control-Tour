import streamlit as st
import pandas as pd

st.title("Tour Analysis Tool")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
location_input = st.text_input("Enter locations (comma separated)")

if uploaded_file and location_input:
    try:
        df = pd.read_excel(uploaded_file)

        # Normalize columns
        df.columns = df.columns.str.strip()

        locations = [loc.strip() for loc in location_input.split(",")]

        filtered = df[df["Location"].isin(locations)]

        if filtered.empty:
            st.warning("No data found for selected locations.")
        else:
            filtered["Completed"] = filtered["Status"].str.lower() == "completed"
            filtered["Non_Completed"] = filtered["Status"].str.lower() != "completed"

            result = filtered.groupby("Executive").agg(
                Total_Tours=("Executive", "count"),
                Completed_Tours=("Completed", "sum"),
                Non_Completed_Tours=("Non_Completed", "sum")
            ).reset_index()

            st.dataframe(result)

            st.download_button(
                label="Download Results",
                data=result.to_csv(index=False),
                file_name="tour_summary.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error: {e}")
