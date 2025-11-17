import streamlit as st
import pandas as pd
import numpy as np
import os

st.title("📦 Duplicate Customer Analyzer")

st.write("""
Upload multiple DB1 and DB2(PO) Excel files.  
This app merges them, cleans the data, identifies duplicate customers,
and generates a downloadable Excel output.
""")

# ----------------------------
# File Upload Section
# ----------------------------
st.header("📁 Upload Files")

db1_files = st.file_uploader(
    "Upload DB1 Excel files",
    accept_multiple_files=True,
    type=["xlsx", "xls"]
)

db2_files = st.file_uploader(
    "Upload DB2 Excel files",
    accept_multiple_files=True,
    type=["xlsx", "xls"]
)

sheet_name_db1 = "Consolidated"
sheet_name_db2 = "Data"

# ----------------------------
# Process Button
# ----------------------------
if st.button("🚀 Run Analysis"):

    try:
        if not db1_files or not db2_files:
            st.error("Please upload both DB1 and DB2 files before running.")
            st.stop()

        st.info("Reading uploaded files... please wait!")

        # ----------------------------
        # Merge uploaded files
        # ----------------------------
        def merge_uploaded_files(uploaded_files, sheet_name):
            df = pd.DataFrame([])
            for file in uploaded_files:
                df_temp = pd.read_excel(file, sheet_name=sheet_name)
                df = pd.concat([df, df_temp])
            df = df.drop_duplicates()
            return df

        df_db1 = merge_uploaded_files(db1_files, sheet_name_db1)
        df_db2 = merge_uploaded_files(db2_files, sheet_name_db2)

        st.success("Files loaded successfully!")

        # ----------------------------
        # Clean Data
        # ----------------------------
        col_id = "RECEIVER MOBILE NO"

        df_db1["BARCODE NO"] = df_db1["BARCODE NO"].astype(str).str.upper()

        threshold = df_db1.shape[1] / 6
        df_db1_clean = df_db1.dropna(thresh=threshold)
        df_db1_clean[col_id] = df_db1_clean[col_id].astype(str)

        reqd_cols = [
            "Date", "BARCODE NO", "RECEIVER CITY", "RECEIVER PINCODE",
            "RECEIVER NAME", "RECEIVER ADD LINE 1", "RECEIVER ADD LINE 2",
            "RECEIVER ADD LINE 3", "RECEIVER MOBILE NO"
        ]

        duplicated_ph_no = (
            df_db1_clean[df_db1_clean[col_id].duplicated()][col_id]
            .unique()
            .tolist()
        )

        duplicated_customer = (
            df_db1_clean[df_db1_clean[col_id].isin(duplicated_ph_no)]
            .sort_values(by=[col_id, "Date"])[reqd_cols]
        )

        duplicated_customer["PHONE_REPEAT_COUNT"] = (
            duplicated_customer[col_id]
            .groupby(duplicated_customer[col_id])
            .transform("count")
        )

        duplicated_customer = duplicated_customer.sort_values(
            by=["PHONE_REPEAT_COUNT", col_id], ascending=False
        )

        st.subheader("📊 Results Summary")
        st.write(f"**Total duplicated entries:** {duplicated_customer.shape[0]}")
        st.write(f"**Number of unique duplicated customers:** {duplicated_customer[col_id].nunique()}")

        bar_code_duplicated_customer = duplicated_customer["BARCODE NO"].tolist()

        reqd_cols_po = [
            "article-number", "booking-date-time", "event-code",
            "event-description", "non-delivery-reason-description",
            "event-office-name"
        ]

        df_db2_duplicated_customer = df_db2[
            df_db2["article-number"].isin(bar_code_duplicated_customer)
        ][reqd_cols_po]

        df_duplicated_customer = pd.merge(
            duplicated_customer,
            df_db2_duplicated_customer,
            left_on="BARCODE NO",
            right_on="article-number",
            how="left"
        )

        st.subheader("🔍 Preview of Final Output")
        st.dataframe(df_duplicated_customer.head(50))

        st.subheader("📥 Download Output Excel")

        output_file = "duplicated_customers.xlsx"
        df_duplicated_customer.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Excel File",
                data=f,
                file_name="duplicated_customers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("Analysis Completed!")

    except Exception as e:
        st.error("❌ An error occurred. Details below:")
        st.exception(e)
