import pandas as pd
import csv
import io
import streamlit as st
from datetime import date, timedelta


REQUIRED_COLUMNS = {"Account Name", "Account Number", "Routing Number", "Amount"}


def generate_ach_csv(df: pd.DataFrame) -> tuple[str, list[int], list[int], str]:
    payment_date = (date.today() + timedelta(days=1)).strftime("%m/%d/%Y")

    header = [
        "New Batch Indicator", "ACH Company ID", "ACH Company Name",
        "ACH Payment Type", "Payment Description", "Payment Date",
        "Recipient Name", "Recipient Bank ID", "Recipient Account Number",
        "Recipient Account Type", "Recipient Status", "Send or Receive",
        "Recipient Amount",
    ]

    rows = [header]
    blank_amounts = []
    skipped_rows = []

    for i, rec in df.iterrows():
        # Skip rows missing any of the three key fields
        if pd.isna(rec["Routing Number"]) or pd.isna(rec["Account Number"]) or pd.isna(rec["Account Name"]):
            skipped_rows.append(i + 2)  # +2 for 1-based row + header row
            continue

        if pd.isna(rec["Amount"]):
            blank_amounts.append(i + 2)

        batch_row = [
            "B", "2562600396", "NATIONAL CHARITY",
            "CCD", "NCS", payment_date,
            "", "", "", "", "", "", "",
        ]
        detail_row = [
            "R", "", "", "", "", "",
            rec["Account Name"],
            str(int(rec["Routing Number"])),
            str(int(rec["Account Number"])),
            "Checking", "Active", "Send",
            rec["Amount"],
        ]
        rows.append(batch_row)
        rows.append(detail_row)

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerows(rows)

    return output.getvalue(), blank_amounts, skipped_rows, payment_date


# --- UI ---

st.set_page_config(page_title="ACH Batch Generator", page_icon="🏦")
st.title("ACH Batch CSV Generator")
st.write("Upload your Excel file to generate a formatted ACH upload CSV.")

uploaded_file = st.file_uploader("Choose an Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        st.error(f"Missing required column(s): **{', '.join(sorted(missing))}**")
        st.write("Columns found in file:")
        st.code(", ".join(df.columns))
        st.stop()

    st.success(f"File loaded: **{len(df)} rows** found.")

    with st.expander("Preview source data"):
        st.dataframe(df[["Account Name", "Routing Number", "Account Number", "Amount"]])

    if st.button("Generate ACH CSV", type="primary"):
        csv_content, blank_amounts, skipped_rows, payment_date = generate_ach_csv(df)

        written = (len(csv_content.splitlines()) - 1) // 2
        st.success(f"Generated **{written} records** — payment date: **{payment_date}**")

        if skipped_rows:
            st.warning(
                f"Skipped **{len(skipped_rows)} row(s)** due to missing Routing Number, "
                f"Account Number, or Account Name — source row(s): {skipped_rows}"
            )

        if blank_amounts:
            st.warning(f"Blank amounts found in source row(s): {blank_amounts}")

        st.download_button(
            label="⬇️ Download ach_upload.csv",
            data=csv_content,
            file_name="ach_upload.csv",
            mime="text/csv",
        )