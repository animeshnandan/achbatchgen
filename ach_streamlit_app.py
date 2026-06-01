import streamlit as st
import pandas as pd
import csv
import io
from datetime import date, timedelta

st.set_page_config(page_title="ACH Batch CSV Generator", page_icon="🏦", layout="centered")

st.title("🏦 ACH Batch Upload Generator")
st.caption("National Charity Services — Wells Fargo ACH Upload")

uploaded_file = st.file_uploader("Upload Excel Export (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        required_columns = {"Account Name", "Account Number", "Routing Number", "Amount", "Stock # Prefix"}
        missing = required_columns - set(df.columns)

        if missing:
            st.error(f"Missing required column(s): {', '.join(sorted(missing))}")
        else:
            df = df.dropna(subset=list(required_columns), how="all").reset_index(drop=True)

            st.success(f"✅ {len(df)} recipient record(s) loaded")

            # Preview table (mask sensitive data)
            st.subheader("Preview")
            preview = df[["Account Name", "Stock # Prefix", "Amount"]].copy()
            st.dataframe(preview, use_container_width=True)

            # Payment date selector
            default_date = date.today() + timedelta(days=1)
            payment_date_input = st.date_input("Payment Date", value=default_date)
            payment_date = payment_date_input.strftime("%m/%d/%Y")

            st.info(f"Payment Date: **{payment_date}**")

            if st.button("Generate ACH CSV", type="primary"):
                header = [
                    "New Batch Indicator",
                    "ACH Company ID",
                    "ACH Company Name",
                    "ACH Payment Type",
                    "Payment Description",
                    "Payment Date",
                    "Recipient Name",
                    "Recipient Bank ID",
                    "Recipient Account Number",
                    "Recipient Account Type",
                    "Recipient Status",
                    "Send or Receive",
                    "Recipient Amount",
                ]

                rows = [header]

                for _, rec in df.iterrows():
                    batch_row = [
                        "B",
                        "2562600396",
                        "NATIONAL CHARITY",
                        "CCD",
                        rec["Stock # Prefix"],
                        payment_date,
                        "", "", "", "", "", "", "",
                    ]
                    detail_row = [
                        "R",
                        "", "", "", "", "",
                        rec["Account Name"],
                        rec["Routing Number"],
                        rec["Account Number"],
                        "Checking",
                        "Active",
                        "Send",
                        rec["Amount"],
                    ]
                    rows.append(batch_row)
                    rows.append(detail_row)

                # Write to in-memory buffer
                buffer = io.StringIO()
                writer = csv.writer(buffer)
                writer.writerows(rows)
                csv_bytes = buffer.getvalue().encode("utf-8")

                total_amount = df["Amount"].sum()
                st.success(
                    f"✅ Generated {len(df)} recipient(s) — "
                    f"{1 + len(df) * 2} total rows — "
                    f"Total Amount: ${total_amount:,.2f}"
                )

                st.download_button(
                    label="⬇️ Download ACH CSV",
                    data=csv_bytes,
                    file_name=f"ach_upload_{date.today().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                )

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Upload your NCS Excel export to get started.")
