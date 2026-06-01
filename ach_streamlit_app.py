import sys
import csv
import pandas as pd
from datetime import date, timedelta
from pathlib import Path


def generate_ach_csv(input_path: str, output_path: str) -> int:
    """
    Generate an ACH batch CSV from an Excel file.

    Returns the number of recipient records written.
    Raises ValueError if required columns are missing.
    """
    df = pd.read_excel(input_path)

    required_columns = {"Account Name", "Account Number", "Routing Number", "Amount", "Stock # Prefix"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(sorted(missing))}")

    # Drop rows where all required fields are blank
    df = df.dropna(subset=list(required_columns), how="all").reset_index(drop=True)

    payment_date = (date.today() + timedelta(days=1)).strftime("%m/%d/%Y")

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
        # B-row: batch-level fields — Payment Description = Stock # Prefix
        batch_row = [
            "B",
            "2562600396",
            "NATIONAL CHARITY",
            "CCD",
            rec["Stock # Prefix"],
            payment_date,
            "", "", "", "", "", "", "",
        ]
        # R-row: recipient detail fields
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

    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

    return len(df)


def main():
    input_path = "/Users/animesh/Downloads/achtest.xlsx"
    output_path = "/Users/animesh/Downloads/ach_upload_test.csv"

    if not Path(input_path).exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    try:
        count = generate_ach_csv(input_path, output_path)
        payment_date = (date.today() + timedelta(days=1)).strftime("%m/%d/%Y")
        print(f"Done. {count} recipient record(s) written to: {output_path}")
        print(f"Payment date: {payment_date}")
        print(f"Total CSV rows: {1 + count * 2}  (1 header + {count} B-rows + {count} R-rows)")
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
