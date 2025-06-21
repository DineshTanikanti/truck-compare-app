import streamlit as st
import pandas as pd
import re
from dateutil.parser import parse as smart_date
import io

st.set_page_config(page_title="Truck Comparison Tool", layout="wide")
st.title("🚚 Truck Sheet Comparison Tool")

# 🔍 Advanced detection of TRUCK NO, DATE, and INVOICE NO
def detect_columns_by_data(df):
    truck_col = None
    date_col = None
    invoice_col = None
    st.write("\n📑 Detecting columns with advanced logic...")

    for col in df.columns:
        sample = df[col].dropna().astype(str).head(30)

        if truck_col is None and sample.str.contains(r'\b[A-Z0-9]{6,11}\b', case=False).sum() >= 2:
            truck_col = col
            st.success(f"✅ Detected Truck Column: {truck_col}")

        if date_col is None:
            detected = 0
            for val in sample:
                try:
                    smart_date(val, dayfirst=True, fuzzy=True)
                    detected += 1
                except:
                    continue
            if detected >= 2:
                date_col = col
                st.success(f"✅ Detected Date Column: {date_col} using advanced parser")

        if invoice_col is None and sample.str.contains(r'\d{4,}', case=False).sum() >= 2:
            invoice_col = col
            st.success(f"✅ Detected Invoice Column: {invoice_col}")

    if not truck_col:
        st.warning("⚠️ Could not detect truck number column.")
    if not date_col:
        st.warning("⚠️ Could not detect date column.")
    if not invoice_col:
        st.warning("⚠️ Could not detect invoice column.")
    return date_col, truck_col, invoice_col

# 📥 Read and merge all sheets from uploaded Excel files
def read_multiple_files(uploaded_files):
    combined_data = []
    for file in uploaded_files:
        try:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                try:
                    df = pd.read_excel(file, sheet_name=sheet, header=None)
                    df.columns = [f"Col_{i}" for i in range(len(df.columns))]  # temporary generic headers
                    df = df.dropna(how='all')
                    df['Source File'] = file.name
                    df['Sheet Name'] = sheet
                    combined_data.append(df)
                except Exception as e:
                    st.warning(f"⚠️ Could not read sheet '{sheet}' from {file.name}: {e}")
        except Exception as e:
            st.warning(f"⚠️ Could not read file {file.name}: {e}")
    return pd.concat(combined_data, ignore_index=True) if combined_data else pd.DataFrame()

# 📂 Upload section
main_files = st.file_uploader("📁 Upload Main Excel File(s)", type=["xlsx"], accept_multiple_files=True)
bill_files = st.file_uploader("📁 Upload Bill Excel File(s)", type=["xlsx"], accept_multiple_files=True)

if main_files and bill_files:
    main = read_multiple_files(main_files)
    bills = read_multiple_files(bill_files)

    st.subheader("📊 Preview Main File Structure")
    st.write("Main file column names:", list(main.columns))
    st.dataframe(main.head(10))

    st.subheader("📊 Preview Bill File Structure")
    st.write("Bill file column names:", list(bills.columns))
    st.dataframe(bills.head(10))

    try:
        date_col_main, truck_col_main, invoice_col_main = detect_columns_by_data(main)
        date_col_bills, truck_col_bills, invoice_col_bills = detect_columns_by_data(bills)

        if not truck_col_main or not truck_col_bills:
            st.error("❌ Could not detect TRUCK NO in one of the sheets.")
            st.stop()

        def extract_info(df, date_col, truck_col, invoice_col, source_column='Source File'):
            df['__Truck'] = df[truck_col].astype(str).str.extract(r'\b([A-Z0-9]{6,11})\b', expand=False)
            def try_parse_date(val):
                try:
                    return smart_date(val, dayfirst=True, fuzzy=True)
                except:
                    return pd.NaT
            df['__Date'] = df[date_col].astype(str).apply(try_parse_date) if date_col else pd.NaT
            df['__Invoice'] = df[invoice_col].astype(str).str.extract(r'(\d{4,})', expand=False) if invoice_col else ""
            df = df.dropna(subset=['__Truck'])
            df['__Truck'] = df['__Truck'].str.strip().str.upper()
            df['Bill Number'] = df[source_column] + " | " + df['Sheet Name']
            return df

        main = extract_info(main, date_col_main, truck_col_main, invoice_col_main)
        bills = extract_info(bills, date_col_bills, truck_col_bills, invoice_col_bills)

        # Primary match key is DATE + TRUCK, fallback to INVOICE + TRUCK
        main['Key1'] = main['__Date'].dt.date.astype(str) + " | " + main['__Truck']
        bills['Key1'] = bills['__Date'].dt.date.astype(str) + " | " + bills['__Truck']

        main['Key2'] = main['__Invoice'].fillna("") + " | " + main['__Truck']
        bills['Key2'] = bills['__Invoice'].fillna("") + " | " + bills['__Truck']

        bill_lookup1 = bills.set_index('Key1')['Bill Number'].to_dict()
        bill_lookup2 = bills.set_index('Key2')['Bill Number'].to_dict()

        def find_match(row):
            if row['Key1'] in bill_lookup1:
                return "✅ Found", bill_lookup1[row['Key1']]
            elif row['Key2'] in bill_lookup2:
                return "✅ Found via Invoice", bill_lookup2[row['Key2']]
            else:
                return "❌ Missing", "Not Found"

        main[['Status', 'Bill Number']] = main.apply(lambda row: pd.Series(find_match(row)), axis=1)

        result = main[['__Date', '__Invoice', '__Truck', 'Status', 'Bill Number', 'Source File', 'Sheet Name']].rename(columns={
            '__Date': 'Date', '__Truck': 'Truck No', '__Invoice': 'Invoice No'
        })

        with st.expander("🔍 Preview Compared Results"):
            st.dataframe(result)

        output_buffer = io.BytesIO()
        result.to_excel(output_buffer, index=False, engine='openpyxl')
        st.download_button("⬇️ Download Result as Excel", data=output_buffer.getvalue(), file_name="final_comparison.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error during comparison: {e}")