import streamlit as st
import pandas as pd
import pdfplumber
import re
from datetime import datetime
import io


def extract_ledger_data(pdf_file):
    """Extracts transaction data from an uploaded PDF file object."""
    transactions = []
    party_name = "Unknown Party"

    with pdfplumber.open(pdf_file) as pdf:
        text = ""
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"

    lines = text.split('\n')

    for i, line in enumerate(lines):
        if "Ledger Account" in line and i > 0:
            party_name = lines[i - 1].strip()
            break

    date_pattern = re.compile(r'(\d{1,2}-[A-Za-z]{3}-\d{2})')
    current_date = None

    for line in lines:
        date_match = date_pattern.search(line)
        if date_match:
            current_date = date_match.group(1)
            date_obj = datetime.strptime(current_date, '%d-%b-%y')
            month_year = date_obj.strftime('%B %Y')

        if not current_date:
            continue

        clean_line = line.replace('|', '').replace(',', '').strip()

        if "Opening Balance" in clean_line:
            numbers = re.findall(r'\d+\.\d{2}', clean_line)
            if numbers:
                transactions.append({"Party Name": party_name, "Month": month_year, "Type": "Opening Balance",
                                     "Amount": float(numbers[-1])})

        elif "Sales" in clean_line and "AH/" in clean_line:
            numbers = re.findall(r'\d+\.\d{2}', clean_line)
            if numbers:
                transactions.append(
                    {"Party Name": party_name, "Month": month_year, "Type": "Sale Debit", "Amount": float(numbers[-1])})

        elif "Credit Note" in clean_line and "CN" in clean_line:
            numbers = re.findall(r'\d+\.\d{2}', clean_line)
            if numbers:
                transactions.append({"Party Name": party_name, "Month": month_year, "Type": "Sale Credit",
                                     "Amount": float(numbers[-1])})

        elif "Receipt" in clean_line:
            numbers = re.findall(r'\d+\.\d{2}', clean_line)
            if numbers:
                transactions.append(
                    {"Party Name": party_name, "Month": month_year, "Type": "Receipt", "Amount": float(numbers[-1])})

    return transactions


def create_excel(all_data):
    """Processes the raw data and generates an Excel file in memory."""
    df = pd.DataFrame(all_data)
    final_rows = []

    months_order = ['April 2026', 'May 2026', 'June 2026', 'July 2026', 'August 2026']

    for party, party_data in df.groupby("Party Name"):
        running_outstanding = 0.0
        party_data['Month_Cat'] = pd.Categorical(party_data['Month'], categories=months_order, ordered=True)
        party_data = party_data.sort_values('Month_Cat')

        for month, month_data in party_data.groupby("Month_Cat", observed=True):
            if month_data.empty:
                continue

            opening_bal = month_data[month_data['Type'] == 'Opening Balance']['Amount'].sum()
            sale_debit = month_data[month_data['Type'] == 'Sale Debit']['Amount'].sum()
            sale_credit = month_data[month_data['Type'] == 'Sale Credit']['Amount'].sum()
            receipt = month_data[month_data['Type'] == 'Receipt']['Amount'].sum()

            if opening_bal > 0:
                running_outstanding += opening_bal

            running_outstanding = (running_outstanding + sale_debit) - (sale_credit + receipt)

            final_rows.append({
                "Party Name": party,
                "Month": month,
                "Total sale offline debit": sale_debit,
                "Sale offline credit": sale_credit,
                "Payment receipt": receipt,
                "Total Outstanding": running_outstanding
            })

    final_df = pd.DataFrame(final_rows)
    final_df['Month_Cat'] = pd.Categorical(final_df['Month'], categories=months_order, ordered=True)
    final_df = final_df.sort_values(by=['Month_Cat', 'Party Name']).drop(columns=['Month_Cat'])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Consolidated Ledger')

    return buffer.getvalue()


# --- STATE MANAGEMENT ---
# Initialize session state variables to manage the UI flow
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "processing_complete" not in st.session_state:
    st.session_state.processing_complete = False
if "excel_data" not in st.session_state:
    st.session_state.excel_data = None


def reset_app():
    """Callback function triggered when the download button is clicked."""
    st.session_state.uploader_key += 1  # Changing the key resets the file uploader widget
    st.session_state.processing_complete = False
    st.session_state.excel_data = None


# --- WEB APP INTERFACE ---
st.set_page_config(page_title="Ledger Processor", layout="centered")
st.title("Welcome")
st.markdown("Please Upload Valid PDFs")

# File Uploader with a dynamic key tied to session state
uploaded_files = st.file_uploader(
    "Upload PDF Files",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

# Flow Logic
if uploaded_files:
    if not st.session_state.processing_complete:
        st.info(f"You have uploaded {len(uploaded_files)} PDF(s).")
        st.write("Do you want to process them?")

        if st.button("Yes"):
            all_extracted_data = []
            progress_bar = st.progress(0)

            for i, file in enumerate(uploaded_files):
                try:
                    pdf_data = extract_ledger_data(file)
                    all_extracted_data.extend(pdf_data)
                except Exception as e:
                    st.error(f"Error processing {file.name}: {e}")

                progress_bar.progress((i + 1) / len(uploaded_files))

            if all_extracted_data:
                # Save the processed data to state and flag it as complete
                st.session_state.excel_data = create_excel(all_extracted_data)
                st.session_state.processing_complete = True
                st.rerun()  # Refresh the page to hide the "Yes" button and show the download button
            else:
                st.warning("No transaction data could be extracted from the uploaded PDFs.")

    # Show the success message and download button if processing is done
    if st.session_state.processing_complete and st.session_state.excel_data:
        st.success("✅ Files processed successfully!")
        st.download_button(
            label="⬇️ Download Excel Sheet",
            data=st.session_state.excel_data,
            file_name="Consolidated_Ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=reset_app  # This resets the app immediately after the file downloads
        )
