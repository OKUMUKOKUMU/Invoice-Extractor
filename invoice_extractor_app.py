import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO

def extract_metadata(raw, header_row):
    """Extract invoice no and order no from merged layout (labels in A–C, values in E–G)."""
    invoice_no, order_no = None, None
    for i in range(header_row):
        row = raw.iloc[i]
        row_vals = [str(x).strip() if pd.notna(x) else "" for x in row]

        left_side = " ".join(row_vals[:3]).lower()  # labels in A–C
        right_side = " ".join(row_vals[4:7]).strip()  # values in E–G

        if "invoice" in left_side and not invoice_no:
            match = re.search(r"\d{6,}", right_side)
            invoice_no = match.group(0) if match else right_side

        if "order" in left_side and not order_no:
            match = re.search(r"\d{4,}", right_side)
            order_no = match.group(0) if match else right_side

    return invoice_no, order_no

def process_excel(file_path):
    """Process one Excel file with multiple sheets."""
    all_items = []
    excel_file = pd.ExcelFile(file_path)

    for sheet_name in excel_file.sheet_names:
        raw = excel_file.parse(sheet_name, header=None, dtype=str).fillna("")

        # find header row (where Description and Line Amount exist)
        header_row = None
        for i, row in raw.iterrows():
            row_text = " ".join([str(x) for x in row if str(x).strip()])
            if "description" in row_text.lower() and "line amount" in row_text.lower():
                header_row = i
                break
        if header_row is None:
            continue

        # extract metadata
        invoice_no, order_no = extract_metadata(raw, header_row)

        # load items table
        items_df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        items_df.columns = [str(c).strip() for c in items_df.columns]

        # map columns
        col_map = {}
        for col in items_df.columns:
            low = col.lower()
            if "description" in low: col_map["Description"] = col
            elif low in ["no.", "no"]: col_map["No."] = col
            elif low.startswith("qty"): col_map["Qty"] = col
            elif "uom" in low: col_map["UoM"] = col
            elif "unit price" in low and "excl" in low: col_map["Unit Price Excl. VAT"] = col
            elif low.startswith("vat"): col_map["VAT %"] = col
            elif "line amount" in low and "excl" in low: col_map["Line Amount Excl. VAT"] = col

        needed = ["No.","Description","Qty","UoM","Unit Price Excl. VAT","VAT %","Line Amount Excl. VAT"]
        final_cols = [col_map.get(k) for k in needed if col_map.get(k)]
        if not final_cols:
            continue

        items_clean = items_df[final_cols].copy()
        items_clean = items_clean.rename(columns={v:k for k,v in col_map.items()})

        # remove blanks + KRA QR Code
        items_clean = items_clean.dropna(how="all", subset=needed)
        items_clean = items_clean[~items_clean["Description"].str.contains("KRA QR Code", case=False, na=False)]

        # add metadata
        items_clean.insert(0, "Order No", order_no if order_no else "")
        items_clean.insert(0, "Invoice No", invoice_no if invoice_no else "")
        items_clean.insert(0, "Sheet Name", sheet_name)

        all_items.append(items_clean)

    return pd.concat(all_items, ignore_index=True) if all_items else pd.DataFrame()

def main():
    st.set_page_config(page_title="Invoice Data Extractor", page_icon="📊", layout="wide")
    
    st.title("📊 Invoice Data Extractor")
    st.markdown("Upload Excel files with invoice data to extract and clean the information.")
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose Excel files", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        all_results = []
        processed_files = []
        
        for uploaded_file in uploaded_files:
            with st.spinner(f"Processing {uploaded_file.name}..."):
                try:
                    # Process the file
                    df = process_excel(uploaded_file)
                    
                    if not df.empty:
                        # Enforce consistent column order
                        cols = ["Sheet Name","Invoice No","Order No","No.","Description","Qty","UoM",
                                "Unit Price Excl. VAT","VAT %","Line Amount Excl. VAT"]
                        df = df[cols]
                        
                        # Add to results
                        all_results.append(df)
                        processed_files.append(uploaded_file.name)
                        
                        st.success(f"✅ Successfully processed {uploaded_file.name}")
                        
                        # Show preview
                        with st.expander(f"Preview: {uploaded_file.name}"):
                            st.dataframe(df.head(), use_container_width=True)
                            st.info(f"Total rows: {len(df)}")
                    else:
                        st.warning(f"⚠️ No valid invoice data found in {uploaded_file.name}")
                        
                except Exception as e:
                    st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
        
        # Combine all results and provide download option
        if all_results:
            st.markdown("---")
            st.subheader("📥 Download Results")
            
            combined_df = pd.concat(all_results, ignore_index=True)
            
            # Show combined preview
            st.dataframe(combined_df.head(), use_container_width=True)
            st.info(f"Total rows across all files: {len(combined_df)}")
            
            # Create download button
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Combined_Data')
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download Combined Excel File",
                data=output,
                file_name="combined_invoice_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Individual file downloads
            st.subheader("Individual File Downloads")
            for i, (file_name, df) in enumerate(zip(processed_files, all_results)):
                output_individual = BytesIO()
                with pd.ExcelWriter(output_individual, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Data')
                
                output_individual.seek(0)
                
                base_name = os.path.splitext(file_name)[0]
                st.download_button(
                    label=f"📥 Download {file_name}",
                    data=output_individual,
                    file_name=f"{base_name}_clean_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{i}"
                )

if __name__ == "__main__":
    main()
