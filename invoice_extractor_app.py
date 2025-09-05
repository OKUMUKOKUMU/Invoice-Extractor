import streamlit as st
import re
import os
import pandas as pd
import io
from datetime import datetime

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

def process_excel(uploaded_file):
    """Process one Excel file with multiple sheets."""
    all_items = []
    
    try:
        excel_file = pd.ExcelFile(uploaded_file)

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
            items_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
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
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return pd.DataFrame()

def main():
    st.set_page_config(
        page_title="Invoice Data Extractor",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Invoice Data Extractor")
    st.markdown("Upload Excel files containing invoice data to extract and clean the information.")
    
    # Sidebar with instructions
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        1. **Upload** your Excel file(s) containing invoice data
        2. **Process** the files to extract clean data
        3. **Download** the cleaned results as Excel files
        
        ### Expected Format:
        - Invoice/Order numbers in the header area
        - Table with columns like Description, Qty, Unit Price, etc.
        - Multiple sheets supported
        """)
        
        st.header("📈 Features")
        st.markdown("""
        - ✅ Multi-sheet Excel support
        - ✅ Automatic metadata extraction
        - ✅ Column mapping and cleaning
        - ✅ Bulk file processing
        - ✅ Download cleaned data
        """)
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose Excel files to process",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Upload one or more Excel files containing invoice data"
    )
    
    if uploaded_files:
        st.success(f"📁 {len(uploaded_files)} file(s) uploaded successfully!")
        
        # Process button
        if st.button("🔄 Process Files", type="primary"):
            progress_bar = st.progress(0)
            results_container = st.container()
            
            processed_files = []
            
            for i, uploaded_file in enumerate(uploaded_files):
                with st.spinner(f"Processing {uploaded_file.name}..."):
                    # Process the file
                    df = process_excel(uploaded_file)
                    
                    if not df.empty:
                        # Enforce consistent column order
                        cols = ["Sheet Name","Invoice No","Order No","No.","Description","Qty","UoM",
                                "Unit Price Excl. VAT","VAT %","Line Amount Excl. VAT"]
                        df = df[cols]
                        
                        # Store processed data
                        base_name = os.path.splitext(uploaded_file.name)[0]
                        processed_files.append((base_name, df))
                        
                        with results_container:
                            st.success(f"✅ Successfully processed: {uploaded_file.name}")
                            
                            # Show preview
                            with st.expander(f"Preview: {uploaded_file.name} ({len(df)} rows)"):
                                st.dataframe(df, use_container_width=True)
                    else:
                        with results_container:
                            st.warning(f"⚠️ No valid invoice data found in: {uploaded_file.name}")
                
                # Update progress
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            # Download section
            if processed_files:
                st.header("📥 Download Results")
                
                # Individual file downloads
                st.subheader("Individual Files")
                cols = st.columns(min(len(processed_files), 3))
                
                for i, (base_name, df) in enumerate(processed_files):
                    output_buffer = io.BytesIO()
                    df.to_excel(output_buffer, index=False, engine='openpyxl')
                    output_buffer.seek(0)
                    
                    col_idx = i % 3
                    with cols[col_idx]:
                        st.download_button(
                            label=f"📄 {base_name}_clean_data.xlsx",
                            data=output_buffer.getvalue(),
                            file_name=f"{base_name}_clean_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                # Combined download option
                if len(processed_files) > 1:
                    st.subheader("Combined File")
                    
                    # Create combined Excel file with multiple sheets
                    combined_buffer = io.BytesIO()
                    with pd.ExcelWriter(combined_buffer, engine='openpyxl') as writer:
                        for base_name, df in processed_files:
                            sheet_name = base_name[:31]  # Excel sheet name limit
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    combined_buffer.seek(0)
                    
                    st.download_button(
                        label="📦 Download All Files Combined",
                        data=combined_buffer.getvalue(),
                        file_name=f"combined_invoice_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                
                # Summary statistics
                st.header("📊 Processing Summary")
                total_rows = sum(len(df) for _, df in processed_files)
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Files Processed", len(processed_files))
                with col2:
                    st.metric("Total Rows Extracted", total_rows)
                with col3:
                    st.metric("Success Rate", f"{len(processed_files)/len(uploaded_files)*100:.1f}%")
                with col4:
                    st.metric("Files Uploaded", len(uploaded_files))
    
    else:
        st.info("👆 Please upload Excel files to get started!")
        
        # Show sample format
        st.header("📖 Sample Expected Format")
        st.markdown("""
        Your Excel files should have a structure similar to this:
        
        **Header Area (rows 1-10):**
        - Invoice information (Invoice No: 123456)
        - Order information (Order No: 7890)
        
        **Data Table:**
        """)
        
        # Sample data table
        sample_data = {
            "No.": [1, 2, 3],
            "Description": ["Product A", "Product B", "Service C"],
            "Qty": [10, 5, 1],
            "UoM": ["PCS", "KG", "HR"],
            "Unit Price Excl. VAT": [100.00, 250.00, 500.00],
            "VAT %": [16, 16, 16],
            "Line Amount Excl. VAT": [1000.00, 1250.00, 500.00]
        }
        st.dataframe(pd.DataFrame(sample_data), use_container_width=True)

if __name__ == "__main__":
    main()