

import streamlit as st
import pandas as pd
import pdfplumber
import io
import zipfile
import os
import re

# --- Custom Sheet Name Sanitization ---
def sanitize_sheet_name(name):
    """
    Sanitizes a string to be a valid Excel sheet name.
    Max 31 characters, no / \ ? * [ ] :
    """
    # Replace invalid characters with underscore
    sane_name = re.sub(r'[\\/?*\[\]:]', '_', name)
    # Replace multiple underscores with a single one
    sane_name = re.sub(r'_+', '_', sane_name)
    # Remove leading/trailing underscores
    sane_name = sane_name.strip('_')
    # Truncate to 31 characters
    sane_name = sane_name[:31]
    # Ensure it's not empty after sanitization
    if not sane_name:
        return "Sheet" # Fallback if sanitization results in an empty string
    return sane_name

# --- PDF Data Extraction (THIS IS THE PART YOU WILL CUSTOMIZE HEAVILY) ---
def extract_data_from_pdf(pdf_file_object):
    """
    Extracts item names and prices from a single PDF file using pdfplumber,
    with a strong focus on flexible regex and contextual parsing,
    in addition to attempting table extraction.
    """
    all_extracted_items = []
    
    try:
        with pdfplumber.open(pdf_file_object) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Removed page analysis info to reduce output clutter
                # st.info(f"  Analyzing Page {page_num + 1}...")

                # --- Attempt 1: Table Extraction (still valuable for structured parts) ---
                table_settings_lines = {
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 5, # Increased tolerance
                    "join_tolerance": 5,
                    "edge_min_length": 5,
                    "min_words_vertical": 1,
                    "min_words_horizontal": 1,
                }
                
                table_settings_text_columns = {
                    "vertical_strategy": "text", # For text that forms columns without lines
                    "horizontal_strategy": "lines", # Still look for horizontal lines to delineate rows
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "edge_min_length": 5,
                }

                tables = page.extract_tables(table_settings_lines)
                
                if not tables:
                    # st.info(f"    No tables found using 'lines' strategy on Page {page_num + 1}. Trying 'text' strategy...")
                    tables = page.extract_tables(table_settings_text_columns)

                for table_idx, table_data in enumerate(tables):
                    if table_data and len(table_data) > 1: # Ensure there's a header and at least one row
                        # st.info(f"    Table {table_idx + 1} detected on Page {page_num + 1}. Attempting to parse as price data.")
                        
                        # Clean headers more robustly
                        # Filter out None and empty strings first
                        raw_headers = [str(col).strip() for col in table_data[0] if col is not None and str(col).strip() != '']
                        # Sanitize remaining headers
                        cleaned_headers = [re.sub(r'[^a-zA-Z0-9_]', '', h.replace(' ', '_')).upper() for h in raw_headers]

                        # Handle potential misalignment between headers and data rows
                        num_cols_in_data = len(table_data[1]) if len(table_data) > 1 else 0
                        if len(cleaned_headers) < num_cols_in_data:
                            # Pad headers if data rows have more columns
                            cleaned_headers.extend([f"COL_{i+1}" for i in range(len(cleaned_headers), num_cols_in_data)])
                        elif len(cleaned_headers) > num_cols_in_data:
                            # Truncate headers if there are too many (e.g., split header cell)
                            cleaned_headers = cleaned_headers[:num_cols_in_data]
                        
                        # Handle duplicate column names if they occur after sanitization/truncation
                        seen_headers = {}
                        final_headers = []
                        for header in cleaned_headers:
                            if header in seen_headers:
                                seen_headers[header] += 1
                                final_headers.append(f"{header}_{seen_headers[header]}")
                            else:
                                seen_headers[header] = 1
                                final_headers.append(header)

                        try:
                            df = pd.DataFrame(table_data[1:], columns=final_headers)
                            
                            # Identify potential item and price columns more selectively
                            # Keywords for item names should be descriptive
                            potential_item_cols = [
                                col for col in df.columns if any(
                                    keyword in col for keyword in [
                                        'DESCRIPTION', 'DESCR', 'ITEM', 'PRODUCT', 'MODEL', 'CATNOS', 'CAT_NO',
                                        'RATED_CURRENT', 'TYPE', 'CATEGORY', 'MODULE', 'NAME'
                                    ]
                                )
                            ]
                            # Keywords for price columns
                            potential_price_cols = [
                                col for col in df.columns if any(
                                    keyword in col for keyword in ['MRP', 'PRICE', 'COST', 'UNIT', 'RATE']
                                )
                            ]
                            
                            # Fallback if no specific price column found
                            if not potential_price_cols:
                                for col in df.columns:
                                    # Check if a column contains mostly numeric values that could be prices
                                    if df[col].astype(str).str.contains(r'^\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?\s*$').sum() > len(df) / 2:
                                        potential_price_cols.append(col)
                                        break
                                
                            # Fallback if no specific item column found
                            if not potential_item_cols:
                                for col in df.columns:
                                    # Pick the first non-numeric looking column
                                    if not df[col].astype(str).str.replace('.', '', 1).str.isdigit().all():
                                        potential_item_cols.append(col)
                                        break
                            
                            for index, row in df.iterrows():
                                item_name_parts = []
                                # Concatenate from relevant item columns, prioritizing more specific ones
                                # Order matters: Cat.No first, then description
                                if 'CATNOS' in potential_item_cols and pd.notna(row['CATNOS']) and str(row['CATNOS']).strip():
                                    item_name_parts.append(f"Cat.No: {str(row['CATNOS']).strip()}")
                                if 'DESCRIPTION' in potential_item_cols and pd.notna(row['DESCRIPTION']) and str(row['DESCRIPTION']).strip():
                                    item_name_parts.append(str(row['DESCRIPTION']).strip())
                                # Add other general item columns if still empty or not specific enough
                                for col in potential_item_cols:
                                    if col not in ['CATNOS', 'DESCRIPTION'] and pd.notna(row[col]) and str(row[col]).strip():
                                        item_name_parts.append(str(row[col]).strip())
                                
                                item_name = ' '.join(item_name_parts).strip()
                                
                                price = None
                                for p_col in potential_price_cols:
                                    if pd.notna(row[p_col]):
                                        price_str = str(row[p_col]).replace(',', '').strip()
                                        price_match = re.search(r'(\d+\.?\d*)', price_str)
                                        if price_match:
                                            try:
                                                price = float(price_match.group(1))
                                                break
                                            except ValueError:
                                                price = None

                                if item_name and price is not None:
                                    all_extracted_items.append({"Item Name": item_name, "Price": price})
                                # Removed st.success messages for individual regex matches to reduce output verbosity
                                elif item_name and not price_match and any(pd.notna(row[col]) for col in potential_price_cols):
                                    # Kept this warning as it indicates a potential issue with data quality
                                    st.warning(f"      Item '{item_name}' found, but no valid price extracted from table row: {row.to_dict()}")

                        except Exception as e:
                            st.error(f"      Error processing table {table_idx + 1} on Page {page_num + 1}: {e}")

                # --- Attempt 2: Flexible Regex-based text extraction (for non-table or missed data) ---
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    
                    price_mrp_regex = re.compile(r'(?:MRP\*?\`?\s*\/Unit\s*|MRP\*?\`?\s*|\$?)\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\b')
                    
                    # Corrected to use triple quotes for multi-line string literal
                    cat_no_desc_price_pattern = re.compile(
                        r'''^(?:Cat\.Nos\s*)?(\d{4,5})\s+ # Group 1: Cat.Nos
                        (.+?)\s+                        # Group 2: Description (non-greedy)
                        (\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s* # Group 3: Price
                        (\d+)?$                         # Group 4: Optional pack quantity
                        ''', re.VERBOSE
                    )
                    
                    # Corrected to use triple quotes for multi-line string literal
                    rated_current_complex_pattern = re.compile(
                        r'''^\s*(\d+)\s+ # Group 1: Rated Current
                        (\d{4,5})\s+ # Group 2: 3P Cat.No
                        (\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s+(\d+)\s+ # Group 3: 3P Price, Group 4: 3P Pack
                        (\d{4,5})\s+ # Group 5: 4P Cat.No
                        (\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s+(\d+)\s*$ # Group 6: 4P Price, Group 7: 4P Pack
                        ''', re.VERBOSE
                    )

                    # Corrected to use triple quotes for multi-line string literal
                    general_item_price_line_pattern = re.compile(
                        r'''(?P<cat_no>\b\d{4,5}\b)?\s* # Optional Cat.No
                        (?P<description>.{5,150}?)\s* # Description (5 to 150 chars, non-greedy)
                        (?:MRP\*?\`?\s*\/Unit\s*|MRP\*?\`?\s*|\$)? # Optional price indicators
                        (?P<price>\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s* # Price
                        (?:\d+)?\s* # Optional trailing numbers (like pack quantity)
                        $''', re.VERBOSE
                    )


                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue
                        
                        match_cat_desc_price = cat_no_desc_price_pattern.match(line)
                        if match_cat_desc_price:
                            cat_no = match_cat_desc_price.group(1)
                            description = match_cat_desc_price.group(2).strip()
                            price_str = match_cat_desc_price.group(3)
                            
                            price = None
                            if price_str != "Price available on request.":
                                try:
                                    price = float(price_str.replace(',', ''))
                                except ValueError:
                                    pass

                            item_name = f"{description} (Cat.No: {cat_no})" if cat_no else description
                            if price is not None:
                                all_extracted_items.append({"Item Name": item_name, "Price": price})
                            elif price_str == "Price available on request.":
                                all_extracted_items.append({"Item Name": item_name, "Price": "Price available on request"})
                            continue

                        match_rated_current = rated_current_complex_pattern.match(line)
                        if match_rated_current:
                            rated_current = match_rated_current.group(1)
                            
                            # 3P data
                            cat_no_3p = match_rated_current.group(2)
                            price_3p_str = match_rated_current.group(3)
                            try:
                                price_3p = float(price_3p_str.replace(',', ''))
                                item_name_3p = f"DPX3 MCCB Rated Current {rated_current}A (3P, Cat.No: {cat_no_3p})"
                                all_extracted_items.append({"Item Name": item_name_3p, "Price": price_3p})
                            except ValueError:
                                st.warning(f"      Failed to parse 3P price from line (Rated Current pattern): '{line}'")

                            # 4P data
                            cat_no_4p = match_rated_current.group(5)
                            price_4p_str = match_rated_current.group(6)
                            try:
                                price_4p = float(price_4p_str.replace(',', ''))
                                item_name_4p = f"DPX3 MCCB Rated Current {rated_current}A (4P, Cat.No: {cat_no_4p})"
                                all_extracted_items.append({"Item Name": item_name_4p, "Price": price_4p})
                            except ValueError:
                                st.warning(f"      Failed to parse 4P price from line (Rated Current pattern): '{line}'")
                            continue

                        # NEW: More general item-price line pattern for less structured entries
                        match_general_line = general_item_price_line_pattern.search(line)
                        if match_general_line:
                            try:
                                item_name = match_general_line.group('description').strip()
                                price = float(match_general_line.group('price').replace(',', ''))
                                cat_no = match_general_line.group('cat_no')

                                # Refine item name by attaching Cat.No if available and not already in description
                                if cat_no and f"(Cat.No: {cat_no})" not in item_name:
                                    item_name = f"{item_name} (Cat.No: {cat_no})"
                                
                                # Avoid capturing irrelevant "items" (like numbers or very short strings)
                                if len(item_name) > 5 and not re.fullmatch(r'\d[\d\sX\/]*', item_name): # Avoid just dimensions/numbers
                                    all_extracted_items.append({"Item Name": item_name, "Price": price})
                            except (AttributeError, ValueError): # AttributeError if groups are None, ValueError for float conversion
                                pass # Silently skip if it doesn't parse cleanly

    except Exception as e: # Catch errors from pdfplumber.open() itself (e.g., malformed PDF)
        st.error(f"❌ Error opening or processing PDF: {e}. This PDF might be corrupted or password protected. Skipping direct extraction for this file.")
        return pd.DataFrame() # Return empty DataFrame if opening fails


    if not all_extracted_items:
        st.warning(f"⚠️ No item data found in this PDF based on the current extraction logic.")

    return pd.DataFrame(all_extracted_items)

# --- Streamlit UI ---
st.set_page_config(layout="wide", page_title="PDF Data Extractor (Handles Nested Folders)")

st.title("📄 PDF Item and Price Extractor (Handles Nested Folders)")

st.write("""
Upload a ZIP file that may contain PDF documents directly or within nested folders.
The application will process each PDF, extract item names and prices,
and then save the data into a single Excel workbook with each PDF's data in a separate sheet.
""")

uploaded_zip_file = st.file_uploader(
    "Upload a ZIP file containing PDFs (can have nested folders)",
    type=["zip"],
    help="File size must be 500.0MB or smaller."
)

if uploaded_zip_file is not None:
    st.success("ZIP file uploaded successfully! Processing PDFs...")

    excel_buffer = io.BytesIO()
    excel_writer = pd.ExcelWriter(excel_buffer, engine='openpyxl')

    pdf_count = 0
    processed_files = []
    skipped_files = []

    with zipfile.ZipFile(uploaded_zip_file, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            file_name = file_info.filename

            if file_name.lower().endswith('.pdf') and not file_info.is_dir():
                pdf_count += 1
                st.subheader(f"📂 Processing File: `{file_name}`")
                try:
                    with zip_ref.open(file_name) as pdf_file:
                        pdf_io = io.BytesIO(pdf_file.read())
                        df = extract_data_from_pdf(pdf_io)

                        if not df.empty:
                            # Use the robust sanitize_sheet_name function
                            base_file_name = os.path.splitext(os.path.basename(file_name))[0]
                            # Use only the base file name for sheet, as full path can be very long
                            sheet_name = sanitize_sheet_name(base_file_name)
                            
                            # Ensure unique sheet names if base names clash (e.g., test.pdf in two folders)
                            original_sheet_name = sheet_name
                            counter = 1
                            while sheet_name in excel_writer.sheets:
                                sheet_name = sanitize_sheet_name(f"{original_sheet_name}_{counter}")
                                counter += 1

                            df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
                            processed_files.append(file_name)
                            st.success(f"✅ Extracted data from `{file_name}`. Data saved to sheet: `{sheet_name}`")
                            # Removed st.dataframe(df) to prevent page hanging
                        else:
                            skipped_files.append(file_name)
                            st.warning(f"⚠️ No structured data extracted from `{file_name}`. Skipping sheet creation.")
                except Exception as e: # General catch-all for any unexpected errors during file open/read/excel write
                    st.error(f"❌ An error occurred while handling `{file_name}`: {e}")
                    skipped_files.append(file_name)

    if processed_files or skipped_files:
        st.subheader("🎉 Processing Complete!")
        st.write(f"Total PDFs found in ZIP: **{pdf_count}**")
        if processed_files:
            st.success(f"Successfully extracted data from **{len(processed_files)}** PDF files.")
            with st.expander("See processed files:"):
                for f in processed_files:
                    st.text(f"- {f}")
        if skipped_files:
            st.warning(f"Skipped **{len(skipped_files)}** PDF files due to issues or no data extracted.")
            with st.expander("See skipped files:"):
                for f in skipped_files:
                    st.text(f"- {f}")

        excel_writer.close()
        excel_buffer.seek(0)

        if processed_files:
            st.download_button(
                label="Download Extracted Data as Excel",
                data=excel_buffer,
                file_name="extracted_pdf_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No data was successfully extracted to create an Excel file.")
    else:
        st.warning("No PDF files found in the uploaded ZIP archive, or an error occurred during ZIP processing.")

st.markdown("---")
st.warning("❗ **IMPORTANT**: The `extract_data_from_pdf` function is highly dependent on the exact layout of your PDFs. "
        "The current regex patterns and table settings are tailored to the examples you provided. "
        "You WILL likely need to inspect your PDFs closely and adjust them for optimal accuracy across your entire dataset.")

st.expander("Guidance for Refining `extract_data_from_pdf`").markdown("""
* **Analyze Your PDFs Visually**: Open your PDFs and look for patterns:
    * **Are the "tables" actually drawn with lines**, or are they just columns of text? (`pdfplumber`'s `vertical_strategy` and `horizontal_strategy` are key here: `lines` for drawn tables, `text` for columnar text).
    * **Are there consistent headers**? Even partial ones are helpful.
    * What are the **unique identifiers** for item names and prices? Keywords like "Cat.Nos", "DPX", "MRP", currency symbols, etc.
* **Debug with `pdfplumber`'s visual tools (Highly Recommended)**:
    * **Save individual pages as images with detected lines/text/tables.** This helps immensely in tuning `table_settings` and understanding where text is.
        ```python
        # Example for debugging a specific page (add this to your local code, not Streamlit)
        # import pdfplumber
        # from PIL import Image
        # from io import BytesIO
        # with pdfplumber.open("your_pdf_file.pdf") as pdf:
        #     page = pdf.pages[0] # Adjust page number
        #     img = page.to_image(resolution=150) # Higher resolution for clarity
        #     img.debug_tablefinder(table_settings) # Pass your table settings
        #     # img.save("debug_page_0.png")
        #     # img.show() # Requires Pillow
        ```
* **Refine Regular Expressions (`re.compile`)**:
    * **Test your regex patterns rigorously** with actual text snippets copy-pasted directly from your PDFs. Online regex testers (like regex101.com) are invaluable.
    * Be as specific as possible but also flexible enough for variations.
    * Consider using `re.IGNORECASE` if casing varies.
* **Contextual Parsing for Free-Form Text**:
    * If a price is found, how far *backwards* should you look for its associated item name?
    * What are common delimiters or line breaks that separate one item-price pair from another?
    * Use lists of "product keywords" to help identify item names.
* **Coordinate-based Extraction (Advanced `pdfplumber`)**:
    * If data is always in the same *area* on a page, you can define bounding boxes (`page.crop((x0, y0, x1, y1))`) and then extract text or tables only from that cropped area. This is very powerful for consistent layouts where a "table" is only in a specific part of the page.
* **Handling Multiple Formats**: You might need an `if/else` or a series of `try/except` blocks within `extract_data_from_pdf` to apply different extraction logic based on keywords or patterns found early in a page's text (e.g., "If page contains 'MCCB' and 'Rated Current', use X logic; else if it contains 'equipment and mounting accessories', use Y logic"). This involves classifying the "type" of page/section.
""")
