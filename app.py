import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

# --- Helper Functions ---

def parse_decimal(num_str):
    """
    Converts European format string (1.000,00) to Python Float.
    """
    if not num_str: return 0.0
    clean = num_str.replace('USD', '').strip()
    clean = clean.replace('.', '').replace(',', '.')
    try:
        return float(clean)
    except:
        return 0.0

def is_item_code(word):
    """
    Checks if a word looks like an Item Code.
    Criteria:
    1. Contains at least one digit.
    2. Is NOT an Order Confirmation number (OC + digits).
    3. Is NOT a PO identifier (PO-...).
    """
    if not word or len(word) < 3: return False
    
    # Exclude Order Confirmations
    if word.startswith("OC") and len(word) > 8 and word[2].isdigit():
        return False
        
    # Exclude PO prefixes
    if word.startswith("PO-"):
        return False
        
    # Must contain at least one digit
    has_digit = any(char.isdigit() for char in word)
    
    # Exclude dates
    if re.match(r'\d{2}/\d{2}/\d{2}', word):
        return False

    return has_digit

def load_and_enrich_data(df):
    """
    Loads 'Item Master.csv' and merges it with the extracted PDF data.
    Calculates Lot No = Fig No - Size MM
    """
    master_file = "Item Master.csv"
    
    # Initialize Enrichment Columns if file missing
    enrichment_cols = ["Fig No", "Size MM", "Lot No", "Product Category", "Origin"]
    
    # 1. Check if file exists
    if not os.path.exists(master_file):
        for c in enrichment_cols:
            df[c] = ""
        return df

    try:
        # 2. Read CSV (Try common encodings)
        try:
            df_master = pd.read_csv(master_file, encoding='utf-8')
        except UnicodeDecodeError:
            df_master = pd.read_csv(master_file, encoding='cp1252')
        
        # 3. Define Mapping from CSV Headers to Our Requirements
        col_mapping = {
            "VIR Item Code": "ItemCode",
            "Fig. Number": "Fig No",
            "Size - mm": "Size MM",
            "Product Category": "Product Category",
            "Origin": "Origin",
            "Description": "Master_Desc"
        }

        if "VIR Item Code" not in df_master.columns:
            st.warning("Found 'Item Master.csv' but could not find 'VIR Item Code' column.")
            return df

        # Rename columns
        df_master = df_master.rename(columns=col_mapping)
        
        # Select relevant columns
        desired_fields = ["ItemCode", "Fig No", "Size MM", "Product Category", "Origin", "Master_Desc"]
        available_fields = [c for c in desired_fields if c in df_master.columns]
        
        df_master_subset = df_master[available_fields].copy()

        # 4. Merge Data (Left Join)
        df_merged = pd.merge(df, df_master_subset, on="ItemCode", how="left")
        
        # 5. Logic: Use Master Description if available
        if "Master_Desc" in df_merged.columns:
            df_merged["Item Desc"] = df_merged["Master_Desc"].fillna(df_merged["Item Desc"])
            df_merged = df_merged.drop(columns=["Master_Desc"])
            
        # 6. --- NEW LOGIC: Calculate Lot No ---
        
        # Ensure Fig No and Size MM are strings and clean up float decimals (e.g. 200.0 -> 200)
        # We handle cases where columns might not exist or are NaN
        if "Fig No" in df_merged.columns:
            df_merged["Fig No"] = df_merged["Fig No"].fillna("").astype(str).str.replace(r'\.0$', '', regex=True)
        else:
            df_merged["Fig No"] = ""

        if "Size MM" in df_merged.columns:
            df_merged["Size MM"] = df_merged["Size MM"].fillna("").astype(str).str.replace(r'\.0$', '', regex=True)
        else:
            df_merged["Size MM"] = ""
            
        def calculate_lot(row):
            fig = row.get("Fig No", "").strip()
            size = row.get("Size MM", "").strip()
            # If both exist, combine them
            if fig and size:
                return f"{fig}-{size}"
            return ""

        df_merged["Lot No"] = df_merged.apply(calculate_lot, axis=1)

        # 7. Fill remaining missing columns (Category, Origin)
        for c in ["Product Category", "Origin"]:
            if c in df_merged.columns:
                df_merged[c] = df_merged[c].fillna("")
            else:
                df_merged[c] = ""
                
        return df_merged

    except Exception as e:
        st.warning(f"Error reading Item Master file: {e}")
        return df

def calculate_weights(df, hs_weight_map):
    """
    Calculates 'Unit Weight' and 'Total Weight'.
    """
    if df.empty:
        return df
        
    hs_qty_sum = df.groupby("HSCode")["Qty"].sum().to_dict()
    
    def get_unit_weight(row):
        hs_code = row["HSCode"]
        total_summary_weight = hs_weight_map.get(hs_code, 0.0)
        total_hs_qty = hs_qty_sum.get(hs_code, 1.0) 
        if total_hs_qty == 0: return 0.0
        return total_summary_weight / total_hs_qty

    df["Unit Weight"] = df.apply(get_unit_weight, axis=1)
    df["Total Weight"] = df["Qty"] * df["Unit Weight"]
    
    return df

def merge_similar_items(df):
    """
    Merges rows by "PO No", "ItemCode", "Price".
    """
    if df.empty:
        return df

    group_cols = [
        "PO No", 
        "ItemCode", 
        "Price"
    ]
    
    agg_dict = {
        "Qty": "sum",
        "Amount": "sum",
        "Total Weight": "sum"
    }
    
    for col in df.columns:
        if col not in group_cols and col not in agg_dict:
            agg_dict[col] = "first"

    df_merged = df.groupby(group_cols, as_index=False).agg(agg_dict)
    
    desired_cols = [c for c in df.columns if c in df_merged.columns]
    return df_merged[desired_cols]

def process_pdf(uploaded_file):
    """
    Main logic to parse the VIR Invoice PDF
    """
    # --- CONFIGURATION ---
    supplier_name = "VIR VALVOINDUSTRIA ING. RIZZIO S.P.A."
    invoice_no = ""
    invoice_date = ""
    currency = "UNKNOWN" 
    
    # State variables
    current_oc = ""
    current_po = ""
    current_po_date = ""
    current_item_code = ""
    current_hs_code = "" 
    
    pending_row = None
    extracted_rows = []
    hs_weight_map = {} 

    with pdfplumber.open(uploaded_file) as pdf:
        total_pages = len(pdf.pages)

        # PRE-PASS 1: WEIGHTS (Scan ALL pages)
        for p in pdf.pages:
            text = p.extract_text()
            if not text: continue
            
            # Matches line like: 40169991 8,95 74,40
            weight_matches = re.findall(r'(\d{8})\s+([\d\.]+,\d+)\s+[\d\.]+,\d+', text)
            for code, weight_str in weight_matches:
                hs_weight_map[code] = parse_decimal(weight_str)

        # PRE-PASS 2: CURRENCY (Scan last 2 pages)
        pages_to_scan_currency = [pdf.pages[-1]]
        if total_pages > 1:
            pages_to_scan_currency.insert(0, pdf.pages[-2])
            
        for p in pages_to_scan_currency:
            text = p.extract_text()
            if not text: continue
            
            curr_match = re.search(r'TOTAL AMOUNT\s+([A-Z]{3})', text)
            if curr_match:
                currency = curr_match.group(1)
                break
            elif "USD" in text and "TOTAL AMOUNT" in text:
                currency = "USD"
            elif "EUR" in text and "TOTAL AMOUNT" in text:
                currency = "EUR"

        # MAIN PASS
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text: continue
            
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                
                # Header Extraction
                if i == 0 and "INVOICE N." in line:
                    match = re.search(r'(?:FATTURA|INVOICE)\s+(\d{4}/[A-Z]{2}/\d+).*DATE\s+(\d{2}/\d{2}/\d{2})', line)
                    if match:
                        invoice_no = match.group(1)
                        invoice_date = match.group(2)
                
                # Capture Data
                if line.startswith("OC") and len(line) > 8 and line[2].isdigit():
                    current_oc = line.split()[0]
                
                # Flexible PO Date Extraction
                if "REF PO-" in line:
                    po_match = re.search(r'REF\s+(PO-[\w\d\-]+)(?:.*?(\d{2}/\d{2}/\d{2}))?', line)
                    if po_match:
                        current_po = po_match.group(1)
                        if po_match.group(2):
                            current_po_date = po_match.group(2)

                if "H.S" in line or "HS" in line:
                    hs_match = re.search(r'H\.?S\.?\s*(\d{8})', line)
                    if hs_match:
                        found_hs = hs_match.group(1)
                        current_hs_code = found_hs
                        if pending_row:
                            pending_row["HSCode"] = found_hs

                if " PZ " not in line:
                    words = line.split()
                    if words and is_item_code(words[0]):
                        current_item_code = words[0]

                # Transaction Line
                if " PZ " in line:
                    if pending_row:
                        extracted_rows.append(pending_row)
                        pending_row = None

                    try:
                        parts = line.split(" PZ ")
                        description = parts[0].strip()
                        math_part = parts[1].strip()
                        
                        desc_words = description.split()
                        if desc_words:
                            first_word = desc_words[0]
                            if is_item_code(first_word):
                                current_item_code = first_word
                                description = " ".join(desc_words[1:])
                        
                        math_tokens = math_part.split()
                        amount_str = math_tokens[-1]
                        price_str = math_tokens[-2] if len(math_tokens) >= 2 else "0"
                        qty_str = math_tokens[0] if len(math_tokens) >= 1 else "0"
                        discount_str = "0"
                        
                        pending_row = {
                            "Inv No.": invoice_no,
                            "Inv Date": invoice_date,
                            "Supplier Name": supplier_name,
                            "Order confirmation number": current_oc,
                            "PO No": current_po,
                            "PO Date": current_po_date,
                            "ItemCode": current_item_code,
                            "Item Desc": description,
                            "Currency": currency,
                            "Qty": parse_decimal(qty_str),
                            "Price": parse_decimal(price_str),
                            "Discount": parse_decimal(discount_str),
                            "Amount": parse_decimal(amount_str),
                            "VAT": "", 
                            "HSCode": current_hs_code, 
                            "Unit Weight": 0.0,
                            "Total Weight": 0.0
                        }
                    except Exception:
                        continue

    if pending_row:
        extracted_rows.append(pending_row)

    # Create DataFrame
    df = pd.DataFrame(extracted_rows)
    
    # 1. Calculate Weights
    df = calculate_weights(df, hs_weight_map)
    
    # 2. Enrich with Master Data (CSV)
    df = load_and_enrich_data(df)
    
    cols = [
        "Inv No.", "Inv Date", "Supplier Name", 
        "Order confirmation number", 
        "PO No", "PO Date", 
        "ItemCode", "Item Desc", "Fig No", "Size MM", "Lot No", "Product Category", "Origin",
        "Currency", 
        "Qty", "Price", "Discount", "Amount", "VAT", "HSCode", 
        "Unit Weight", "Total Weight"
    ]
    
    for c in cols:
        if c not in df.columns: df[c] = ""
            
    return df[cols]

def to_excel(df):
    """
    Converts dataframe to Excel bytes with a TOTALS row.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Calculate Totals
        total_qty = df['Qty'].sum()
        total_amount = df['Amount'].sum()
        total_weight = df['Total Weight'].sum()
        
        # Create Totals Row
        total_data = {col: [None] for col in df.columns}
        total_data['Qty'] = [total_qty]
        total_data['Amount'] = [total_amount]
        total_data['Total Weight'] = [total_weight]
        
        df_total = pd.DataFrame(total_data)
        
        # Write Data
        df_total.to_excel(writer, index=False, header=False, startrow=0, sheet_name='InvoiceData')
        df.to_excel(writer, index=False, header=True, startrow=1, sheet_name='InvoiceData')
        
    return output.getvalue()

# --- Streamlit UI ---

st.set_page_config(page_title="LITSOL PDF Converter", layout="wide")

st.title("📄 LITSOL - PDF Invoice to Excel Converter")
st.markdown("Upload the **VIR Invoice PDF**. The app also looks for `Item Master.csv` to populate additional details.")

uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])

if uploaded_file is not None:
    with st.spinner('Parsing PDF and looking up Master Data...'):
        try:
            # 1. Process Raw Data
            df_raw = process_pdf(uploaded_file)
            
            # 2. Create Merged Version
            df_merged = merge_similar_items(df_raw)
            
            st.success(f"Success! Found {len(df_raw)} raw items. Merged into {len(df_merged)} unique items.")
            
            # 3. Display
            st.subheader("Preview (Merged Data)")
            st.dataframe(df_merged, use_container_width=True)
            
            # 4. Downloads
            col1, col2 = st.columns(2)
            
            with col1:
                excel_raw = to_excel(df_raw)
                st.download_button(
                    label="📥 Download Detailed Excel (All Lines)",
                    data=excel_raw,
                    file_name="invoice_detailed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col2:
                # Merged: Exclude OC Column
                df_merged_download = df_merged.drop(columns=["Order confirmation number"], errors='ignore')
                excel_merged = to_excel(df_merged_download)
                
                st.download_button(
                    label="📥 Download Merged Excel (Combined Qty)",
                    data=excel_merged,
                    file_name="invoice_merged.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            
        except Exception as e:
            st.error(f"An error occurred: {e}")