import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

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
    Checks if a word looks like an Item Code (e.g., F0900B032.2683)
    """
    # Regex: Starts with F or K, alphanumeric chars, literal dot, digits
    return re.match(r'^[FK][A-Z0-9]+\.\d+$', word) is not None

def process_pdf(uploaded_file):
    """
    Main logic to parse the VIR Invoice PDF
    """
    # Configuration
    supplier_name = "VIR VALVOINDUSTRIA ING. RIZZIO S.P.A."
    invoice_no = ""
    currency = "UNKNOWN" 
    
    # State variables
    current_oc = ""
    current_po = ""
    current_item_code = ""
    # We keep track of the last seen HS code to backfill if needed
    current_hs_code = "" 
    
    # --- CRITICAL CHANGE: Move pending_row OUTSIDE the page loop ---
    # This ensures that if an item starts on Page 6 and the HS Code 
    # is on Page 7, the row stays "open" across the page break.
    pending_row = None
    
    extracted_rows = []
    hs_weight_map = {} 

    with pdfplumber.open(uploaded_file) as pdf:
        total_pages = len(pdf.pages)

        # --- PRE-PASS 1: WEIGHTS (Scan last 2 pages) ---
        # Scan last 2 pages to be safe
        pages_to_scan = [pdf.pages[-1]]
        if total_pages > 1:
            pages_to_scan.insert(0, pdf.pages[-2])
            
        for p in pages_to_scan:
            text = p.extract_text()
            if not text: continue
            weight_matches = re.findall(r'(\d{8})\s+([\d\.]+,\d+)\s+[\d\.]+,\d+', text)
            for code, weight_str in weight_matches:
                hs_weight_map[code] = parse_decimal(weight_str)

        # --- PRE-PASS 2: CURRENCY ---
        try:
            last_page_text = pdf.pages[-1].extract_text()
            if "USD" in last_page_text: currency = "USD"
            elif "EUR" in last_page_text: currency = "EUR"
        except:
            pass

        # --- MAIN PASS: PROCESS LINE ITEMS ---
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text: continue
            
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                
                # 1. HEADER INFO
                if i == 0 and "INVOICE N." in line:
                    match = re.search(r'FATTURA\s+([\w/]+)', line)
                    if match: invoice_no = match.group(1)
                
                # 2. CAPTURE DATA
                
                # Order Confirmation
                if line.startswith("OC") and len(line) > 8 and line[2].isdigit():
                    current_oc = line.split()[0]
                
                # PO Number
                if "REF PO-" in line:
                    po_match = re.search(r'REF\s+(PO-[\w\d]+)', line)
                    if po_match:
                        current_po = po_match.group(1)

                # HS Code Logic
                # Check for "H.S", "HS", "H.S." followed by 8 digits
                if "H.S" in line or "HS" in line:
                    hs_match = re.search(r'H\.?S\.?\s*(\d{8})', line)
                    if hs_match:
                        found_hs = hs_match.group(1)
                        current_hs_code = found_hs # Update sticky HS Code
                        
                        # IF we have a pending row waiting for an HS code, update it now.
                        # This works even if this line is on the NEXT page from the item description.
                        if pending_row:
                            pending_row["HSCode"] = found_hs
                            pending_row["net weight"] = hs_weight_map.get(found_hs, 0.0)

                # Item Code (Standalone line)
                if (re.match(r'^[FK][A-Z0-9]+\.\d+', line) or re.match(r'^[A-Z]{3}\d+', line)) and " PZ " not in line:
                    current_item_code = line.split()[0]

                # 3. TRANSACTION LINE ("PZ")
                if " PZ " in line:
                    # COMMIT PREVIOUS ROW: 
                    # We only save the previous row when we hit a NEW item line.
                    if pending_row:
                        extracted_rows.append(pending_row)
                        pending_row = None

                    try:
                        # Parsing logic
                        parts = line.split(" PZ ")
                        description = parts[0].strip()
                        math_part = parts[1].strip()
                        
                        # Handle Item Code merged in Description
                        desc_words = description.split()
                        if desc_words and is_item_code(desc_words[0]):
                            current_item_code = desc_words[0]
                            description = " ".join(desc_words[1:])
                        
                        math_tokens = math_part.split()
                        amount_str = math_tokens[-1]
                        price_str = math_tokens[-2] if len(math_tokens) >= 2 else "0"
                        qty_str = math_tokens[0] if len(math_tokens) >= 1 else "0"
                        discount_str = "0"
                        
                        # CREATE NEW PENDING ROW
                        # We initialize HSCode with 'current_hs_code' in case the HS Code 
                        # appeared ABOVE the PZ line (Screenshot 1).
                        # If the HS Code appears BELOW or on NEXT PAGE (Screenshot 2), 
                        # the logic above will overwrite this with the correct value later.
                        
                        # Determine initial weight
                        init_weight = hs_weight_map.get(current_hs_code, 0.0)
                        
                        pending_row = {
                            "Inv No.": invoice_no,
                            "Supplier Name": supplier_name,
                            "Order confirmation number": current_oc,
                            "PO No": current_po,
                            "ItemCode": current_item_code,
                            "Item Desc": description,
                            "Currency": currency,
                            "Qty": parse_decimal(qty_str),
                            "Price": parse_decimal(price_str),
                            "Discount": parse_decimal(discount_str),
                            "Amount": parse_decimal(amount_str),
                            "VAT": "", 
                            "HSCode": current_hs_code, # Use sticky value tentatively
                            "net weight": init_weight
                        }
                        
                    except Exception:
                        continue
            
            # REMOVED: "End of Page Commit". 
            # We do NOT save pending_row here. We let it persist to the next page.

    # FINAL COMMIT
    # Only after processing ALL pages do we save the very last item.
    if pending_row:
        extracted_rows.append(pending_row)

    # Create DataFrame
    df = pd.DataFrame(extracted_rows)
    
    cols = ["Inv No.", "Supplier Name", "Order confirmation number", "PO No", 
            "ItemCode", "Item Desc", "Currency", "Qty", "Price", 
            "Discount", "Amount", "VAT", "HSCode", "net weight"]
    
    for c in cols:
        if c not in df.columns: df[c] = ""
            
    return df[cols]

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='InvoiceData')
    return output.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="PDF to Excel Utility", layout="wide")
st.title("📄 LITSOL - PDF Invoice to Excel Converter")
st.markdown("Upload the **VIR Invoice PDF**. Supports multi-page items.")

uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])

if uploaded_file is not None:
    with st.spinner('Parsing PDF...'):
        try:
            df_result = process_pdf(uploaded_file)
            st.success(f"Success! Extracted {len(df_result)} line items.")
            st.dataframe(df_result, use_container_width=True)
            
            st.download_button(
                label="📥 Download Excel Sheet",
                data=to_excel(df_result),
                file_name="converted_invoice.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")