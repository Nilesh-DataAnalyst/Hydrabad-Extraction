import pandas as pd
import os
import re

# ===== FILE PATHS =====
INPUT_PATH = r"D:\Nilesh\Auto_Script\Hydrabad_Final_Merged_File.xlsx"
OUTPUT_PATH = r"D:\Nilesh\Auto_Script\Hydrabad_Final_Merged_Extracted.xlsx"

# Party type mappings (Seller types, Buyer types)
SELLER_TYPES = {"EX", "MR", "DR", "RR", "PL", "LR","FP"}  # Added PL and LR as Sellers
BUYER_TYPES = {"CL", "ME", "DE", "RE", "AY", "LE","SP"}   # Added AY and LE as Buyers
ALL_PARTY_TYPES = SELLER_TYPES | BUYER_TYPES  # Union of all types

BOUND_START_RE = re.compile(r"(?i)\bbound\w*\s*:\s*")

def clean_spaces(s: str) -> str:
    """Clean extra spaces but preserve intentional newlines"""
    if not isinstance(s, str):
        return s
    # Split into lines, clean each line, then rejoin
    lines = s.split('\n')
    cleaned_lines = [re.sub(r"\s+", " ", line).strip() for line in lines]
    return "\n".join(cleaned_lines)

def normalize_basic(s: str) -> str:
    s = str(s)
    s = re.sub(r"(?i)VILL\s*/\s*COL", "VILL/COL", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def pick_unit_token(v: str, field_name: str = "") -> str:
    if not v:
        return ""
    
    # For BUILT field, return the full value as is (including unit)
    if field_name == "BUILT":
        return v.strip()
    
    # For EXTENT field, keep the original logic
    for p in v.split()[:12]:
        if "SQ" in p.upper():
            return p
    return v

def convert_extent_to_sq_ft(extent_value: str) -> str:
    """Convert extent value from SQ.Yd to SQ.Ft (direct conversion)"""
    if pd.isna(extent_value) or not isinstance(extent_value, str) or not extent_value:
        return ""
    
    extent_value = extent_value.strip()
    
    # Check if it's in SQ.Yd format (e.g., "190SQ.Yd" or "190 SQ.Yd" or "190.5SQ.Yd")
    sq_yd_match = re.search(r"([\d.]+)\s*SQ\.?Yd\.?", extent_value, re.IGNORECASE)
    
    if sq_yd_match:
        try:
            sq_yd = float(sq_yd_match.group(1))
            sq_ft = sq_yd * 9  # Direct conversion: 1 Square Yard = 9 Square Feet
            # Format to 2 decimal places - REMOVED UNIT
            return f"{sq_ft:.2f}"
        except ValueError:
            return ""
    
    return ""

def convert_built_to_sq_ft(built_value: str) -> str:
    """Convert built-up area value from SQ.Ft to SQ.Ft (no conversion needed, just formatting)"""
    if pd.isna(built_value) or not isinstance(built_value, str) or not built_value:
        return ""
    
    built_value = built_value.strip()
    
    # Check if it's in SQ.Ft format - handles "50SQ. FT", "50 SQ.FT", "50SQ.Ft", "50 SQ. FT", etc.
    sq_ft_match = re.search(r"([\d.]+)\s*SQ\.?\s*FT\.?", built_value, re.IGNORECASE)
    
    if sq_ft_match:
        try:
            sq_ft = float(sq_ft_match.group(1))
            # Format to 2 decimal places - REMOVED UNIT
            return f"{sq_ft:.2f}"
        except ValueError:
            return ""
    
    return ""

def extract_dates(date_text: str) -> dict:
    """Extract Registration, Execution, and Presentation dates"""
    out = {"Registration Date": "", "Execution Date": "", "Presentation Date": ""}
    if pd.isna(date_text) or not isinstance(date_text, str):
        return out
    
    # Pattern to match (R) date, (E) date, (P) date
    r_match = re.search(r"\(R\)\s*(\d{1,2}-\d{1,2}-\d{4})", date_text, re.IGNORECASE)
    e_match = re.search(r"\(E\)\s*(\d{1,2}-\d{1,2}-\d{4})", date_text, re.IGNORECASE)
    p_match = re.search(r"\(P\)\s*(\d{1,2}-\d{1,2}-\d{4})", date_text, re.IGNORECASE)
    
    if r_match:
        out["Registration Date"] = r_match.group(1)
    if e_match:
        out["Execution Date"] = e_match.group(1)
    if p_match:
        out["Presentation Date"] = p_match.group(1)
    
    return out

def extract_document_info(doc_text: str) -> dict:
    """Extract Document type code, Document Type, Market Value, Consideration Value"""
    out = {
        "Document type code": "", 
        "Document Type": "", 
        "Market Value": "", 
        "Consideration Value": ""
    }
    if pd.isna(doc_text) or not isinstance(doc_text, str):
        return out
    
    # Extract document type code (first 4 digits)
    code_match = re.search(r"^(\d{4})", doc_text.strip())
    if code_match:
        out["Document type code"] = code_match.group(1)
    
    # Extract document type (between code and Mkt.Value)
    doc_type_match = re.search(r"^\d{4}\s+(.+?)(?:\s+Mkt\.Value:|$)", doc_text, re.IGNORECASE)
    if doc_type_match:
        out["Document Type"] = clean_spaces(doc_type_match.group(1))
    
    # Extract Market Value
    mkt_match = re.search(r"Mkt\.Value:\s*(?:Rs\.?)?\s*([0-9,]+)", doc_text, re.IGNORECASE)
    if mkt_match:
        out["Market Value"] = mkt_match.group(1).replace(",", "")
    
    # Extract Consideration Value
    cons_match = re.search(r"Cons\.Value:\s*(?:Rs\.?)?\s*([0-9,]+)", doc_text, re.IGNORECASE)
    if cons_match:
        out["Consideration Value"] = cons_match.group(1).replace(",", "")
    
    return out

def extract_parties(parties_text: str) -> dict:
    """Extract Seller and Buyer from party information - supports multiple sellers and buyers"""
    out = {"Seller": "", "Buyer": ""}
    if pd.isna(parties_text) or not isinstance(parties_text, str):
        return out
    
    # Clean the text first
    parties_text = clean_spaces(parties_text)
    
    # Remove duplicate content (if the text is repeated)
    text_length = len(parties_text)
    half_length = text_length // 2
    
    if text_length > 20 and parties_text[:half_length] == parties_text[half_length:]:
        parties_text = parties_text[:half_length]
    
    sellers = []
    buyers = []
    seen_sellers = set()
    seen_buyers = set()
    
    # FIRST: Split by numbered entries (1., 2., 3., etc.) - THIS IS CRITICAL
    # This regex splits on spaces followed by a number and dot
    entries = re.split(r'\s+(?=\d+\.)', parties_text)
    
    # If splitting didn't work well, try alternative split
    if len(entries) <= 1:
        # Find all numbered entries
        entries = re.findall(r'\d+\.[^.]*(?:\([^)]+\)[^.]*)*', parties_text)
    
    # Process each numbered entry separately
    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue
        
        # Extract the number and the rest
        number_match = re.match(r'(\d+)\.\s*(.*)', entry)
        if number_match:
            number, content = number_match.groups()
        else:
            content = entry
        
        # Check for party type in this entry
        found_type = None
        for party_type in ALL_PARTY_TYPES:
            type_pattern = rf'\(({party_type})\)'
            type_match = re.search(type_pattern, content, re.IGNORECASE)
            if type_match:
                found_type = party_type.upper()
                # Remove the party type tag from content
                content = re.sub(type_pattern, '', content, flags=re.IGNORECASE).strip()
                break
        
        if not found_type:
            continue
        
        # Clean up the name
        name = content.strip()
        
        # Remove any trailing number patterns
        name = re.sub(r'\s+\d+\.\s*$', '', name)
        name = re.sub(r'^\s+|\s+$', '', name)
        
        # Skip if name is empty
        if not name:
            continue
        
        # Create a normalized version for duplicate checking
        normalized_name = re.sub(r'\s+', '', name.upper())
        normalized_name = re.sub(r'[^\w\s]', '', normalized_name)
        
        # Add to appropriate list based on party type
        if found_type in SELLER_TYPES:
            if normalized_name and normalized_name not in seen_sellers:
                sellers.append(name)
                seen_sellers.add(normalized_name)
        elif found_type in BUYER_TYPES:
            if normalized_name and normalized_name not in seen_buyers:
                buyers.append(name)
                seen_buyers.add(normalized_name)
    
    # Join with newlines
    out["Seller"] = "\n".join(sellers) if sellers else ""
    out["Buyer"] = "\n".join(buyers) if buyers else ""
    
    return out

def segment_fields(text: str) -> dict:
    """Extract property description fields from text"""
    out = {k: "" for k in ["VILL/COL", "W-B", "SURVEY", "PLOT", "HOUSE", "APARTMENT", "BLOCK", "FLAT", "EXTENT", "BUILT", "Boundires"]}
    if pd.isna(text):
        return out

    t = normalize_basic(text)
    
    # ---- Fix 1: Ensure proper spacing before "Boundires:" ----
    # Add space before "Boundires:" if it's attached to other text
    t = re.sub(r"([^ ])Boundires:", r"\1 Boundires:", t, flags=re.IGNORECASE)
    t = re.sub(r"([^ ])bound\w*:", r"\1 bound:", t, flags=re.IGNORECASE)
    
    # ---- 1) Boundaries: regex slice (most reliable) ----
    bb = BOUND_START_RE.search(t)
    if bb:
        # Get everything after the boundary marker
        remaining_text = t[bb.end():]
        # Find where the next label might start (to capture complete boundaries)
        next_label_pos = len(remaining_text)
        
        # Look for any of the other labels that might come after boundaries
        for label in ["VILL/COL:", "W-B:", "SURVEY:", "PLOT:", "HOUSE:", "APARTMENT:", "BLOCK:", "FLAT:", "EXTENT:", "BUILT:"]:
            pos = remaining_text.upper().find(label.upper())
            if 0 < pos < next_label_pos:
                next_label_pos = pos
        
        out["Boundires"] = clean_spaces(remaining_text[:next_label_pos])
        # Remove boundaries part from text so it doesn't interfere with other parsing
        t_main = t[:bb.start()].strip()
    else:
        # ---- Alternative: Try to find boundaries at the end of string ----
        # Look for boundary pattern at the end (common pattern with [N], [S], etc.)
        bound_pattern_at_end = re.search(r"(?i)(bound\w*\s*:.*?)(?:\[[NSEW]\].*?)+$", t)
        if bound_pattern_at_end:
            out["Boundires"] = clean_spaces(bound_pattern_at_end.group(1).split(":", 1)[1])
            t_main = t[:bound_pattern_at_end.start()].strip()
        else:
            t_main = t

    # ---- NEW: Handle VILL/COL when it's at the beginning without label ----
    # Check if text starts with something that's not a known label (like "VENKATGIRI-1")
    if t_main and not t_main.upper().startswith(("VILL/COL:", "W-B:", "SURVEY:", "PLOT:", "HOUSE:", "APARTMENT:", "BLOCK:", "FLAT:", "EXTENT:", "BUILT:")):
        # Extract the first part until we hit a known label
        first_part_match = re.match(r"^([^:]+?)\s+(?=W-B:|SURVEY:|PLOT:|HOUSE:|APARTMENT:|BLOCK:|FLAT:|EXTENT:|BUILT:|Boundires:)", t_main, re.IGNORECASE)
        if first_part_match:
            out["VILL/COL"] = clean_spaces(first_part_match.group(1))
            # Remove the extracted part from t_main for further processing
            t_main = t_main[len(first_part_match.group(1)):].strip()

    # ---- 2) Parse other fields by simple label splits ----
    # Ensure labels have ":" (only for known labels)
    # Be careful with HOUSE to not match "/HOUSE SITE" in VILL/COL
    for k in ["VILL/COL", "W-B", "SURVEY", "PLOT", "APARTMENT", "BLOCK", "FLAT", "EXTENT", "BUILT"]:
        t_main = re.sub(rf"(?i)\b{k}\b\s*(?!:)", f"{k}:", t_main)
    
    # Handle HOUSE separately with a more precise pattern
    # Only add colon if HOUSE is at word boundary and not part of VILL/COL
    t_main = re.sub(rf"(?i)(?<!/)\bHOUSE\b\s*(?!:)", "HOUSE:", t_main)

    # Clean up any double colons
    t_main = re.sub(r"::+", ":", t_main)

    def grab(label, s):
        # Improved to better detect field boundaries
        # Special handling for HOUSE to avoid matching in VILL/COL
        if label == "HOUSE":
            # More precise pattern for HOUSE
            m = re.search(rf"(?i)(?<!/)\b{re.escape(label)}\s*:\s*(.*?)(?=\s+(?:VILL/COL|W-B|SURVEY|PLOT|APARTMENT|BLOCK|FLAT|EXTENT|BUILT|Boundires)\s*:|$)", s)
        else:
            m = re.search(rf"(?i)\b{re.escape(label)}\s*:\s*(.*?)(?=\s+(?:VILL/COL|W-B|SURVEY|PLOT|HOUSE|APARTMENT|BLOCK|FLAT|EXTENT|BUILT|Boundires)\s*:|$)", s)
        
        if m:
            value = m.group(1)
            # Additional cleanup: remove any trailing text that might contain next field's label
            value = re.sub(r'\s+(?:VILL/COL|W-B|SURVEY|PLOT|HOUSE|APARTMENT|BLOCK|FLAT|EXTENT|BUILT|Boundires)\s*:.*$', '', value, flags=re.IGNORECASE)
            return clean_spaces(value)
        return ""

    # Only grab VILL/COL if we didn't already extract it from the beginning
    if not out["VILL/COL"]:
        out["VILL/COL"] = grab("VILL/COL", t_main)
    
    out["W-B"]      = grab("W-B", t_main)
    out["SURVEY"]   = grab("SURVEY", t_main)
    out["PLOT"]     = grab("PLOT", t_main)
    out["HOUSE"]    = grab("HOUSE", t_main)
    out["APARTMENT"] = grab("APARTMENT", t_main)
    out["BLOCK"]    = grab("BLOCK", t_main)
    out["FLAT"]     = grab("FLAT", t_main)
    out["EXTENT"]   = grab("EXTENT", t_main)
    out["BUILT"]    = grab("BUILT", t_main)

    # ---- 3) Clean EXTENT/BUILT ----
    # Handle the case where boundaries text might have leaked into EXTENT
    if out["EXTENT"] and any(bound_word in out["EXTENT"].upper() for bound_word in ["BOUND", "[N]", "[S]", "[E]", "[W]"]):
        # Split on boundary markers
        for bound_marker in [" bound", " Bound", " BOUND", "[N]", "[S]", "[E]", "[W]"]:
            if bound_marker in out["EXTENT"]:
                out["EXTENT"] = out["EXTENT"].split(bound_marker)[0].strip()
                break
    
    out["EXTENT"] = pick_unit_token(out["EXTENT"], "EXTENT")
    out["BUILT"]  = pick_unit_token(out["BUILT"], "BUILT")

    # ---- Final check: If boundaries still empty but we see boundary patterns in main text ----
    if not out["Boundires"]:
        # Look for boundary pattern anywhere in original text
        bound_match = re.search(r"(?i)(?:bound\w*\s*:|\b(?:north|south|east|west|n|s|e|w)[\s:]*).*?(?:\[[NSEW]\].*?)+", t)
        if bound_match:
            # Extract just the boundary description part
            bound_text = bound_match.group(0)
            if ":" in bound_text:
                out["Boundires"] = clean_spaces(bound_text.split(":", 1)[1])
            else:
                out["Boundires"] = clean_spaces(bound_text)

    return out

# ===== CLASSIFICATION FUNCTION FOR TRANSACTION TYPE =====
def classify_transaction(doc_type):
    sales_types = [
        "Sale Deed",
        "AGREEMENT OF SALE CUM GPA",
        "Sale Agreement Without Possess",
        "Sale Agreement With Possession",
        "CONVEYANCE FOR CONSIDERATION",
        "CONVEYANCE DEED(WITHOUT CONSID",
        "Sale deed executed by A.P.Hous",
        "Sale Deeds executed by Courts",
        "Sale Certificate",
        "Exchange",
        "Assignment deed",
        "RECONVEYANCE DEED EXECUTED BY",
        "Sale of life interest",
        "Development Agreement Cum GPA",
        "DEVELOPMENT AGREEMENT OR CONST",
        "Sale deed executed by or infav",
        "Sale deed executed by Society",
        "Sale deed in favour of State o"
    ]

    lease_types = [
        "Lease Deed",
        "Lease in favour of State/Centr",
        "Lease(others)",
        "Surrender of Lease",
        "Transfer of Lease"
    ]
    
    if pd.isna(doc_type):
        return "Others"

    doc_type = str(doc_type).strip()

    if doc_type in sales_types:
        return "Sales"
    elif doc_type in lease_types:
        return "Lease"
    else:
        return "Others"

# ---- Apply ----
# Create output directory if it doesn't exist
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

print(f"Reading input file: {INPUT_PATH}")
df = pd.read_excel(INPUT_PATH)
src_col = "Description of property"

# Extract property description fields
print("Extracting property description fields...")
parsed = df[src_col].apply(segment_fields).apply(pd.Series)
df_out = pd.concat([df, parsed], axis=1)

# Add EXTENT in SqFt column (direct conversion from SQ.Yd to SQ.Ft)
print("Converting EXTENT to SqFt...")
df_out["EXTENT in SqFt"] = df_out["EXTENT"].apply(convert_extent_to_sq_ft)

# Add BUILT in SqFt column (no conversion, just formatting)
print("Converting BUILT to SqFt...")
df_out["BUILT in SqFt"] = df_out["BUILT"].apply(convert_built_to_sq_ft)

# Extract dates from Reg.Date Exe.Date Pres.Date column
if "Reg.Date Exe.Date Pres.Date" in df_out.columns:
    print("Extracting dates...")
    dates_parsed = df_out["Reg.Date Exe.Date Pres.Date"].apply(extract_dates).apply(pd.Series)
    # Insert date columns after the original date column
    date_col_idx = df_out.columns.get_loc("Reg.Date Exe.Date Pres.Date") + 1
    for i, col in enumerate(["Registration Date", "Execution Date", "Presentation Date"]):
        df_out.insert(date_col_idx + i, col, dates_parsed[col])

# Extract document info from Nature & Mkt.Value Con. Value column
if "Nature & Mkt.Value Con. Value" in df_out.columns:
    print("Extracting document information...")
    doc_parsed = df_out["Nature & Mkt.Value Con. Value"].apply(extract_document_info).apply(pd.Series)
    # Insert document columns after the original document column
    doc_col_idx = df_out.columns.get_loc("Nature & Mkt.Value Con. Value") + 1
    for i, col in enumerate(["Document type code", "Document Type", "Market Value", "Consideration Value"]):
        df_out.insert(doc_col_idx + i, col, doc_parsed[col])

# Extract seller and buyer from Name of Parties Executant(EX) & Claimants(CL) column
if "Name of Parties Executant(EX) & Claimants(CL)" in df_out.columns:
    print("Extracting seller and buyer information...")
    parties_parsed = df_out["Name of Parties Executant(EX) & Claimants(CL)"].apply(extract_parties).apply(pd.Series)
    # Insert party columns after the original party column
    party_col_idx = df_out.columns.get_loc("Name of Parties Executant(EX) & Claimants(CL)") + 1
    for i, col in enumerate(["Seller", "Buyer"]):
        df_out.insert(party_col_idx + i, col, parties_parsed[col])

# ===== ADD TRANSACTION TYPE COLUMN AFTER DOCUMENT TYPE =====
# Check if Document Type column exists (it should from the extraction above)
if "Document Type" in df_out.columns:
    print("Adding Transaction Type column...")
    # Create transaction type values
    transaction_values = df_out["Document Type"].apply(classify_transaction)
    
    # Insert Transaction Type column after Document Type
    doc_type_idx = df_out.columns.get_loc("Document Type")
    df_out.insert(doc_type_idx + 1, "Transaction Type", transaction_values)
    
    # ===== PRINT TRANSACTION TYPE SUMMARY =====
    print("\n" + "="*60)
    print("TRANSACTION TYPE SUMMARY")
    print("="*60)
    transaction_summary = df_out["Transaction Type"].value_counts()
    sales_count = transaction_summary.get("Sales", 0)
    lease_count = transaction_summary.get("Lease", 0)
    others_count = transaction_summary.get("Others", 0)
    total_count = len(df_out)
    
    # Print in the requested format
    print(f"Sales\tLease\tOthers\tTotal")
    print(f"{sales_count}\t{lease_count}\t{others_count}\t{total_count}")
    print("="*60)
else:
    print("Warning: Document Type column not found. Cannot add Transaction Type.")

# ---- Debug: Show rows where Boundaries is still empty ----
empty_boundaries = df_out[df_out["Boundires"] == ""]
if not empty_boundaries.empty:
    print(f"Found {len(empty_boundaries)} rows with empty boundaries:")
    for idx, row in empty_boundaries.head(10).iterrows():
        print(f"\nRow {idx}:")
        print(f"Original text: {row[src_col]}")
        print("-" * 50)

# Debug: Show sample of BUILT and BUILT in SqFt to verify conversion
print("\n--- BUILT Conversion Sample (first 10 rows) ---")
sample_rows = df_out[["BUILT", "BUILT in SqFt"]].head(10)
print(sample_rows.to_string())

# Debug: Show sample of EXTENT and EXTENT in SqFt to verify conversion
print("\n--- EXTENT Conversion Sample (first 10 rows) ---")
sample_rows = df_out[["EXTENT", "EXTENT in SqFt"]].head(10)
print(sample_rows.to_string())

# Debug: Show sample of parties extraction with new mappings
print("\n--- Parties Extraction Sample with MR/ME, DR/DE, RR/RE Mappings (first 20 rows) ---")
if "Name of Parties Executant(EX) & Claimants(CL)" in df_out.columns:
    # Get rows that might contain the new party types
    sample_df = df_out[["Name of Parties Executant(EX) & Claimants(CL)", "Seller", "Buyer"]].head(20)
    
    # Also check specifically for rows with MR, ME, DR, DE, RR, RE
    print("\nRows with new party types (MR/ME, DR/DE, RR/RE, PL/AY, LR/LE):")
    for idx, row in df_out.head(50).iterrows():
        text = str(row["Name of Parties Executant(EX) & Claimants(CL)"])
        if any(x in text for x in ['(MR)', '(ME)', '(DR)', '(DE)', '(RR)', '(RE)', '(PL)', '(AY)', '(LR)', '(LE)', '(FP)', '(SP)']):
            print(f"\nRow {idx}:")
            print(f"Original: {text[:100]}..." if len(text) > 100 else f"Original: {text}")
            print(f"Seller:\n{row['Seller']}")
            print(f"Buyer:\n{row['Buyer']}")
            print("-" * 40)
    
    print("\nFirst 20 rows sample:")
    print(sample_df.to_string())

# Debug: Show sample of Transaction Type
print("\n--- Transaction Type Sample (first 20 rows) ---")
if "Transaction Type" in df_out.columns:
    trans_sample = df_out[["Document Type", "Transaction Type"]].head(20)
    print(trans_sample.to_string())

# Save the output file
df_out.to_excel(OUTPUT_PATH, index=False)
print(f"\nSaved: {OUTPUT_PATH}")