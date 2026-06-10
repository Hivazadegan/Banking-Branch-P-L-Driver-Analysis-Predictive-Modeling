import os
import re
import pandas as pd
import logging

# --- Configuration ---
INPUT_FILE = "14030231.xlsx"
OUTPUT_DIR = "output"
OUTPUT_CLEANED = os.path.join(OUTPUT_DIR, "cleaned_banking_pnl_advanced.csv")
OUTPUT_SUMMARY = os.path.join(OUTPUT_DIR, "banking_pnl_summary_advanced.csv")
LOG_FILE = os.path.join(OUTPUT_DIR, "processing.log")

# --- Setup Logging ---
os.makedirs(OUTPUT_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)


# --- Helper Functions ---
def normalize_text(text):
    """Normalize Persian/Arabic text for cleaner matching."""
    if pd.isna(text):
        return ""
    text = str(text)
    # Replace common variations of Persian/Arabic characters
    text = re.sub(r'[ي]', 'ی', text)
    text = re.sub(r'[ك]', 'ک', text)
    # Remove zero-width non-joiner and other control characters if any
    text = re.sub(r'[\u200c\u200d]', '', text)
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)  # Consolidate whitespace
    return text


def clean_column_names(df):
    """Normalize column names for consistency."""
    logging.info("Cleaning column names...")
    original_columns = df.columns.tolist()
    normalized_columns = [normalize_text(col) for col in original_columns]
    df.columns = normalized_columns
    logging.info(f"Columns normalized: {', '.join(normalized_columns[:5])}...")  # Log first few
    return df


def to_numeric_safe(series, column_name=""):
    """Convert numeric-like values to numbers safely, handling Persian numerals and separators."""
    if series.dtype == "O":  # Object type, likely strings
        s = series.astype(str)
        # Remove common separators and invalid characters
        s = s.str.replace(",", "", regex=False)
        s = s.str.replace("،", "", regex=False)  # Persian comma
        s = s.str.replace("ـ", "", regex=False)  # Arabic hyphen
        s = s.str.replace(" ", "", regex=False)
        s = s.str.replace("(", "", regex=False)
        s = s.str.replace(")", "", regex=False)

        # Convert Persian digits to English digits
        persian_digits = "۰۱۲۳۴۵۶۷۸۹"
        english_digits = "0123456789"
        translation_table = str.maketrans(persian_digits, english_digits)
        s = s.apply(lambda x: x.translate(translation_table) if isinstance(x, str) else x)

        # Attempt conversion to numeric, coercing errors to NaN
        numeric_series = pd.to_numeric(s, errors="coerce")

        # Log if significant portion became NaN during conversion
        if numeric_series.isna().sum() > 0 and numeric_series.isna().sum() / len(series) > 0.1:
            logging.warning(
                f"Column '{column_name}': Coerced {numeric_series.isna().sum()} values to NaN during numeric conversion.")

        return numeric_series

    # If already numeric, ensure it's float for potential NaNs
    return pd.to_numeric(series, errors="coerce")


def find_columns(df, keywords):
    """Find columns containing any of the given keywords, case-insensitive and with normalization."""
    keywords_norm = [normalize_text(k) for k in keywords]
    matches = []
    for col in df.columns:
        col_norm = normalize_text(col)
        if any(k in col_norm for k in keywords_norm):
            matches.append(col)
    return matches


def load_data(file_path):
    """Load the Excel file with error handling."""
    logging.info(f"Loading data from {file_path}...")
    if not os.path.exists(file_path):
        logging.error(f"Input file not found: {file_path}")
        raise FileNotFoundError(f"Input file not found: {file_path}")
    try:
        # Read only the first sheet by default if not specified
        df = pd.read_excel(file_path, engine='openpyxl')
        logging.info(f"Successfully loaded {len(df)} rows and {len(df.columns)} columns.")
        return df
    except Exception as e:
        logging.error(f"Error loading Excel file {file_path}: {e}")
        raise


def clean_data(df):
    """Perform comprehensive data cleaning and normalization."""
    logging.info("Starting data cleaning process...")

    # Drop rows where all values are NaN
    initial_rows = len(df)
    df.dropna(how='all', inplace=True)
    rows_dropped = initial_rows - len(df)
    if rows_dropped > 0:
        logging.info(f"Dropped {rows_dropped} rows with all NaN values.")

    df = clean_column_names(df)

    # Define identifier columns and potential financial/descriptive columns
    # Adjust these based on typical banking P&L structures
    id_cols_keywords = {
        "BranchCode": ["کد شعبه", "branch code"],
        "BranchName": ["نام شعبه", "branch name"],
        "ProvinceCode": ["کد استان", "province code"],
        "ProvinceName": ["نام استان", "province name"],
        "Date": ["تاریخ", "date", "ماه", "سال"],  # Will need more sophisticated date parsing
        "AccountNumber": ["شماره حساب", "account number"]
    }

    financial_keywords = [
        "درآمد", "هزینه", "سود", "زیان", "کارمزد", "تسهیلات", "سپرده",
        "مطالبات", "تامین مالی", "بهره", "سود خالص", "سود عملیاتی",
        "هزینه کارکنان", "هزینه اداری", "استهلاک", "ارزی", "تراکنش",
        "Balance", "Revenue", "Expense", "Profit", "Loss", "Interest", "Fee"
    ]

    # Apply cleaning and conversion
    for col in df.columns:
        normalized_col = normalize_text(col)

        # Try to identify and convert identifier columns
        is_id_col = False
        for target_name, kws in id_cols_keywords.items():
            if any(kw in normalized_col for kw in kws):
                logging.debug(f"Attempting to clean ID column: '{col}'")
                # Convert to string and normalize for cleaner IDs
                df[col] = df[col].astype(str).apply(normalize_text)
                # Attempt to convert to Int64 if it looks like a number
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
                except ValueError:
                    pass  # Keep as string if not convertible
                is_id_col = True
                break

        if is_id_col: continue  # Skip further processing for known ID cols

        # Check if column is likely financial/numeric based on keywords
        is_financial = any(kw in normalized_col for kw in financial_keywords)

        # If it contains financial keywords OR if it's not purely textual (after normalization)
        if is_financial or df[col].dtype == 'object' and df[col].apply(
            lambda x: isinstance(x, (int, float)) or re.match(r'^[\d\.,-]+$', normalize_text(x)) or re.match(
                r'^[\d\.,-]+$', x) if pd.notna(x) else False