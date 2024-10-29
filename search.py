
import os
import pandas as pd
import re

# Define the directory where the Excel files are stored
base_dir ="Output"

# List of financial variables to search for
variables = ["Stock Ticker", "Year", "Sales", "Total Assets"]

# Create an empty DataFrame to store the extracted data
compiled_data = pd.DataFrame(columns=variables)

# Function to search for variable within the Excel sheet
def search_for_variable(df, variable_name):
    try:
        # Convert all columns to strings
        df = df.astype(str)
        
        for col in df.columns:
            mask = df[col].str.contains(variable_name, case=False, na=False)
            if mask.any():
                row_index = mask.idxmax()  # Get the first match
                col_index = df.columns.get_loc(col)
                # Start checking from the next column and continue until a suitable value is found
                for offset in range(1, len(df.columns) - col_index):
                    value = df.loc[row_index, df.columns[col_index + offset]]
                    # Check if the value is non-empty and not a bracketed single digit
                    if pd.notna(value) and value.strip() and not re.match(r"^\[\d\]$", value.strip()):
                        return value
    except Exception as e:
        print(f"Error searching for {variable_name}: {e}")
    return None


def search_for_goodwill(df):
    if not isinstance(df, pd.DataFrame):
        print("Error: Input is not a pandas DataFrame")
        return None

    try:
        # Convert all columns to strings
        df = df.astype(str)
        
        # Pattern to match exactly 'Goodwill' but not 'Goodwill Impairment'
        goodwill_pattern = r'\bGoodwill\b'
        
        for col in df.columns:
            # This mask checks for 'Goodwill' exactly and excludes any 'Impairment' mentions
            mask = df[col].str.contains(goodwill_pattern, case=False, na=False) & ~df[col].str.contains("Impairment", case=False, na=False)
            if mask.any():
                row_index = mask.idxmax()  # Get the first match
                # Check the next column for the value if exists
                next_col_index = df.columns.get_loc(col) + 1
                if next_col_index < len(df.columns):
                    value = df.loc[row_index, df.columns[next_col_index]]  # Get the value in the next column
                    # Ensure value is meaningful
                    if pd.notna(value) and not re.match(r"\[\d+\]", str(value).strip()):
                        return value
    except Exception as e:
        print(f"Error searching for Goodwill: {e}")
    return None

# Function to search for multiple words (e.g., "plant" and "net") within the Excel sheet
def search_for_variable_with_multiple_keywords(df, keywords):
    try:
        df = df.astype(str)  # Convert all columns to strings

        for col in df.columns:
            mask = df[col].str.contains(keywords[0], case=False, na=False) & df[col].str.contains(keywords[1], case=False, na=False)
            if mask.any():
                row_index = mask.idxmax()  # Get the first match
                col_index = df.columns.get_loc(col)

                # Start checking from the next column and continue until a non-empty value is found
                for offset in range(1, len(df.columns) - col_index):
                    value = df.loc[row_index, df.columns[col_index + offset]]
                    if pd.notna(value) and not re.match(r"\[\d+\]", str(value).strip()):
                        return value  # Return the first non-empty, non-whitespace value
    except Exception as e:
        print(f"Error searching for {keywords}: {e}")
    return None

def search_for_variable_with_any_keywords(df, keywords):
    try:
        # Convert all columns to strings
        df = df.astype(str)
        
        # Loop through each column to search for the combination of keywords
        for col in df.columns:
            mask = df[col].str.contains(keywords[0], case=False, na=False) | df[col].str.contains(keywords[1], case=False, na=False)
            
            if mask.any():
                row_index = mask.idxmax()  # Get the first match
                col_index = df.columns.get_loc(col)

                # Check subsequent columns for a valid value
                for offset in range(1, len(df.columns) - col_index):
                    next_col_index = col_index + offset
                    value = df.loc[row_index, df.columns[next_col_index]]

                    # Check if value is valid (not empty, not bracketed number)
                    if pd.notna(value) and value.strip() and not re.match(r"^\[\d+\]$", value):
                        return value  # Return the first valid value

    except Exception as e:
        print(f"Error searching for {keywords}: {e}")
    return None




# Function to search for EBIT (Operating Income) within the Excel sheet
def search_for_ebit(df):
    try:
        # Convert all columns to strings
        df = df.astype(str)
        
        # Escape special characters in the strings to avoid regex issues
        ebit_pattern_1 = re.escape("Operating Income")
        ebit_pattern_2 = re.escape("Income (Loss)")
        
        # Loop through each column to search for EBIT keywords
        for col in df.columns:
            # Create a mask for rows that contain EBIT-related terms, escaping any special characters
            mask = df[col].str.contains(ebit_pattern_1, case=False, na=False) | df[col].str.contains(ebit_pattern_2, case=False, na=False)
            
            if mask.any():
                row_index = mask.idxmax()  # Get the first match
                col_index = df.columns.get_loc(col)
                
                # Check if the next column exists
                if col_index + 1 < len(df.columns):
                    value = df.loc[row_index, df.columns[col_index + 1]]  # Get the value in the next column
                    return value
                else:
                    return None  # Skip if there is no next column
    except Exception as e:
        print(f"Error searching for EBIT: {e}")
    return None



# Function to extract the year from a specific cell
def extract_year(df):
    try:
        # Convert all columns to strings
        df = df.astype(str)
        
        # Search for "Document Period End Date" which might indicate the year
        mask = df.apply(lambda x: x.str.contains('Document Period End Date', case=False, na=False)).any(axis=1)
        if mask.any():
            row_index = mask.idxmax()  # Get the first match
            year = df.loc[row_index, df.columns[1]]  # Year should be next to "Document Period End Date"
            return year
    except Exception as e:
        print(f"Error extracting year: {e}")
    return None

def search_for_inventory(df):
    try:

        # Convert all columns to strings
        df = df.astype(str)
        
        # Define the patterns to search for, ordered by priority
        inventory_patterns = [
            r"^\s*Total inventories\s*$",  # Exact match for 'Total inventories'
            r"inventories.*net",           # 'Inventories' with 'net' somewhere in the text
            r"inventories"                 # Any form of 'inventories'
        ]

        for pattern in inventory_patterns:
            for col in df.columns:
                mask = df[col].str.contains(pattern, case=False, na=False, regex=True)
                if mask.any():
                    row_index = mask.idxmax()  # Get the first match
                    next_col_index = df.columns.get_loc(col) + 1
                    if next_col_index < len(df.columns):
                        value = df.loc[row_index, df.columns[next_col_index]]  # Get the value in the next column
                        if pd.notna(value) and value.strip() != "":
                            return value  # Return the first non-empty, non-whitespace value
    except Exception as e:
        print(f"Error searching for inventory: {e}")
    return None


def search_for_stock_based_compensation(df):
    try:
        # Ensure df is a DataFrame
        if not isinstance(df, pd.DataFrame):
            print("Error: Input is not a pandas DataFrame")
            return None

        # Convert all columns to strings
        df = df.astype(str)
        
        # Define regex pattern for stock-based compensation, focusing on 'stock' and 'compensation' keywords
        pattern = r"stock[-].*compensation.|compensation[-].*stock.|share[-].*compensation."

        for col in df.columns:
            mask = df[col].str.contains(pattern, case=False, na=False, regex=True)
            if mask.any():
                row_index = mask.idxmax()  # Get the first match
                col_index = df.columns.get_loc(col)

                # Check subsequent columns for a valid value
                for offset in range(1, len(df.columns) - col_index):
                    next_col_index = col_index + offset
                    value = df.loc[row_index, df.columns[next_col_index]]

                    # Check if value is valid (not empty, not bracketed number)
                    if pd.notna(value) and value.strip() and not re.match(r"^\[\d+\]$", value):
                        return value  # Return the first valid value

    except Exception as e:
        print(f"Error searching for stock-based compensation: {e}")
    return None

# Iterate over each company's folder
for company_folder in os.listdir(base_dir):
    company_path = os.path.join(base_dir, company_folder)
    
    if os.path.isdir(company_path):
        for file in os.listdir(company_path):
            # Skip temporary Excel files starting with ~$
            if file.startswith("~$"):
                print(f"Skipping temporary file: {file}")
                continue
            
            if file.endswith(".xlsx"):
                file_path = os.path.join(company_path, file)
                
                # Read the Excel file using openpyxl engine
                try:
                    # Load the Excel file with all sheets
                    xls = pd.ExcelFile(file_path, engine="openpyxl")
                    
                    # Initialize variables for the current file
                    sales = None
                    total_assets = None
                    year = None
                    goodwill=None
                    provision=None
                    rd=None
                    provision=None
                    plant=None
                    EBIT=None
                    intangible_asset=None
                    inventory=None
                    compensation=None
                    subsidiaries=None
                    auditor_fee=None
                    
                    stock_based_compensation = None
                    EBIT=None
                    
                    # Iterate over each sheet in the Excel file
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        
                        # Extract the year from the specific sheet
                        if year is None:
                            year = extract_year(df)
                        
                        # Search for the required financial variables
                        if sales is None:
                            sales = search_for_variable_with_any_keywords(df, ["Net sales","sales"])
                        if total_assets is None:
                            total_assets = search_for_variable(df, "Total assets")
                        if goodwill is None:
                            goodwill=search_for_goodwill(df)
                        if provision is None:
                            provision=search_for_variable_with_any_keywords(df,["provision","provisions","provisional"])
                        if rd is None:
                            rd=search_for_variable(df,"Research and development expense")
                         # Search for stock-based compensation
                        if stock_based_compensation is None:
                            stock_based_compensation = search_for_stock_based_compensation(df)
                       
                        # Iterate over each company's folder (rest of the code remains the same)

                        if plant is None:
                            plant = search_for_variable_with_multiple_keywords(df, ["plant", "property"])

                        if intangible_asset is None:
                            intangible_asset=search_for_variable_with_any_keywords(df,["Intangible assets","intangibles","intangible"])
                        if inventory is None:
                            inventory=search_for_inventory(df)
                        if EBIT is None:
                            EBIT = search_for_ebit(df)
                        if auditor_fee is None:
                            auditor_fee=search_for_variable(df,"Selling, general and administrative expenses")
                       
                        
                        
                     
                        
                        

                        
                        # Stop searching once we have all required data
                        if sales is not None and total_assets is not None and year is not None :
                            break
                    
                    # Use folder name as stock ticker
                    stock_ticker = company_folder
                    if sales is None:
                        sales = "Nan"
                    if total_assets is None:
                        total_assets = "Nan"
                    if year is None:
                        year = "Unknown"
                    if goodwill is None:
                        goodwill = "Nan"
                    if provision is None:
                        provision = "Nan"
                    if plant is None:
                        plant = "Nan"
                    if intangible_asset is None:
                        intangible_asset = "Nan"
                    if inventory is None:
                        inventory = "Nan"
                    if stock_based_compensation is None:
                        stock_based_compensation = "Nan"
                    if EBIT is None:
                        EBIT = "Nan"
                    if auditor_fee is None:
                        auditor_fee = "Nan"
                    
                    
                    # Create a temporary DataFrame for the current file's data
                    temp_data = pd.DataFrame([{
                        "Stock Ticker": stock_ticker,
                        "Year": year,
                        "Sales": sales,
                        "Total Assets": total_assets,
                        "Goodwill": goodwill,
                        "Corporate Tax(Provision)":provision,
                        "Plant,Property and equipment":plant,
                        "Intangible Assets": intangible_asset,
                        "Goodwill": goodwill,
                        "Inventories": inventory,
                        "Executive compensation" :stock_based_compensation,
                        "EBIT": EBIT,
                        "Auditor fee": auditor_fee,
                        
                    }])

                    # Concatenate the temp_data with the compiled_data DataFrame
                    compiled_data = pd.concat([compiled_data, temp_data], ignore_index=True)

                except Exception as e:
                    print(f"Error processing {file_path}: {e}")

# Save the compiled data to a CSV file
compiled_data.to_csv("plantfinal.csv", index=False)

print("Data extraction complete. Compiled data saved to 'f3.csv'.")
