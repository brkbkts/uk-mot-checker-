import requests
import json
import pandas as pd
import time
from datetime import datetime, timedelta
import logging
import sys
import os
import openpyxl

# Global token variables
access_token = None
token_expiry = None


# API Configuration
# Note: These values need to be replaced with your actual credentials from DVSA
CLIENT_ID = ""
CLIENT_SECRET = ""  
API_KEY = ""
SCOPE_URL = "https://tapi.dvsa.gov.uk/.default"
TOKEN_URL = ""
API_BASE_URL = "https://history.mot.api.gov.uk/v1/trade/vehicles/registration/"

# Global variable to store token and expiry
access_token = None
token_expiry = None

def get_access_token():
    """Get an OAuth access token, reusing existing token if still valid"""
    global access_token, token_expiry
    
    # If we have a valid token, return it
    if access_token and token_expiry and datetime.now() < token_expiry:
        logging.debug("Using cached token")
        return access_token
    
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": SCOPE_URL
    }
    
    try:
        resp = requests.post(TOKEN_URL, headers=headers, data=data)
        resp.raise_for_status()
        token_data = resp.json()
        access_token = token_data["access_token"]
        token_expiry = datetime.now() + timedelta(seconds=int(token_data["expires_in"]) - 60)
        return access_token
    except Exception as e:
        logging.error(f"Access token error: {e}")
        return None

def get_mot_information(registration_number):
    """Get MOT information for a vehicle by registration number"""
    if not registration_number or pd.isna(registration_number):
        return "No registration number"
    
    registration_number = str(registration_number).strip().replace(" ", "")
    token = get_access_token()
    
    if not token:
        return "Error: No token"
    
    headers = {"Authorization": f"Bearer {token}", "X-API-Key": API_KEY}
    url = f"{API_BASE_URL}{registration_number}"
    
    try:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            data = resp.json()
            
            # Check for vehicles with MOT tests
            if "motTests" in data and data["motTests"]:
                test = sorted(data["motTests"], key=lambda x: x.get("completedDate", ""), reverse=True)[0]
                if "expiryDate" in test:
                    date = test["expiryDate"].split("T")[0]
                    try:
                        return datetime.strptime(date, "%Y-%m-%d").strftime("%d/%m/%Y")
                    except ValueError:
                        return date
                else:
                    return "No MOT expiry date found"
            
            # Check for newly registered vehicles with motTestDueDate
            elif "motTestDueDate" in data and data["motTestDueDate"]:
                date = data["motTestDueDate"].split("T")[0]
                try:
                    return datetime.strptime(date, "%Y-%m-%d").strftime("%d/%m/%Y")
                except ValueError:
                    return date
            
            # Check for very new vehicles with just registration date
            elif "registrationDate" in data and data["registrationDate"]:
                reg_date = data["registrationDate"].split("T")[0]
                try:
                    # Calculate MOT due date (3 years from registration for new vehicles)
                    reg_date_obj = datetime.strptime(reg_date, "%Y-%m-%d")
                    mot_due_date_obj = reg_date_obj + timedelta(days=3*365)  # Approximately 3 years
                    return f"New - Due: {mot_due_date_obj.strftime('%d/%m/%Y')}"
                except ValueError:
                    return f"New - Reg date: {reg_date}"
            
            else:
                return "No MOT info"
                
        elif resp.status_code == 404:
            return "Vehicle not found"
        else:
            return f"Error: {resp.status_code}"
            
    except Exception as e:
        logging.error(f"MOT request failed: {e}")
        return "Error during request"

def is_valid_reg_number(value):
    """Check if a value looks like a valid registration number"""
    if pd.isna(value): 
        return False
    
    value_str = str(value).strip()
    
    # Basic validation: at least 4 characters, contains both letters and numbers
    return len(value_str) >= 4 and any(c.isalpha() for c in value_str) and any(c.isdigit() for c in value_str)

def update_excel_file(excel_file):
    """Update MOT information in all sheets of the Excel file"""
    if not os.path.exists(excel_file):
        logging.error("Excel file not found")
        return False
    
    try:
        # Read all sheets at once with dtype=str to handle formatting issues
        dfs = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl', dtype=str)
        
        # Clean sheet names by stripping whitespace
        dfs = {name.strip(): df for name, df in dfs.items()}
        
        # Create output filename with timestamp to avoid overwriting
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f"{os.path.splitext(excel_file)[0]}_MOT_Updated_{timestamp}.xlsx"
        
        updated_dfs = {}
        total_vehicles = 0
        
        # Count total vehicles first for better progress reporting
        total_reg_count = 0
        for sheet_name, df in dfs.items():
            if df.empty or df.shape[1] == 0:
                continue
                
            first_col = df.columns[0]
            for idx, row in df.iterrows():
                reg = row[first_col]
                if is_valid_reg_number(reg):
                    total_reg_count += 1
        
        logging.info(f"Found {len(dfs)} sheets with {total_reg_count} valid registration numbers")
        
        # Process each sheet
        for sheet_name, df in dfs.items():
            logging.info(f"Processing sheet: {sheet_name}")
            
            # Skip empty sheets but include them in the output
            if df.empty or df.shape[1] == 0:
                updated_dfs[sheet_name] = df
                logging.info(f"Sheet {sheet_name} is empty or has no columns")
                continue
            
            # Get the first column (registration numbers)
            first_col = df.columns[0]
            logging.info(f"First column in sheet {sheet_name}: '{first_col}'")
            
            # Add MOT Due column if it doesn't exist (checking case-insensitive)
            if not any(col.strip().lower() == "mot due" for col in df.columns):
                df["MOT Due"] = ""
            
            # Process each row
            sheet_count = 0
            for idx, row in df.iterrows():
                reg = row[first_col]
                
                # Skip invalid registration numbers
                if not is_valid_reg_number(reg):
                    continue
                
                # Get MOT information
                mot_due = get_mot_information(reg)
                
                # Update the dataframe
                df.at[idx, "MOT Due"] = mot_due
                
                sheet_count += 1
                total_vehicles += 1
                
                # Show progress
                logging.info(f"Vehicle {total_vehicles}/{total_reg_count}: {reg} => {mot_due}")
                
                # Small delay to avoid overwhelming the API
                time.sleep(1)
            
            # Add the processed dataframe to our collection
            updated_dfs[sheet_name] = df
            logging.info(f"Completed sheet {sheet_name}: processed {sheet_count} vehicles")
        
        # Write all dataframes to the output file
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for name, df in updated_dfs.items():
                df.to_excel(writer, sheet_name=name, index=False)
                logging.info(f"Saved sheet: {name} ({df.shape[0]} rows)")
        
        logging.info(f"Finished processing {total_vehicles} vehicles across {len(updated_dfs)} sheets")
        logging.info(f"Results saved to {output_file}")
        return True
    
    except Exception as e:
        logging.error(f"Processing error: {e}")
        logging.debug(traceback.format_exc())
        return False

def main():
    """Main function"""
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("mot_checker.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    print("MOT Checker Tool")
    print("===============")
    print("Hello Team, time for coffee while I take care of the job")
    print("This tool checks MOT information for all vehicles in your Excel file.")
    print("It will process ALL sheets and ALL registration numbers.")
    print()
    
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = input("Enter Excel path: ").strip()
    
    print(f"\nProcessing file: {excel_file}")
    print("This may take several minutes depending on the number of vehicles.")
    print("Please wait...\n")
    
    if update_excel_file(excel_file):
        print("\n✅ Done. MOT information updated successfully.")
        print(f"Results saved to a new file with '_MOT_Updated' added to the filename.")
    else:
        print("\n❌ Failed. See the log file for details.")
    
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()
