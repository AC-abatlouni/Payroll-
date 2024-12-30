import pandas as pd
import logging
from datetime import datetime
import os
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
import re

def extract_dept_code(memo: str) -> Optional[str]:
    """Extract department code from memo, handling various formats like '27-', '27 -', ' 27-', etc."""
    if not memo:
        return None
    # Strip whitespace and standardize format
    memo = str(memo).strip()
    # Look for 2 digits at start, optionally followed by space/hyphen
    match = re.match(r'^\s*(\d{2})(?:\s*-|\s+|$)', memo)
    return match.group(1) if match else None

def sum_spiffs_for_dept(spiffs_df: pd.DataFrame, tech_name: str, dept_code: str) -> float:
    """Sum positive spiffs for a specific technician and department code."""
    dept_spiffs = spiffs_df[
        (spiffs_df['Technician'] == tech_name) & 
        (spiffs_df['Memo'].apply(lambda x: extract_dept_code(str(x)) == dept_code)) & 
        (spiffs_df['Amount'] > 0)  # Ignore negative spiffs
    ]
    
    if dept_spiffs.empty:
        return 0.0
        
    return dept_spiffs['Amount'].sum()

# Constants
LOCATION_ID = 'L100'
COMPANY_CODE = 'J6P'

# Department code mapping with descriptions
DEPARTMENT_CODES = {
    '20': {'code': '2000000', 'desc': 'HVAC SERVICE'},
    '21': {'code': '2100000', 'desc': 'HVAC INSTALL'},
    '22': {'code': '2200000', 'desc': 'MAINTENANCE MVP'},
    '24': {'code': '2400000', 'desc': 'OIL SERVICE'},
    '25': {'code': '2500000', 'desc': 'OIL MAINTENANCE MVP'},
    '27': {'code': '2700000', 'desc': 'HVAC DUCT CLEANING'},
    '30': {'code': '3000000', 'desc': 'PLUMBING SERVICE'},
    '31': {'code': '3100000', 'desc': 'PLUMBING INSTALL'},
    '33': {'code': '3300000', 'desc': 'PLUMBING DRAIN CLEANING'},
    '34': {'code': '3400000', 'desc': 'PLUMBING EXCAVATION'},
    '40': {'code': '4000000', 'desc': 'ELECTRICAL SERVICE'},
    '41': {'code': '4100000', 'desc': 'ELECTRICAL INSTALL'}
}

@dataclass
class PayrollEntry:
    company_code: str = COMPANY_CODE
    badge_id: str = ''
    date: str = ''
    amount: float = 0.0
    pay_code: str = ''
    dept: str = ''
    location_id: str = LOCATION_ID

def setup_logging() -> logging.Logger:
    """Configure logging with both file and console handlers."""
    logger = logging.getLogger('payroll_import')
    logger.setLevel(logging.DEBUG)
    
    # Create timestamp for log file
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f'payrollimport_{timestamp}.log'
    
    # File handler - Debug level with detailed formatting
    fh = logging.FileHandler(log_filename)
    fh.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )
    fh.setFormatter(file_formatter)
    
    # Console handler - Info level with simpler formatting
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    ch.setFormatter(console_formatter)
    
    logger.addHandler(fh)
    logger.addHandler(ch)
    
    return logger

def read_tech_department_data(file_path: str, logger: logging.Logger) -> pd.DataFrame:
    """Read technician department data from combined file."""
    try:
        logger.debug(f"Reading technician department data from {file_path}")
        tech_df = pd.read_excel(file_path, sheet_name='Sheet1_Tech')
        logger.debug(f"Successfully loaded {len(tech_df)} technician records")
        return tech_df
    except Exception as e:
        logger.error(f"Error reading technician department data: {str(e)}")
        raise

def determine_pay_code(business_unit: str) -> Optional[str]:
    """Determine pay code based on business unit description."""
    if pd.isna(business_unit):
        return None
    
    bu_upper = str(business_unit).upper()
    if 'ADMINISTRATIVE' in bu_upper:
        return None
    elif 'SERVICE' in bu_upper:
        return 'PC'
    else:
        return 'IC'

def process_paystats(paystats_file: str, tech_data: pd.DataFrame, target_date: str, logger: logging.Logger) -> List[PayrollEntry]:
    """Process paystats file to generate payroll entries."""
    logger.info("Processing paystats file for payroll entries...")
    payroll_entries = []

    try:
        # Load paystats data
        stats_df = pd.read_excel(paystats_file)
        logger.debug(f"Successfully loaded {len(stats_df)} records from paystats")

        # Load spiffs data from Direct Payroll Adjustments
        combined_file = r'C:\Users\abatlouni\Downloads\combined_data.xlsx'
        adj_df = pd.read_excel(combined_file, sheet_name='Direct Payroll Adjustments')

        # Process each technician
        for _, row in stats_df.iterrows():
            badge_id = row['Badge ID']
            tech_name = row['Technician']
            commission_rate = row['Commission Rate %'] / 100

            # Get business unit to determine pay code
            tech_bu = tech_data.loc[tech_data['Name'] == tech_name, 'Technician Business Unit'].iloc[0] \
                if not tech_data[tech_data['Name'] == tech_name].empty else None
            pay_code = determine_pay_code(tech_bu)

            if not pay_code:
                logger.debug(f"Skipping {tech_name} - No valid pay code")
                continue

            # Process each subdepartment
            for dept_code in DEPARTMENT_CODES.keys():
                total_col = f"{dept_code} Total"

                if total_col in row:
                    try:
                        # Get total amount for subdepartment
                        total_str = str(row[total_col])
                        if total_str == '$0.00' or not total_str:
                            continue

                        # Parse total amount and handle any formatting
                        total_amount = float(total_str.replace('$', '').replace(',', ''))
                        if total_amount <= 0:
                            continue

                        # Calculate spiffs total using helper function
                        spiffs_total = sum_spiffs_for_dept(adj_df, tech_name, dept_code)

                        # Adjust revenue by subtracting spiffs
                        adjusted_amount = round(total_amount - spiffs_total, 2)

                        # Calculate commission with proper rounding
                        commission_amount = round(adjusted_amount * commission_rate, 2)

                        if commission_amount > 0:
                            entry = PayrollEntry(
                                company_code=COMPANY_CODE,
                                badge_id=badge_id,
                                date=target_date,
                                amount=commission_amount,
                                pay_code=pay_code,
                                dept=DEPARTMENT_CODES[dept_code]['code'],
                                location_id=LOCATION_ID
                            )
                            payroll_entries.append(entry)

                            logger.debug(
                                f"{tech_name} - Dept {dept_code}: "
                                f"Total=${total_amount:,.2f}, "
                                f"Spiffs=${spiffs_total:,.2f}, "
                                f"Adjusted=${adjusted_amount:,.2f}, "
                                f"Commission=${commission_amount:,.2f}"
                            )

                    except (ValueError, KeyError) as e:
                        logger.warning(f"Error processing {total_col} for {tech_name}: {str(e)}")
                        continue

        logger.info(f"Generated {len(payroll_entries)} payroll entries")
        return payroll_entries

    except Exception as e:
        logger.error(f"Error processing paystats file: {str(e)}")
        raise

def process_gp_entries(combined_file: str, tech_data: pd.DataFrame, target_date: str, logger: logging.Logger) -> List[PayrollEntry]:
    """Process GP values from Invoices sheet to generate payroll entries."""
    logger.info("Processing GP entries from Invoices sheet...")
    payroll_entries = []

    try:
        # Read the Invoices sheet
        invoices_df = pd.read_excel(combined_file, sheet_name='Invoices')
        logger.debug(f"Loaded {len(invoices_df)} records from Invoices sheet")

        # Merge with tech data to get Payroll ID
        merged_df = invoices_df.merge(
            tech_data[['Name', 'Payroll ID']],
            left_on='Technician',  # Match technician name in Invoices
            right_on='Name',       # Match with Name in Sheet1_Tech
            how='left'
        )

        # Check if the merge added the necessary columns
        if 'Payroll ID' not in merged_df.columns:
            logger.error("Missing 'Payroll ID' after merging with tech data")
            raise KeyError("Payroll ID")

        # Group GP values by Technician, Business Unit, and Payroll ID
        grouped = merged_df.groupby(['Technician', 'Business Unit', 'Payroll ID'])['GP'].sum().reset_index()

        for _, row in grouped.iterrows():
            technician = row['Technician']
            business_unit = row['Business Unit']
            badge_id = row['Payroll ID']
            total_gp = row['GP']

            # Skip zero-amount entries
            if total_gp == 0:
                logger.debug(f"Skipping entry for {technician} (GP: {total_gp:.2f})")
                continue

            # Extract subdepartment code from the Business Unit (e.g., "20" from "HVAC SERVICE 20")
            subdepartment_code = business_unit.split()[-1] if isinstance(business_unit, str) else None

            # Map subdepartment code to department code
            dept_code = DEPARTMENT_CODES.get(subdepartment_code, {}).get('code')

            if not dept_code:
                logger.warning(f"Could not determine department code for business unit: {business_unit}")
                continue

            # Create PayrollEntry for the GP total
            entry = PayrollEntry(
                company_code=COMPANY_CODE,
                badge_id=badge_id if pd.notna(badge_id) else "UNKNOWN",
                date=target_date,
                amount=total_gp,
                pay_code='IC',  # Pay Code for GP entries
                dept=dept_code,
                location_id=LOCATION_ID
            )
            payroll_entries.append(entry)
            logger.debug(
                f"Created GP payroll entry for {technician} - "
                f"Business Unit: {business_unit}, Dept: {dept_code}, "
                f"Total GP: ${total_gp:.2f}"
            )

        logger.info(f"Generated {len(payroll_entries)} GP payroll entries")
        return payroll_entries

    except Exception as e:
        logger.error(f"Error processing GP entries: {str(e)}")
        raise




def save_payroll_file(entries: List[PayrollEntry], output_file: str, logger: logging.Logger):
    """Save payroll entries to Excel file with specific formatting."""
    try:
        # Create DataFrame with entries
        df = pd.DataFrame([{
            'Company Code': entry.company_code,
            'Badge ID': entry.badge_id,
            'Date': entry.date,
            'Amount': entry.amount,
            'Pay Code': entry.pay_code,
            'Dept': entry.dept,
            'Location ID': entry.location_id
        } for entry in entries])
        
        # Create Excel writer
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write DataFrame without index
            df.to_excel(writer, index=False)
            
            # Get the worksheet
            worksheet = writer.sheets['Sheet1']
            
            # Autofit columns
            for idx, col in enumerate(df.columns):
                # Get maximum length of column content
                max_length = max(
                    df[col].astype(str).apply(len).max(),  # Max length of values
                    len(str(col))  # Length of column header
                ) + 2  # Add padding
                
                # Get the column letter
                col_letter = get_column_letter(idx + 1)
                
                # Set column width
                worksheet.column_dimensions[col_letter].width = max_length
                
                # Format the header cell
                header_cell = worksheet[f"{col_letter}1"]
                header_cell.font = Font(bold=True)
                
                # Right-align Amount column
                if col == 'Amount':
                    for cell in worksheet[col_letter]:
                        cell.alignment = Alignment(horizontal='right')
                        if cell.row > 1:  # Skip header
                            cell.number_format = '#,##0.00'
        
        logger.info(f"Successfully saved payroll file: {output_file}")
        
    except Exception as e:
        logger.error(f"Error saving payroll file: {str(e)}")
        raise
    
def process_adjustments(combined_file: str, logger: logging.Logger) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Process adjustments from combined file."""
    logger.info("Processing adjustments data...")
    
    try:
        adj_df = pd.read_excel(combined_file, sheet_name='Direct Payroll Adjustments')
        logger.debug(f"Successfully loaded {len(adj_df)} adjustment records")
        
        # Initialize DataFrames for positive and negative adjustments
        positive_adj = []
        negative_adj = []
        
        for _, row in adj_df.iterrows():
            try:
                amount = float(str(row['Amount']).replace('$', '').replace(',', ''))
                memo = str(row['Memo']).strip()
                
                # Extract department code from memo
                dept_code = memo[:2] if memo[:2].isdigit() else '00'
                
                adj_record = {
                    'Technician': row['Technician'],
                    'Amount': abs(amount),
                    'Department': dept_code,
                    'Memo': memo,
                    'Type': 'Commission' if 'commission' in memo.lower() else 
                           'TGL' if 'tgl' in memo.lower() else 'Regular'
                }
                
                if amount >= 0:
                    positive_adj.append(adj_record)
                else:
                    negative_adj.append(adj_record)
                    
            except ValueError as e:
                logger.warning(f"Error processing adjustment record: {str(e)}")
                continue
        
        pos_df = pd.DataFrame(positive_adj)
        neg_df = pd.DataFrame(negative_adj)
        
        logger.info(f"Processed {len(pos_df)} positive and {len(neg_df)} negative adjustments")
        return pos_df, neg_df
        
    except Exception as e:
        logger.error(f"Error processing adjustments: {str(e)}")
        raise

def save_adjustments_files(pos_adj: pd.DataFrame, neg_adj: pd.DataFrame, 
                          pos_file: str, neg_file: str, logger: logging.Logger):
    """Save adjustment analysis files."""
    try:
        pos_adj.to_excel(pos_file, index=False)
        neg_adj.to_excel(neg_file, index=False)
        logger.info(f"Successfully saved adjustment files: {pos_file} and {neg_file}")
    except Exception as e:
        logger.error(f"Error saving adjustment files: {str(e)}")
        raise

def main():
    # Setup logging
    logger = setup_logging()
    logger.info("Starting payroll import process...")

    try:
        # Get user input for date
        target_date = input("Enter target date (MM/DD/YYYY): ")

        # Define file paths
        base_path = r'C:\Users\abatlouni\Downloads'
        combined_file = os.path.join(base_path, 'combined_data.xlsx')
        paystats_file = os.path.join(base_path, 'paystats.xlsx')

        # Output files
        payroll_file = os.path.join(base_path, 'payroll.xlsx')
        adj_pos_file = os.path.join(base_path, 'adjustments.xlsx')
        adj_neg_file = os.path.join(base_path, 'adjustments_negative.xlsx')

        # Process technician department data
        tech_data = read_tech_department_data(combined_file, logger)

        # Generate payroll entries from paystats
        payroll_entries = process_paystats(paystats_file, tech_data, target_date, logger)

        # Generate payroll entries from GP in Invoices
        gp_entries = process_gp_entries(combined_file, tech_data, target_date, logger)

        # Combine all payroll entries
        all_payroll_entries = payroll_entries + gp_entries

        # Process adjustments
        pos_adj, neg_adj = process_adjustments(combined_file, logger)

        # Save output files
        save_payroll_file(all_payroll_entries, payroll_file, logger)
        save_adjustments_files(pos_adj, neg_adj, adj_pos_file, adj_neg_file, logger)

        logger.info("Payroll import process completed successfully!")

    except Exception as e:
        logger.error(f"Fatal error in payroll import process: {str(e)}")
        raise

if __name__ == "__main__":
    main()
