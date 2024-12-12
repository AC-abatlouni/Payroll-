import pandas as pd
import logging
import sys
import os
import glob
import shutil
import re
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple, List
from dataclasses import dataclass
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from pathlib import Path

logger = logging.getLogger('commission_processor')

# Constants
LOCATION_ID = 'L100'
COMPANY_CODE = 'J6P'

# Column order for output
COLUMN_ORDER = [
    'Badge ID', 'Technician', 'Main Dept',
    # HVAC Department Subdepartments
    '20 Revenue', '20 Sales', '20 Total',
    '21 Revenue', '21 Sales', '21 Total',
    '22 Revenue', '22 Sales', '22 Total',
    '24 Revenue', '24 Sales', '24 Total',
    '25 Revenue', '25 Sales', '25 Total',
    '27 Revenue', '27 Sales', '27 Total',
    # HVAC Department Totals
    'HVAC Revenue', 'HVAC Sales', 'HVAC Spiffs', 'HVAC Total', 'HVAC Commission',
    # Plumbing Department Subdepartments
    '30 Revenue', '30 Sales', '30 Total',
    '31 Revenue', '31 Sales', '31 Total',
    '33 Revenue', '33 Sales', '33 Total',
    '34 Revenue', '34 Sales', '34 Total',
    # Plumbing Department Totals
    'Plumbing Revenue', 'Plumbing Sales', 'Plumbing Spiffs', 'Plumbing Total', 'Plumbing Commission',
    # Electric Department Subdepartments
    '40 Revenue', '40 Sales', '40 Total',
    '41 Revenue', '41 Sales', '41 Total',
    # Electric Department Totals
    'Electric Revenue', 'Electric Sales', 'Electric Spiffs', 'Electric Total', 'Electric Commission',
    # Performance Metrics
    'Completed Job Revenue', 'Tech-Sourced Install Sales',
    'Service Completion %', 'Install Contribution %',
    'Excused Hours', 'Spiffs', 'Valid TGLs', 'Avg Ticket $',
    'TGL Threshold Reduction', 'Base Threshold Scale', 'Adjusted Threshold Scale',
    'Commission Rate %', 'Total Revenue', 'Commissionable Revenue',
    'Status', 'Total Commission'
]

# Define threshold tables
HVAC_THRESHOLDS = {
    0: [7000, 8000, 9000, 10000],
    10: [7500, 8500, 9500, 11000],
    20: [8000, 9000, 10000, 11000],
    30: [9000, 10000, 11000, 12000],
    40: [10000, 11000, 12000, 13000],
    50: [12000, 14000, 16000, 18000],
    60: [14000, 16000, 18000, 20000],
    70: [15000, 17000, 19000, 21000],
    80: [16500, 18500, 20500, 23000],
    90: [18500, 20500, 23500, 26000],
    100: [22000, 24000, 26000, 29000]
}

PLUMBING_ELECTRICAL_THRESHOLDS = {
    0: [7000, 8000, 9000, 10000],
    10: [7500, 8500, 9500, 11000],
    20: [8000, 9000, 10000, 11000],
    30: [9000, 10000, 11000, 12000],
    40: [10000, 11000, 12000, 13000],
    50: [11500, 12500, 14000, 15000],
    60: [13000, 14000, 15500, 17000],
    70: [14500, 15500, 17000, 18000],
    80: [15500, 16500, 18000, 19000],
    90: [17000, 18500, 19500, 21000],
    100: [18000, 20000, 22000, 24000]
}

# Department code mapping
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

SUBDEPARTMENT_MAP = {
    '00': '00 - ADMINISTRATIVE',
    '41': '41 - ELECTRICAL INSTALL',
    '40': '40 - ELECTRICAL SERVICE',
    '42': '42 - GENERATOR MAINTENANCE',
    '27': '27 - HVAC DUCT CLEANING',
    '21': '21 - HVAC INSTALL',
    '23': '23 - HVAC SALES',
    '20': '20 - HVAC SERVICE',
    '22': '22 - MAINTENANCE MVP',
    '25': '25 - OIL MAINTENANCE MVP',
    '24': '24 - OIL SERVICE',
    '33': '33 - PLUMBING DRAIN CLEANING',
    '34': '34 - PLUMBING EXCAVATION',
    '31': '31 - PLUMBING INSTALL',
    '32': '32 - PLUMBING MAINTENANCE',
    '30': '30 - PLUMBING SERVICE'
}

# Technician Badge Map
TECH_BADGE_MAP = {
    "Andrew Wycoff": "J6P100665",
    "Andy Ventura": "J6P100622",
    "Artie Straniti": "J6P100524",
    "Brett Allen": "J6P100591",
    "Carter Bruce": "J6P100426",
    "Chris Smith": "J6P100430",
    "David Knox": "J6P100633",
    "Ethan Ficklin": "J6P100310",
    "Garrett Caine": "J6P100522",
    "Glenn Griffin": "J6P100297",
    "Hunter Stanley": "J6P100536",
    "Jacob Simpson": "J6P100512",
    "Jake West": "J6P100520",
    "John Williams": "J6P100529",
    "Josue Rodriguez": "J6P100553",
    "Justin Barron": "J6P100377",
    "Kevin Stanley": "J6P100655",
    "Pablo Silvas": "J6P100594",
    "Patrick Bowerman": "J6P100502",
    "Robert McGhee": "J6P100667",
    "Ronnie Bland": "J6P100133",
    "Shawn Hollingsworth": "J6P100696",
    "Stephen Starner": "J6P100521",
    "Thomas Shawaryn": "J6P100434",
    "Tim Miller": "J6P100485",
    "WT Settle": "J6P100283",
    "Will Winfree": "J6P100708"
}

# Dataclass definitions
@dataclass
class PayrollEntry:
    company_code: str = COMPANY_CODE
    badge_id: str = ''
    date: str = ''
    amount: float = 0.0
    pay_code: str = ''
    dept: str = ''
    location_id: str = LOCATION_ID

# Utility Functions
def setup_logging(name='commission_calculator'):
    """Configure logging with both file and console handlers."""
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f'{name}_{current_time}.log'
    
    # File handler - Debug level with detailed formatting
    fh = logging.FileHandler(log_filename)
    fh.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(file_formatter)
    
    # Console handler - Info level with simpler formatting
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    ch.setFormatter(console_formatter)
    
    # Add console filter for commission calculator
    if name == 'commission_calculator':
        class ConsoleFilter(logging.Filter):
            def filter(self, record):
                return record.levelno == logging.INFO and not any(x in record.msg for x in [
                    'Department Summary',
                    'Completed:',
                    'Sales:',
                    'Total:',
                    'Spiffs:',
                    'Commissionable Revenue:',
                    'Commission:',
                    'Processing technician'
                ])
        ch.addFilter(ConsoleFilter())
    
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger

def create_output_directory(base_path: str, start_of_week: datetime, end_of_week: datetime, logger: logging.Logger) -> str:
    """Create a directory for the output files with a name based on the week range."""
    # Format week range for folder name
    folder_name = f"Commission Output {start_of_week.strftime('%m_%d_%y')}-{end_of_week.strftime('%m_%d_%y')}"
    output_dir = os.path.join(base_path, folder_name)
    
    try:
        # Create the directory
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"Created output directory: {output_dir}")
        return output_dir
    except Exception as e:
        logger.error(f"Error creating output directory: {str(e)}")
        raise


def autofit_columns(worksheet):
    """Autofit column widths in Excel worksheet."""
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

def format_currency(amount):
    """Format number as currency string."""
    try:
        if isinstance(amount, str):
            amount = float(amount.replace('$', '').replace(',', ''))
        return f"${amount:,.2f}"
    except (ValueError, TypeError):
        return "$0.00"

def get_user_date() -> datetime:
    """Get date input from user."""
    while True:
        try:
            date_str = input("Enter a date (mm/dd/yy): ")
            return datetime.strptime(date_str, '%m/%d/%y')
        except ValueError:
            print("Invalid date format. Please use mm/dd/yy format (e.g., 03/15/24)")

def extract_subdepartment_code(business_unit):
    """Extract specific two-digit subdepartment code from business unit."""
    try:
        unit_str = str(business_unit).split('-')[0].strip()
        subdept = ''.join(filter(str.isdigit, unit_str))
        return subdept[:2] if len(subdept) >= 2 else '00'
    except:
        return '00'

def extract_department_number(business_unit):
    """Extract department number from business unit."""
    try:
        unit_str = str(business_unit).split('-')[0].strip()
        return int(''.join(filter(str.isdigit, unit_str)))
    except:
        return 0

def get_department_with_code(dept_num):
    """Get department name with code range."""
    if 20 <= dept_num <= 29:
        return 'HVAC (20-29)'
    elif 30 <= dept_num <= 39:
        return 'Plumbing (30-39)'
    elif 40 <= dept_num <= 49:
        return 'Electric (40-49)'
    return 'Unknown (0)'

def get_department_from_number(dept_num):
    """Get department name from number."""
    if 20 <= dept_num <= 29:
        return 'HVAC'
    elif 30 <= dept_num <= 39:
        return 'Plumbing'
    elif 40 <= dept_num <= 49:
        return 'Electric'
    return 'Unknown'

def extract_dept_code(memo: str) -> Optional[str]:
    """Extract department code from memo string."""
    if not memo:
        return None
    memo = str(memo).strip()
    match = re.match(r'^\s*(\d{2})(?:\s*-|\s+|$)', memo)
    return match.group(1) if match else None

def format_threshold_scale(thresholds):
    """Format threshold scale for display."""
    return (f"2%: ${thresholds[0]:,.0f} | "
            f"3%: ${thresholds[1]:,.0f} | "
            f"4%: ${thresholds[2]:,.0f} | "
            f"5%: ${thresholds[3]:,.0f}")

def extract_department_range(business_unit, logger):
    """Extract department range from business unit."""
    try:
        unit_number = int(''.join(filter(str.isdigit, business_unit)))
        base = (unit_number // 10) * 10
        if 20 <= base <= 20:
            return range(20, 30)
        elif 30 <= base <= 30:
            return range(30, 40)
        elif 40 <= base <= 40:
            return range(40, 50)
        return range(0, 0)
    except (ValueError, TypeError) as e:
        logger.warning(f"Could not extract department range from business unit: {business_unit}")
        return range(0, 0)

def is_same_department(source_unit, target_unit, logger=None):
    """Check if two business units are in the same department."""
    source_range = extract_department_range(source_unit, logger)
    target_range = extract_department_range(target_unit, logger)
    
    return (len(source_range) > 0 and 
            len(target_range) > 0 and 
            source_range.start == target_range.start)

def get_valid_tgls(file_path: str, tech_name: str) -> List[dict]:
    """Get valid TGLs for a technician."""
    try:
        tgl_df = pd.read_excel(file_path, sheet_name='Sheet1_TGL')
        logger.debug(f"Processing TGLs for {tech_name} from Sheet1_TGL")
        logger.debug(f"Available columns: {tgl_df.columns.tolist()}")
        
        valid_tgls = []
        tech_tgls = tgl_df[
            (tgl_df['Lead Generated By'] == tech_name) & 
            (tgl_df['Status'] != 'Canceled')
        ]
        
        for _, tgl in tech_tgls.iterrows():
            source_unit = str(tgl['Business Unit'])
            target_unit = str(tgl['Lead Generated from Business Unit'])
            
            if is_same_department(source_unit, target_unit, logger):
                valid_tgls.append({
                    'job_number': tgl.get('Job #', 'N/A'),
                    'status': tgl['Status'],
                    'business_unit': source_unit,
                    'target_unit': target_unit,
                    'created_date': tgl['Created Date']
                })
                logger.debug(f"Valid TGL found - Job #: {tgl.get('Job #', 'N/A')}, "
                           f"From: {source_unit} To: {target_unit}")
        
        return valid_tgls
        
    except Exception as e:
        logger.error(f"Error processing TGL data for {tech_name}: {str(e)}")
        return []

def get_spiffs_total(file_path: str, tech_name: str) -> tuple[float, dict[str, float]]:
    """Calculate total spiffs and department breakdown for a technician."""
    try:
        spiffs_df = pd.read_excel(file_path, sheet_name='Direct Payroll Adjustments')
        tech_spiffs = spiffs_df[spiffs_df['Technician'] == tech_name]
        
        spiffs_total = 0
        department_spiffs = {
            'HVAC': 0,
            'Plumbing': 0,
            'Electric': 0
        }
        
        for idx, spiff in tech_spiffs.iterrows():
            amount = float(str(spiff['Amount']).replace('$', '').replace(',', '').strip()) if pd.notnull(spiff['Amount']) else 0
            memo = str(spiff['Memo']).strip() if pd.notnull(spiff['Memo']) else ''
            
            # Skip if memo contains 'commission' or 'tgl' in any case
            if 'commission' in memo.lower() or 'tgl' in memo.lower():
                logger.debug(f"Skipping entry: ${amount:,.2f} - Memo: {memo}")
                continue
            
            if not memo[:2].isdigit():
                logger.error(f"Invalid memo format - must start with department number. Row {idx + 2}: {memo}")
                raise ValueError(f"Invalid memo format - must start with department number. Row {idx + 2}: {memo}")
            
            dept_num = int(memo[:2])
            if not (20 <= dept_num <= 29 or 30 <= dept_num <= 39 or 40 <= dept_num <= 49):
                logger.error(f"Invalid department number in memo (must be 20-29, 30-39, or 40-49). Row {idx + 2}: {memo}")
                raise ValueError(f"Invalid department number in memo (must be 20-29, 30-39, or 40-49). Row {idx + 2}: {memo}")
            
            memo_content = memo[2:].strip().lstrip('-').strip()
            if not memo_content:
                logger.warning(f"Warning: Memo contains no description after department number. Row {idx + 2}: {memo}")
            
            if 20 <= dept_num <= 29:
                department_spiffs['HVAC'] += amount
                logger.debug(f"Added HVAC spiff: ${amount:,.2f} - Memo: {memo}")
            elif 30 <= dept_num <= 39:
                department_spiffs['Plumbing'] += amount
                logger.debug(f"Added Plumbing spiff: ${amount:,.2f} - Memo: {memo}")
            elif 40 <= dept_num <= 49:
                department_spiffs['Electric'] += amount
                logger.debug(f"Added Electric spiff: ${amount:,.2f} - Memo: {memo}")
            
            spiffs_total += amount
        
        return spiffs_total, department_spiffs
        
    except Exception as e:
        logger.error(f"Error processing spiffs data for {tech_name}: {str(e)}")
        raise

def get_excused_hours(file_path: str, base_date: datetime, sheet_name: str = '2024') -> Dict[str, int]:
    """Get excused hours from time off sheet."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        start_of_week = base_date - timedelta(days=base_date.weekday())
        end_of_week = start_of_week + timedelta(days=4)
        
        def format_date(date):
            day = date.day
            if 4 <= day <= 20 or 24 <= day <= 30:
                suffix = "th"
            else:
                suffix = ["st", "nd", "rd"][day % 10 - 1]
            return f"{date.strftime('%B')} {day}{suffix}"
        
        target_week = f"{format_date(start_of_week)} - {format_date(end_of_week)}"
        logger.debug(f"Looking for week range: {target_week}")
        
        start_col = None
        for row_idx in [0, 1]:
            for col_idx, header in enumerate(df.iloc[row_idx]):
                if isinstance(header, str) and target_week.strip() == header.strip():
                    start_col = col_idx
                    logger.debug(f"Found target week in row {row_idx}, column {col_idx}")
                    break
            if start_col is not None:
                break
        
        if start_col is None:
            logger.warning(f"Week range '{target_week}' not found in time off sheet")
            return {}
        
        week_columns = list(range(start_col, start_col + 5))
        hours_summary = {}
        
        for index in range(2, len(df)):
            row = df.iloc[index]
            technician_name = row[0]
            
            if pd.isna(technician_name) or str(technician_name).strip() == '':
                continue
            
            technician_name = str(technician_name).strip()
            total_hours = 0
            
            for col in week_columns:
                try:
                    cell_value = str(row[col]).strip().lower()
                    if cell_value in ['x', 'r', 'v']:
                        total_hours += 8
                        logger.debug(f"Found time off marker '{cell_value}' for {technician_name} in column {col}")
                except Exception as e:
                    logger.warning(f"Error processing cell for {technician_name} in column {col}: {str(e)}")
                    continue
            
            if total_hours > 0:
                hours_summary[technician_name] = total_hours
                logger.debug(f"Found {total_hours} excused hours for {technician_name}")
        
        return hours_summary
        
    except Exception as e:
        logger.error(f"Error analyzing time off data: {str(e)}")
        return {}

def calculate_box_metrics(data: pd.DataFrame, tech_name: str, base_date: datetime) -> Tuple[float, float, float, Dict[str, Dict[str, float]]]:
    """Calculate Box A (CJR), Box B (TSIS), and Box C (Total) metrics with subdepartment breakdowns."""
    subdept_breakdown = {
        'completed': {code: 0.0 for code in SUBDEPARTMENT_MAP.keys()},
        'sales': {code: 0.0 for code in SUBDEPARTMENT_MAP.keys()},
        'total': {code: 0.0 for code in SUBDEPARTMENT_MAP.keys()}
    }

    # Calculate week range
    start_of_week = base_date - timedelta(days=base_date.weekday())
    end_of_week = start_of_week + timedelta(days=6)
    
    logger.debug(f"\nCalculating metrics for {tech_name} for week {start_of_week.strftime('%Y-%m-%d')} to {end_of_week.strftime('%Y-%m-%d')}")
    
    # Ensure datetime format
    if not pd.api.types.is_datetime64_any_dtype(data['Invoice Date']):
        data['Invoice Date'] = pd.to_datetime(data['Invoice Date'])

    # Get primary tech jobs within date range
    primary_jobs = data[
        (data['Primary Technician'] == tech_name) &
        (data['Invoice Date'].dt.date >= start_of_week.date()) &
        (data['Invoice Date'].dt.date <= end_of_week.date())
    ]
    box_a = primary_jobs['Jobs Total Revenue'].fillna(0).sum()
    
    # Get sales jobs within date range
    sold_jobs = data[
        (data['Sold By'] == tech_name) & 
        (data['Primary Technician'] != tech_name) &
        (data['Invoice Date'].dt.date >= start_of_week.date()) &
        (data['Invoice Date'].dt.date <= end_of_week.date())
    ]
    box_b = sold_jobs['Jobs Total Revenue'].fillna(0).sum()
    
    # Calculate subdepartment breakdowns
    for _, row in primary_jobs.iterrows():
        revenue = row['Jobs Total Revenue'] or 0
        subdept = extract_subdepartment_code(row.get('Business Unit', ''))
        subdept_breakdown['completed'][subdept] += revenue
        subdept_breakdown['total'][subdept] += revenue
        
    for _, row in sold_jobs.iterrows():
        revenue = row['Jobs Total Revenue'] or 0
        subdept = extract_subdepartment_code(row.get('Business Unit', ''))
        subdept_breakdown['sales'][subdept] += revenue
        subdept_breakdown['total'][subdept] += revenue

    box_c = box_a + box_b
    
    return box_a, box_b, box_c, subdept_breakdown

def calculate_percentages(box_a: float, box_c: float) -> Tuple[int, int]:
    """Calculate Service Completion and Install Contribution percentages."""
    if box_c == 0:
        return 0, 0
        
    raw_scp = (box_a / box_c * 100)
    raw_icp = 100 - raw_scp
    
    raw_scp = max(0, raw_scp)
    raw_icp = max(0, raw_icp)
    
    total = raw_scp + raw_icp
    if total != 0:
        raw_scp = (raw_scp / total) * 100
        raw_icp = (raw_icp / total) * 100
        
    scp = round(raw_scp / 10) * 10
    icp = round(raw_icp / 10) * 10
    
    if scp + icp != 100:
        if scp > icp:
            scp = 100
            icp = 0
        else:
            scp = 0
            icp = 100
    
    return int(scp), int(icp)

def calculate_average_ticket_value(data: pd.DataFrame, tech_name: str, box_a: float, box_b: float, base_date: datetime, logger: logging.Logger) -> Dict[str, float]:
    """Calculate average ticket value using total revenue divided by opportunity count."""
    avg_tickets = {'overall': 0.0}
    
    # Calculate week range
    start_of_week = base_date - timedelta(days=base_date.weekday())
    end_of_week = start_of_week + timedelta(days=6)
    
    logger.debug(f"\nCalculating average ticket value for {tech_name}")
    logger.debug(f"Week range: {start_of_week.date()} to {end_of_week.date()}")
    
    if not pd.api.types.is_datetime64_any_dtype(data['Invoice Date']):
        data['Invoice Date'] = pd.to_datetime(data['Invoice Date'])

    # Get completed jobs within date range to count opportunities
    completed_jobs = data[
        (data['Primary Technician'] == tech_name) &
        (data['Invoice Date'].dt.date >= start_of_week.date()) &
        (data['Invoice Date'].dt.date <= end_of_week.date())
    ]
    
    # Get count of opportunity jobs
    opportunity_jobs = completed_jobs[completed_jobs['Opportunity'] == True]
    opportunity_count = len(opportunity_jobs)
    
    # Calculate average ticket using total revenue
    total_revenue = box_a + box_b
    avg_ticket = round(total_revenue / opportunity_count, 2) if opportunity_count > 0 else 0
    avg_tickets['overall'] = avg_ticket
    
    logger.debug(f"Total Revenue (CJR + TSIS): ${total_revenue:,.2f}")
    logger.debug(f"True Opportunity Count: {opportunity_count}")
    logger.debug(f"Average Ticket: ${avg_ticket:,.2f}")
    
    # Log opportunity jobs for reference
    logger.debug("\nOpportunity Jobs:")
    for _, job in opportunity_jobs.iterrows():
        invoice = job.get('Invoice #', 'N/A')
        customer_name = job.get('Customer Name', 'Unknown')
        invoice_date = job['Invoice Date'].strftime('%m/%d/%y')
        logger.debug(f"  Invoice #{invoice} - {invoice_date} - {customer_name}")
    
    return avg_tickets

def format_department_revenue(revenue_data: Dict[str, Dict[str, float]], 
                            commission_rate: float,
                            department_spiffs: Dict[str, float],
                            subdept_breakdown: Dict[str, Dict[str, float]]) -> Dict[str, str]:
    """Format department and subdepartment revenue data into strings for Excel output."""
    formatted = {}
    
    # Format subdepartment totals
    for subdept_code in ['20', '21', '22', '24', '25', '27', '30', '31', '33', '34', '40', '41']:
        completed = subdept_breakdown['completed'].get(subdept_code, 0)
        sales = subdept_breakdown['sales'].get(subdept_code, 0)
        total = subdept_breakdown['total'].get(subdept_code, 0)
        
        formatted[f"{subdept_code} Revenue"] = f"${completed:,.2f}"
        formatted[f"{subdept_code} Sales"] = f"${sales:,.2f}"
        formatted[f"{subdept_code} Total"] = f"${total:,.2f}"
    
    # Format main department totals
    for dept in ['HVAC', 'Plumbing', 'Electric']:
        completed = revenue_data['completed'][dept]
        sales = revenue_data['sales'][dept]
        combined = revenue_data['combined'][dept]
        dept_spiffs = department_spiffs[dept]
        adjusted_combined = combined - dept_spiffs
        
        formatted[f"{dept} Revenue"] = f"${completed:,.2f}"
        formatted[f"{dept} Sales"] = f"${sales:,.2f}"
        formatted[f"{dept} Spiffs"] = f"${dept_spiffs:,.2f}"
        formatted[f"{dept} Total"] = f"${adjusted_combined:,.2f}"
        formatted[f"{dept} Commission"] = f"${(adjusted_combined * commission_rate):,.2f}"
    
    return formatted

def calculate_department_revenue(data: pd.DataFrame, tech_name: str, base_date: datetime) -> Dict[str, Dict[str, float]]:
    """Calculate department revenue using the same date filtering."""
    # Calculate week range
    start_of_week = base_date - timedelta(days=base_date.weekday())
    end_of_week = start_of_week + timedelta(days=6)
    
    revenue_by_dept = {
        'completed': {'HVAC': 0.0, 'Plumbing': 0.0, 'Electric': 0.0, 'Unknown': 0.0},
        'sales': {'HVAC': 0.0, 'Plumbing': 0.0, 'Electric': 0.0, 'Unknown': 0.0},
        'combined': {'HVAC': 0.0, 'Plumbing': 0.0, 'Electric': 0.0, 'Unknown': 0.0}
    }
    
    if not pd.api.types.is_datetime64_any_dtype(data['Invoice Date']):
        data['Invoice Date'] = pd.to_datetime(data['Invoice Date'])

    # Get completed jobs within date range
    completed_jobs = data[
        (data['Primary Technician'] == tech_name) &
        (data['Invoice Date'].dt.date >= start_of_week.date()) &
        (data['Invoice Date'].dt.date <= end_of_week.date())
    ]
    
    # Get sales within date range
    sold_jobs = data[
        (data['Sold By'] == tech_name) & 
        (data['Primary Technician'] != tech_name) &
        (data['Invoice Date'].dt.date >= start_of_week.date()) &
        (data['Invoice Date'].dt.date <= end_of_week.date())
    ]
    
    # Calculate revenues
    for _, job in completed_jobs.iterrows():
        dept_num = extract_department_number(str(job.get('Business Unit', '')))
        dept = get_department_from_number(dept_num)
        revenue = job.get('Jobs Total Revenue', 0) or 0
        revenue_by_dept['completed'][dept] += revenue
        revenue_by_dept['combined'][dept] += revenue

    for _, job in sold_jobs.iterrows():
        dept_num = extract_department_number(str(job.get('Business Unit', '')))
        dept = get_department_from_number(dept_num)
        revenue = job.get('Jobs Total Revenue', 0) or 0
        revenue_by_dept['sales'][dept] += revenue
        revenue_by_dept['combined'][dept] += revenue

    return revenue_by_dept

def get_commission_rate(total_revenue: float, flipped_percent: float, department: str, 
                       excused_hours: int, tgl_reduction: float) -> Tuple[float, list, list]:
    """Calculate commission rate and thresholds based on revenue and department."""
    logger.debug(f"\nDetailed threshold calculation:")
    
    flipped_percent = min(100, max(0, int(round(flipped_percent / 10) * 10)))
    
    if department in ['Electric', 'Plumbing']:
        thresholds = PLUMBING_ELECTRICAL_THRESHOLDS
    else:
        thresholds = HVAC_THRESHOLDS
    
    tier_thresholds = thresholds[flipped_percent].copy()
    logger.debug(f"Using {department} thresholds for ICP {flipped_percent}%: {tier_thresholds}")
    
    days_off = min(5, excused_hours / 8)
    reduction_factor = max(0, 1 - (0.20 * days_off))
    logger.debug(f"Time off reduction factor: {reduction_factor} ({days_off} days)")
    
    # First apply time off reduction
    time_off_adjusted = [threshold * reduction_factor for threshold in tier_thresholds]
    logger.debug(f"After time off adjustment: {time_off_adjusted}")
    
    # Then apply TGL reduction
    adjusted_thresholds = [max(0, threshold - tgl_reduction) for threshold in time_off_adjusted]
    logger.debug(f"After TGL reduction: {adjusted_thresholds}")
    
    # Determine commission rate based on highest threshold met
    if total_revenue >= adjusted_thresholds[3]:
        rate = 0.05
    elif total_revenue >= adjusted_thresholds[2]:
        rate = 0.04
    elif total_revenue >= adjusted_thresholds[1]:
        rate = 0.03
    elif total_revenue >= adjusted_thresholds[0]:
        rate = 0.02
    else:
        rate = 0
    
    logger.debug(f"Final commission rate: {rate}")
    return rate, adjusted_thresholds, tier_thresholds

def process_commission_calculations(data: pd.DataFrame, tech_data: pd.DataFrame, 
                                 file_path: str, base_date: datetime,
                                 excused_hours_dict: Dict[str, int]) -> pd.DataFrame:
    """Process commission calculations for all technicians."""
    results = []
    technicians_to_track = list(TECH_BADGE_MAP.keys())
    tech_dept_map = dict(zip(tech_data['Name'], tech_data['Technician Business Unit']))

    for tech_name in technicians_to_track:
        logger.info(f"\nProcessing technician: {tech_name}")
        badge_id = TECH_BADGE_MAP.get(tech_name, '')
        
        # Basic calculations with subdepartment tracking
        box_a, box_b, box_c, subdept_breakdown = calculate_box_metrics(data, tech_name, base_date)
        scp, icp = calculate_percentages(box_a, box_c)
        
        # Calculate department-specific revenue
        dept_revenue = calculate_department_revenue(data, tech_name, base_date)
        
        # Get spiffs with department breakdown
        spiffs_total, department_spiffs = get_spiffs_total(file_path, tech_name)
        
        # Get valid TGLs
        valid_tgls = get_valid_tgls(file_path, tech_name)
        
        # Calculate average ticket values
        avg_tickets = calculate_average_ticket_value(data, tech_name, box_a, box_b, base_date, logger)
        default_ticket = 0.0
        avg_ticket_value = avg_tickets.get('overall', default_ticket) if avg_tickets else default_ticket
        
        # Get department
        dept_unit_str = tech_dept_map.get(tech_name, '0')
        dept_num = extract_department_number(dept_unit_str)
        department = get_department_from_number(dept_num)
        department_with_code = get_department_with_code(dept_num)

        # Calculate TGL threshold reduction
        tgl_reduction = avg_ticket_value * len(valid_tgls) if avg_ticket_value > 0 else 0
        
        excused_hours = excused_hours_dict.get(tech_name, 0)
        
        # Calculate commission rate using ICP
        commission_rate, adjusted_thresholds, base_thresholds = get_commission_rate(
            box_c, icp, department, excused_hours, tgl_reduction
        )
        
        # Format threshold scales
        base_threshold_scale = format_threshold_scale(base_thresholds)
        adjusted_threshold_scale = format_threshold_scale(adjusted_thresholds)
        
        # Format revenue data including subdepartments
        formatted_dept_data = format_department_revenue(
            dept_revenue,
            commission_rate,
            department_spiffs,
            subdept_breakdown
        )
        
        # Calculate total commissionable revenue
        commissionable_revenue = box_c - spiffs_total
        final_commission = commissionable_revenue * commission_rate
        
        # Build result dictionary
        result = {
            'Badge ID': badge_id,
            'Technician': tech_name,
            'Main Dept': department_with_code,
            'Total Revenue': box_c,
            'Completed Job Revenue': box_a,
            'Tech-Sourced Install Sales': box_b,
            'Service Completion %': scp,
            'Install Contribution %': icp,
            'Excused Hours': excused_hours,
            'Spiffs': spiffs_total,
            'Valid TGLs': len(valid_tgls),
            'Avg Ticket $': avg_ticket_value,
            'TGL Threshold Reduction': tgl_reduction,
            'Base Threshold Scale': base_threshold_scale,
            'Adjusted Threshold Scale': adjusted_threshold_scale,
            'Commissionable Revenue': commissionable_revenue,
            'Commission Rate %': commission_rate * 100,
            'Total Commission': round(final_commission, 2),
            'Status': f"Qualified for {commission_rate*100}% tier" if commission_rate > 0 else "Did not qualify"
        }
        
        # Add formatted department and subdepartment data
        result.update(formatted_dept_data)
        results.append(result)

    results_df = pd.DataFrame(results)
    
    # Ensure all required columns exist and are in correct order
    for col in COLUMN_ORDER:
        if col not in results_df.columns:
            results_df[col] = ''
    
    return results_df[COLUMN_ORDER]

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

def sum_spiffs_for_dept(spiffs_df: pd.DataFrame, tech_name: str, dept_code: str) -> float:
    """Sum spiffs for a specific technician and department code."""
    dept_spiffs = spiffs_df[
        (spiffs_df['Technician'] == tech_name) & 
        (spiffs_df['Memo'].apply(lambda x: extract_dept_code(str(x)) == dept_code)) &
        (~spiffs_df['Memo'].str.lower().str.contains('tgl')) &
        (~spiffs_df['Memo'].str.lower().str.contains('commission'))
    ]
    
    if dept_spiffs.empty:
        return 0.0
        
    amounts = dept_spiffs['Amount'].apply(
        lambda x: float(str(x).replace('$', '').replace(',', '')) if pd.notnull(x) else 0.0
    )
    
    return amounts.sum()

def process_paystats(paystats_file: str, tech_data: pd.DataFrame, target_date: str, logger: logging.Logger) -> List[PayrollEntry]:
    """Process paystats file to generate payroll entries."""
    logger.info("Processing paystats file for payroll entries...")
    payroll_entries = []

    try:
        # Load paystats data
        stats_df = pd.read_excel(paystats_file)
        logger.debug(f"Successfully loaded {len(stats_df)} records from paystats")

        # Load spiffs data from Direct Payroll Adjustments
        combined_file = os.path.join(os.path.expanduser(r'~\Downloads'), 'combined_data.xlsx')
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
                try:
                    # Initialize variables
                    total_amount = 0.0
                    spiffs_total = sum_spiffs_for_dept(adj_df, tech_name, dept_code)

                    # Get total amount if it exists
                    if total_col in row:
                        total_str = str(row[total_col])
                        if total_str:
                            try:
                                total_amount = float(total_str.replace('$', '').replace(',', ''))
                            except ValueError:
                                total_amount = 0.0

                    # Calculate adjusted amount
                    adjusted_amount = round(total_amount - spiffs_total, 2)
                    final_amount = round(adjusted_amount * commission_rate, 2)

                    # Skip if final amount is zero
                    if final_amount == 0:
                        continue

                    entry = PayrollEntry(
                        company_code=COMPANY_CODE,
                        badge_id=badge_id,
                        date=target_date,
                        amount=final_amount,
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
                        f"Final Amount=${final_amount:,.2f}"
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
            left_on='Technician',
            right_on='Name',
            how='left'
        )

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

            if total_gp == 0:
                logger.debug(f"Skipping entry for {technician} (GP: {total_gp:.2f})")
                continue

            subdepartment_code = business_unit.split()[-1] if isinstance(business_unit, str) else None
            dept_code = DEPARTMENT_CODES.get(subdepartment_code, {}).get('code')

            if not dept_code:
                logger.warning(f"Could not determine department code for business unit: {business_unit}")
                continue

            entry = PayrollEntry(
                company_code=COMPANY_CODE,
                badge_id=badge_id if pd.notna(badge_id) else "UNKNOWN",
                date=target_date,
                amount=total_gp,
                pay_code='IC',
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

def process_adjustments(combined_file: str, logger: logging.Logger) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Process adjustments data and split into positive and negative adjustments."""
    logger.info("Processing adjustments data...")
    
    try:
        adj_df = pd.read_excel(combined_file, sheet_name='Direct Payroll Adjustments')
        logger.debug(f"Successfully loaded {len(adj_df)} adjustment records")
        
        positive_adj = []
        negative_adj = []
        
        for _, row in adj_df.iterrows():
            try:
                amount = float(str(row['Amount']).replace('$', '').replace(',', ''))
                memo = str(row['Memo']).strip()
                
                dept_code = memo[:2] if memo[:2].isdigit() else '00'
                dept_desc = DEPARTMENT_CODES.get(dept_code, {'desc': 'Unknown'})['desc']
                dept_combined = f"{dept_code} - {dept_desc}"
                
                adj_record = {
                    'Technician': row['Technician'],
                    'Department': dept_combined,
                    'Memo': memo,
                    'Type': 'Commission' if 'commission' in memo.lower() else 
                           'TGL' if 'tgl' in memo.lower() else 'Spiff'
                }
                
                if amount >= 0:
                    if not ('tgl' in memo.lower() or 'commission' in memo.lower()):
                        adj_record['Amount'] = amount
                        positive_adj.append(adj_record)
                else:
                    adj_record['Amount'] = amount
                    negative_adj.append(adj_record)
                    logger.debug(f"Negative adjustment found: {adj_record}")
                    
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

def save_payroll_file(entries: List[PayrollEntry], output_file: str, logger: logging.Logger):
    """Save payroll entries to Excel file with specific formatting."""
    try:
        df = pd.DataFrame([{
            'Company Code': entry.company_code,
            'Badge ID': entry.badge_id,
            'Date': entry.date,
            'Amount': entry.amount,
            'Pay Code': entry.pay_code,
            'Dept': entry.dept,
            'Location ID': entry.location_id
        } for entry in entries])
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            
            for idx, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(str(col))
                ) + 2
                
                col_letter = get_column_letter(idx + 1)
                worksheet.column_dimensions[col_letter].width = max_length
                
                header_cell = worksheet[f"{col_letter}1"]
                header_cell.font = Font(bold=True)
                
                if col == 'Amount':
                    for cell in worksheet[col_letter]:
                        cell.alignment = Alignment(horizontal='right')
                        if cell.row > 1:
                            cell.number_format = '#,##0.00'
        
        logger.info(f"Successfully saved payroll file: {output_file}")
        
    except Exception as e:
        logger.error(f"Error saving payroll file: {str(e)}")
        raise

def save_adjustments_files(pos_adj: pd.DataFrame, neg_adj: pd.DataFrame, 
                          pos_file: str, neg_file: str, logger: logging.Logger):
    """Save adjustment analysis files with autofit columns."""
    try:
        with pd.ExcelWriter(pos_file, engine='openpyxl') as writer:
            pos_adj.to_excel(writer, index=False)
            autofit_columns(writer.sheets['Sheet1'])
            
        with pd.ExcelWriter(neg_file, engine='openpyxl') as writer:
            neg_adj.to_excel(writer, index=False)
            autofit_columns(writer.sheets['Sheet1'])
            
        logger.info(f"Successfully saved adjustment files: {pos_file} and {neg_file}")
        
    except Exception as e:
        logger.error(f"Error saving adjustment files: {str(e)}")
        raise

def validate_required_files(directory):
    """Validate that all required files are present before processing."""
    try:
        # Check for UUID file
        if not glob.glob(os.path.join(directory, "????????-????-????-????-????????????.xlsx")):
            raise FileNotFoundError("UUID file not found")
            
        # Check for Jobs Report
        if not glob.glob(os.path.join(directory, "Copy of Jobs Report for Performance -DE2_Dated *.xlsx")):
            raise FileNotFoundError("Jobs Report file not found")
            
        # Check for Tech Department file
        if not glob.glob(os.path.join(directory, "Technician Department_Dated *.xlsx")):
            raise FileNotFoundError("Tech Department file not found")
            
        # Check for Time Off file
        time_off_file = os.path.join(directory, "Approved_Time_Off 2023.xlsx")
        if not os.path.exists(time_off_file):
            raise FileNotFoundError("Time Off file not found")
            
        return True
        
    except FileNotFoundError as e:
        print(f"\nError: {str(e)}")
        print("Please ensure all required files are in the directory before running.")
        sys.exit(1)

def find_latest_files(directory):
    """Find the latest version of each required file."""
    try:
        uuid_file = glob.glob(os.path.join(directory, "????????-????-????-????-????????????.xlsx"))[0]
    except IndexError:
        raise FileNotFoundError("UUID file not found")
    
    try:
        jobs_file = max(glob.glob(os.path.join(directory, "Copy of Jobs Report for Performance -DE2_Dated *.xlsx")))
    except ValueError:
        raise FileNotFoundError("Jobs Report file not found")
    
    try:
        tech_file = max(glob.glob(os.path.join(directory, "Technician Department_Dated *.xlsx")))
    except ValueError:
        raise FileNotFoundError("Tech Department file not found")
    
    time_off_file = os.path.join(directory, "Approved_Time_Off 2023.xlsx")
    if not os.path.exists(time_off_file):
        raise FileNotFoundError("Time Off file not found")
    
    tgl_files = glob.glob(os.path.join(directory, "TGLs Set _Dated *.xlsx"))
    tgl_file = max(tgl_files) if tgl_files else None
    
    return {
        'uuid': uuid_file,
        'jobs': jobs_file,
        'tech': tech_file,
        'time_off': time_off_file,
        'tgl': tgl_file
    }

def combine_workbooks(directory, output_file):
    """Combine all workbooks into a single file."""
    validate_required_files(directory)
    files = find_latest_files(directory)
    
    # Start by copying the UUID file as our base
    shutil.copy2(files['uuid'], output_file)
    
    # Open target workbook
    target_wb = load_workbook(filename=output_file)
    
    # AutoFit UUID file sheets
    for sheet in target_wb.sheetnames:
        autofit_columns(target_wb[sheet])
    
    # Copy from Jobs Report
    source_wb = load_workbook(filename=files['jobs'], data_only=True)
    source_ws = source_wb['Sheet1']
    if 'Sheet1' in target_wb.sheetnames:
        target_wb.remove(target_wb['Sheet1'])
    target_wb.create_sheet('Sheet1')
    target_ws = target_wb['Sheet1']
    
    for row in source_ws:
        for cell in row:
            target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
    autofit_columns(target_ws)
    
    # Copy from Tech Department
    source_wb = load_workbook(filename=files['tech'], data_only=True)
    source_ws = source_wb['Sheet1']
    target_wb.create_sheet('Sheet1_Tech')
    target_ws = target_wb['Sheet1_Tech']
    
    for row in source_ws:
        for cell in row:
            target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
    autofit_columns(target_ws)
    
    # Copy from Time Off
    source_wb = load_workbook(filename=files['time_off'], data_only=True)
    source_ws = source_wb['2024']
    if '2024' in target_wb.sheetnames:
        target_wb.remove(target_wb['2024'])
    target_wb.create_sheet('2024')
    target_ws = target_wb['2024']
    
    for row in source_ws:
        for cell in row:
            target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
    autofit_columns(target_ws)
    
    # Copy from TGLs Set if file exists
    if files['tgl']:
        source_wb = load_workbook(filename=files['tgl'], data_only=True)
        source_ws = source_wb['Sheet1']
        if 'Sheet1_TGL' in target_wb.sheetnames:
            target_wb.remove(target_wb['Sheet1_TGL'])
        target_wb.create_sheet('Sheet1_TGL')
        target_ws = target_wb['Sheet1_TGL']
        
        for row in source_ws:
            for cell in row:
                target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
        autofit_columns(target_ws)
    
    target_wb.save(output_file)
    return files

def process_calculations(base_path: str, output_dir: str, logger: logging.Logger,
                         start_of_week: datetime, end_of_week: datetime):
    """Process all calculations and generate output files."""
    try:
        combined_file = os.path.join(output_dir, 'combined_data.xlsx')
        paystats_file = os.path.join(output_dir, 'paystats.xlsx')

        # Read necessary data
        data = pd.read_excel(combined_file, sheet_name='Sheet1')
        tech_data = read_tech_department_data(combined_file, logger)

        # Calculate commission results for the specified week
        excused_hours_dict = get_excused_hours(combined_file, start_of_week)
        results_df = process_commission_calculations(
            data, tech_data, combined_file, start_of_week, excused_hours_dict
        )

        # Save results to paystats file
        with pd.ExcelWriter(paystats_file, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='Technician Revenue Totals', index=False)
            autofit_columns(writer.sheets['Technician Revenue Totals'])

        logger.info("Commission calculations completed")

    except Exception as e:
        logger.error(f"Error in calculations: {str(e)}")
        raise

def process_payroll(base_path: str, output_dir: str, base_date: datetime, logger: logging.Logger):
    """Process payroll and adjustments."""
    try:
        combined_file = os.path.join(output_dir, 'combined_data.xlsx')
        paystats_file = os.path.join(output_dir, 'paystats.xlsx')
        payroll_file = os.path.join(output_dir, 'payroll.xlsx')
        adj_pos_file = os.path.join(output_dir, 'adjustments.xlsx')
        adj_neg_file = os.path.join(output_dir, 'adjustments_negative.xlsx')
        
        # Format target date
        target_date = base_date.strftime('%m/%d/%Y')
        
        # Process technician department data
        tech_data = read_tech_department_data(combined_file, logger)
        
        # Generate payroll entries from paystats
        payroll_entries = process_paystats(paystats_file, tech_data, target_date, logger)
        
        # Generate payroll entries from GP
        gp_entries = process_gp_entries(combined_file, tech_data, target_date, logger)
        
        # Combine all payroll entries
        all_payroll_entries = payroll_entries + gp_entries
        
        # Process adjustments
        pos_adj, neg_adj = process_adjustments(combined_file, logger)
        
        # Save output files
        save_payroll_file(all_payroll_entries, payroll_file, logger)
        save_adjustments_files(pos_adj, neg_adj, adj_pos_file, adj_neg_file, logger)
        
        logger.info("Payroll processing completed successfully!")
        
    except Exception as e:
        logger.error(f"Error in payroll processing: {str(e)}")
        raise

def main():
    """Main program entry point."""
    try:
        # Display welcome message and instructions
        print("\n" + "="*80)
        print("\nWelcome to the Commission Processing System\n")
        print("Before proceeding, please ensure the following files are in your Downloads folder:")
        print("\n1. UUID file (format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.xlsx)")
        print("2. Jobs Report (format: Copy of Jobs Report for Performance -DE2_Dated *.xlsx)")
        print("3. Tech Department file (format: Technician Department_Dated *.xlsx)")
        print("4. Time Off file (name: Approved_Time_Off 2023.xlsx)")
        print("5. TGL file (format: TGLs Set _Dated *.xlsx)")
        print("\nThe program will:")
        print("1. Combine all worksheets into 'combined_data.xlsx'")
        print("2. Calculate commissions and create 'paystats.xlsx'")
        print("3. Process payroll and create 'payroll.xlsx'")
        print("4. Generate adjustment reports ('adjustments.xlsx' and 'adjustments_negative.xlsx')")
        print("\nAll output files will be created in your Downloads folder.")
        print("\n" + "="*80)

        while True:
            response = input("\nAre all required files ready in the Downloads folder? (Y/N): ").strip().upper()
            if response == 'Y':
                break
            elif response == 'N':
                print("\nPlease gather all required files and run the program again.")
                sys.exit(0)
            else:
                print("Invalid input. Please enter Y or N.")

        # Setup logging
        global logger
        logger = setup_logging('commission_processor')
        logger.info("Starting commission processing...")

        # Define base path
        base_path = os.path.expanduser(r'~\Downloads')

        # Get the base date and calculate the week range
        base_date = get_user_date()
        start_of_week = base_date - timedelta(days=base_date.weekday())
        end_of_week = start_of_week + timedelta(days=6)
        logger.info(f"Using week range: {start_of_week.date()} to {end_of_week.date()}")

        # Create output directory
        output_dir = create_output_directory(base_path, start_of_week, end_of_week, logger)

        # Define output file for combined data
        combined_file = os.path.join(output_dir, 'combined_data.xlsx')

        # Step 1: Combine workbooks
        logger.info("Combining workbooks...")
        combine_workbooks(base_path, combined_file)
        logger.info("Workbook combination completed!")

        # Step 2: Process calculations
        logger.info("\nStarting calculations...")
        process_calculations(base_path, output_dir, logger, start_of_week, end_of_week)

        # Step 3: Process payroll
        logger.info("\nStarting payroll processing...")
        process_payroll(base_path, output_dir, base_date, logger)

        logger.info("\nAll processing completed successfully!")

    except Exception as e:
        logger.error(f"Fatal error in main process: {str(e)}")
        sys.exit(1)



if __name__ == "__main__":
    main()