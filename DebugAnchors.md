# Anchor comments x Payroll_Plus.py

[parse_filename_date_range](Payroll_Plus.py#L107):
Determine technician type based on business unit.

Returns:
    str: One of three values:
    - 'ADMIN' for administrative/sales staff (to be disregarded)
    - 'SERVICE' for service technicians (PCM, eligible for paystats and spiffs/TGL)
    - 'INSTALL' for installers (ICM, eligible for GP and spiffs/TGL)

[get_week_range](Payroll_Plus.py#L122):
Get the Monday-Sunday dates for the week containing the input date.
Any date entered will map to its corresponding week's Monday-Sunday range.

[validate_files_for_date](Payroll_Plus.py#L173):
This function validates the presence and correctness of files in a given directory for a specified date range:

- Determines the start and end dates of the week containing `user_date` using `get_week_range`.
- Formats these dates to create expected file naming patterns.
- Defines file patterns for three expected file types: "tech", "tgl", and "jobs".
- Searches the directory for files matching these patterns:
  - If no files match, it checks for files with similar patterns but different dates.
  - Adds errors for missing or mismatched files.
- Searches for UUID files (glob pattern for UUID format). Adds an error if none are found.
- Verifies the presence of a "Time Off" file and adds it to the found files dictionary if available.
- Returns:
  - A boolean indicating whether all validations passed.
  - A list of errors describing any issues.
  - A list of UUID files found (if any).
  - A dictionary of found files categorized by type.

[validate_files_for_date_with_uuid](Payroll_Plus.py#L229):
This function validates files in a directory for a specified date range, focusing on a specific UUID file:

- Calls `validate_files_for_date` to perform general file validation.
- Removes any UUID-related errors, as the validation will target a specific UUID file.
- Validates the selected UUID file by:
  - Analyzing its date range using `analyze_uuid_file_dates`.
  - Ensuring the file's data range covers the mid-week period (Wednesday to Friday) within the specified week.
- Adds an error if the UUID file's data range does not meet the mid-week coverage requirement.
- Returns:
  - A boolean indicating whether the validation passed.
  - A list of errors describing any issues found.

[get_validated_user_date](Payroll_Plus.py#L256):
This function validates user input for a date and ensures all required files, including a UUID file, are present and correct:

- Instantiates a `DateValidator` for file validation.
- Defines a nested `get_uuid_file_choice` function to handle UUID file selection:
  - If multiple UUID files are found, the user is prompted to confirm or select one.
  - Allows the user to input a specific UUID file name or exit the program.
- Prompts the user to enter a date in the format `mm/dd/yy` or type 'exit' to quit.
  - Converts the input to a `datetime` object and validates files using `validate_files_for_date`.
  - If multiple UUID files are present, the user selects one, and the selected file is further validated using `validate_files_for_date_with_uuid`.
  - If exactly one UUID file is present, it is used directly.
- If all required files are found with correct date ranges:
  - Returns a tuple containing the validated date, selected UUID file, and a dictionary of found files.
- If validation fails:
  - Displays errors and prompts the user to take corrective action, including entering a new date, downloading correct files, or exiting the program.

[determine_tech_type](Payroll_Plus.py#L381):
This function determines the type of technician (ADMIN, SERVICE, or INSTALL) based on the provided business unit:

- Handles empty or missing `business_unit` values:
  - Returns 'ADMIN' and logs a debug message if the business unit is empty or `NaN`.
- Converts the business unit to uppercase for consistent comparison and logs the input.
- Identifies technician type based on predefined patterns:
  1. **ADMIN**:
     - Matches patterns: 'ADMINISTRATIVE', '23 -', or 'SALES'.
     - Logs and returns 'ADMIN' if a match is found.
  2. **SERVICE**:
     - Matches pattern: 'SERVICE'.
     - Logs and returns 'SERVICE' if a match is found.
  3. **INSTALL**:
     - Default category for all other business units (e.g., INSTALL, EXCAVATION).
     - Logs and returns 'INSTALL' as the default technician type.
- Handles exceptions gracefully:
  - Logs an error message if an exception occurs.
  - Defaults to 'ADMIN' as a fallback to ensure safe operation.

[get_main_department_code](Payroll_Plus.py#L427):
Get the main department code (20/30/40) from a subdepartment code.

Args:
    subdept_code (str): Two-digit subdepartment code (e.g., '21', '31', '42')

Returns:
    str: Main department code ('20', '30', or '40')

[get_full_department_code](Payroll_Plus.py#L440):
Get the full 7-digit department code based on main department.

Args:
    subdept_code (str): Two-digit subdepartment code

Returns:
    str: Seven-digit department code (e.g., '2000000', '3000000', '4000000')

[get_tech_home_department](Payroll_Plus.py#L444):
This function retrieves a technician's home department code based on their business unit:

- **Input Example**: Converts a business unit like "PLUMBING SERVICE 30" to a 7-digit department code, e.g., "3000000".
- **Steps**:
  1. Extracts all numeric characters from the `tech_business_unit` using `filter` and `str.isdigit`.
  2. Checks the length of the extracted numbers:
     - If at least 2 digits are present:
       - Uses the first digit to map the department code:
         - '2' -> '2000000' (HVAC)
         - '3' -> '3000000' (Plumbing)
         - '4' -> '4000000' (Electric)
     - If fewer than 2 digits are found, returns '0000000' as a fallback.
- **Error Handling**:
  - Handles potential `IndexError` (e.g., when slicing the number string) or `AttributeError` (e.g., if the input is `None` or invalid).
  - Returns '0000000' if any error occurs, ensuring a default value.

[get_service_department_code](Payroll_Plus.py#L463):
This function retrieves a technician's home department code based on their business unit:

- **Input Example**: Converts a business unit like "PLUMBING SERVICE 30" to a 7-digit department code, e.g., "3000000".
- **Steps**:
  1. Extracts all numeric characters from the `tech_business_unit` using `filter` and `str.isdigit`.
  2. Checks the length of the extracted numbers:
     - If at least 2 digits are present:
       - Uses the first digit to map the department code:
         - '2' -> '2000000' (HVAC)
         - '3' -> '3000000' (Plumbing)
         - '4' -> '4000000' (Electric)
     - If fewer than 2 digits are found, returns '0000000' as a fallback.
- **Error Handling**:
  - Handles potential `IndexError` (e.g., when slicing the number string) or `AttributeError` (e.g., if the input is `None` or invalid).
  - Returns '0000000' if any error occurs, ensuring a default value.

[consolidate_negative_spiffs](Payroll_Plus.py#L467):
This function consolidates and categorizes negative and positive SPIFF (Sales Performance Incentive Fund) entries for technicians:

- **Inputs**:
  - `tech_data`: A DataFrame containing technician details (e.g., name, business unit, badge ID).
  - `spiffs_df`: A DataFrame containing SPIFF details (e.g., technician name, amount, memo).

- **Outputs**:
  - A tuple containing two lists of dictionaries:
    1. `negative_entries`: Consolidated negative SPIFF amounts per technician.
    2. `positive_entries`: Individual positive SPIFF entries with additional details.

- **Processing Steps**:
  1. Groups `spiffs_df` by the 'Technician' column to process SPIFFs for each technician individually.
  2. Retrieves the technician's home department code and badge ID from `tech_data`.
  3. **Negative SPIFFs**:
     - Iterates through each SPIFF entry for the technician.
     - Identifies and sums all negative amounts.
     - Appends a consolidated entry for the total negative SPIFFs, if any exist, to `negative_entries`.
  4. **Positive SPIFFs**:
     - Identifies all positive SPIFF entries.
     - Extracts the service department code from the memo and categorizes the entry.
     - Appends each positive SPIFF entry with its details to `positive_entries`.
  5. **Error Handling**:
     - Logs errors for invalid or missing data for a technician but continues processing others.
     - Handles potential conversion errors (e.g., malformed amounts) gracefully.

- **Error Logging**:
  - Logs any exceptions that occur during processing for debugging.

- **Key Helper Functions**:
  - `get_tech_home_department`: Retrieves the home department code based on the technician's business unit.
  - `extract_dept_code`: Extracts department codes from SPIFF memo strings.
  - `get_service_department_code`: Maps extracted department codes to service department codes.

[process_pcm_entry](Payroll_Plus.py#L523):
This function processes a Payroll Cost Management (PCM) entry for a technician:

- **Inputs**:
  - `tech_data`: A DataFrame containing technician details (e.g., name, badge ID).
  - `tech_name`: The name of the technician.
  - `subdept_code`: The sub-department code associated with the PCM entry.
  - `amount`: The amount for the PCM entry.
  - `date`: The date for the PCM entry.

- **Outputs**:
  - Returns a `PayrollEntry` object with the required details if the entry is successfully processed.
  - Returns `None` if the technician data is missing, invalid, or an error occurs.

- **Processing Steps**:
  1. Searches `tech_data` for the technician's information using their name.
  2. Checks if technician information is found:
     - If no matching data is found or the badge ID is missing (`NaN`), the function returns `None`.
  3. Converts the `subdept_code` to a full department code using `get_full_department_code`.
  4. Creates a `PayrollEntry` object with the following attributes:
     - `company_code`: Uses a predefined `COMPANY_CODE`.
     - `badge_id`: Extracted from `tech_data`.
     - `date`: Provided as input.
     - `amount`: Provided as input.
     - `pay_code`: Defaults to `'PCM'`.
     - `dept`: The full department code.
     - `location_id`: Uses a predefined `LOCATION_ID`.
  5. Handles exceptions:
     - Logs an error message and returns `None` if an exception occurs during processing.

[set_up_logging](Payroll_Plus.py#L553):
This function configures logging with both file and console handlers to log messages for the application:

- **Purpose**:
  - Logs messages to a file for detailed debugging.
  - Displays simplified messages in the console for user-friendly output.

- **Parameters**:
  - `name`: (Optional) The name of the logger, defaulting to `'commission_calculator'`.

- **Setup Steps**:
  1. Sets the logger level to `DEBUG` for comprehensive logging.
  2. Clears any existing handlers to avoid duplicate logs.
  3. Generates a log filename with a timestamp for uniqueness (`<name>_YYYYMMDD_HHMMSS.log`).
  4. **File Handler**:
     - Logs messages at the `DEBUG` level to the generated file.
     - Uses a detailed format: `'%(asctime)s - %(levelname)s - %(message)s'`.
  5. **Console Handler**:
     - Logs messages at the `INFO` level to the console (`sys.stdout`).
     - Uses a simpler format: `'%(message)s'`.
  6. **Console Filter** (specific to `'commission_calculator'`):
     - Filters out detailed technical messages from console output for improved readability.
     - Excludes messages containing certain keywords (e.g., 'Department Summary', 'Sales:', 'Commission:').
  7. Adds both handlers to the logger.

- **Returns**:
  - The configured logger instance.

- **Usage**:
  - This function is designed for use in applications like commission calculators where both detailed logs and user-friendly console output are required.

[create_output_directory](Payroll_Plus.py#L595):

This function creates an output directory named based on a specific week range:

- **Parameters**:
  - `base_path`: The base directory where the output folder will be created.
  - `start_of_week`: The start date of the week (datetime object).
  - `end_of_week`: The end date of the week (datetime object).
  - `logger`: A configured logging instance for logging messages.

- **Functionality**:
  1. **Folder Name**:
     - Generates a folder name in the format `Commission Output MM_DD_YY-MM_DD_YY`,
       where the dates represent the start and end of the week.
  2. **Directory Creation**:
     - Combines `base_path` and the generated folder name to form the full directory path.
     - Uses `os.makedirs` with `exist_ok=True` to create the directory if it doesn't already exist.
  3. **Logging**:
     - Logs an informational message upon successfully creating the directory.
     - Logs an error message and raises the exception if directory creation fails.

- **Returns**:
  - The full path of the created output directory.

- **Exceptions**:
  - Raises any exceptions that occur during directory creation after logging the error.

[autofit_columns](Payroll_Plus.py#L611):
This function adjusts the column widths in an Excel worksheet to automatically fit the content:

- **Parameters**:
  - `worksheet`: An Excel worksheet object from the `openpyxl` library.

- **Functionality**:
  1. Iterates through all columns in the worksheet.
  2. Determines the maximum length of the content in each column:
     - Converts each cell's value to a string and measures its length.
     - Skips cells with invalid or `None` values.
  3. Adjusts the width of each column:
     - Uses the column letter (e.g., 'A', 'B', etc.) to set the column width.
     - Adds padding of 2 characters to the maximum content length for better readability.

- **Notes**:
  - Utilizes `get_column_letter` from `openpyxl.utils` to convert column indices to Excel-style letters.
  - Ignores any exceptions when accessing or measuring cell content.

[get_valid_tgls](Payroll_Plus.py#L716):
This function retrieves valid TGLs (Technician Generated Leads) for a specified technician from an Excel file:

- **Parameters**:
  - `file_path`: The path to the Excel file containing TGL data.
  - `tech_name`: The name of the technician whose TGLs need to be retrieved.

- **Functionality**:
  1. Reads the Excel file from the specified `file_path`, accessing the `Sheet1_TGL` sheet.
  2. Filters TGL records:
     - Matches entries where the 'Lead Generated By' column equals `tech_name`.
     - Ensures the 'Status' is `'Completed'`.
  3. For each filtered TGL record:
     - Extracts the source and target business units.
     - Checks if the source and target units belong to the same department using `is_same_department`.
     - Appends valid TGLs to the result list with relevant details:
       - Job number, status, source unit, target unit, and creation date.
  4. Logs debug messages for:
     - Available columns in the data.
     - Each valid TGL found with its details.

- **Error Handling**:
  - Logs an error and returns an empty list if any exceptions occur during file reading or processing.

- **Returns**:
  - A list of dictionaries containing valid TGL details for the technician.

- **Notes**:
  - The `is_same_department` function is used to verify department alignment between source and target units.
  - Missing values for 'Job #' are replaced with `'N/A'` in the output.

[get_subdepartment_spiffs](Payroll_Plus.py#L750):
This function retrieves SPIFF amounts for a technician, categorized by subdepartment, for display purposes:

- **Parameters**:
  - `file_path`: Path to the Excel file containing SPIFF data.
  - `tech_name`: The name of the technician whose SPIFFs need to be retrieved.

- **Functionality**:
  1. Reads the Excel sheet named `'Direct Payroll Adjustments'` to load SPIFF data.
  2. Filters the data to include only entries for the specified technician.
  3. Initializes a dictionary with subdepartment codes as keys and their SPIFF totals set to 0.
  4. Iterates through the technician's SPIFF records:
     - Skips entries with missing or invalid `'Amount'` or `'Memo'` values.
     - Converts the `'Amount'` to a float, ignoring negative amounts.
     - Extracts the first two characters of the `'Memo'` field to determine the subdepartment code.
     - Adds the amount to the corresponding subdepartment total if the code is valid.
  5. Returns a dictionary containing SPIFF totals for each subdepartment.

- **Error Handling**:
  - Logs any exceptions encountered during file reading or processing.
  - Returns a dictionary with all subdepartment totals set to 0 if an error occurs.

- **Returns**:
  - A dictionary where keys are subdepartment codes and values are the total SPIFF amounts for each subdepartment.

- **Notes**:
  - Negative SPIFF amounts are excluded from the totals.
  - Assumes subdepartment codes are the first two characters of the `'Memo'` field.
  - Subdepartment codes are predefined and include:
    `'20', '21', '22', '24', '25', '27', '30', '31', '33', '34', '40', '41', '42'`.

[get_spiffs_total](Payroll_Plus.py#L795):
This function calculates the total SPIFFs (Sales Performance Incentive Funds) for a technician and categorizes them by department.

- **Parameters**:
  - `file_path`: The path to the Excel file containing SPIFF data.
  - `tech_name`: The name of the technician whose SPIFFs are being calculated.

- **Functionality**:
  1. Reads the Excel sheet `'Direct Payroll Adjustments'` to load SPIFF data.
  2. Filters the data to include only entries for the specified technician.
  3. Initializes a dictionary `department_spiffs` with department categories (`HVAC`, `Plumbing`, `Electric`) set to 0.
  4. Iterates through the technician's SPIFF records:
     - Skips entries with missing or invalid `'Amount'` or `'Memo'` fields.
     - Converts the `'Amount'` to a float and excludes negative or zero values.
     - Extracts the department number from the first two characters of the `'Memo'` field.
     - Validates the department number, ensuring it falls within specified ranges:
       - `20-29` for `HVAC`.
       - `30-39` for `Plumbing`.
       - `40-49` for `Electric`.
     - Adds the SPIFF amount to the appropriate department total.
  5. Logs debug messages for successfully processed SPIFFs and warnings for invalid entries.
  6. Computes the total SPIFF amount across all departments.

- **Error Handling**:
  - Logs detailed errors for invalid rows or formatting issues.
  - Raises exceptions for critical errors encountered during file reading or processing.

- **Returns**:
  - A tuple containing:
    1. `spiffs_total`: The total SPIFF amount for the technician.
    2. `department_spiffs`: A dictionary with SPIFF totals categorized by department.

- **Notes**:
  - Assumes department numbers are extracted from the first two characters of the `'Memo'` field.
  - Department ranges are predefined as:
    - `20-29` for `HVAC`.
    - `30-39` for `Plumbing`.
    - `40-49` for `Electric`.
  
[get_excused_hours](Payroll_Plus.py#L857):
This function calculates excused hours for technicians from a time-off sheet for a given week:

- **Parameters**:
  - `file_path`: The path to the Excel file containing time-off data.
  - `base_date`: A date within the target week (e.g., a Monday).
  - `sheet_name`: (Optional) The name of the Excel sheet to read. Defaults to `'2024'`.

- **Functionality**:
  1. **Week Range Calculation**:
     - Determines the start (`Monday`) and end (`Friday`) of the week containing `base_date`.
     - Formats the week range as `<Month> <Day><Suffix> - <Month> <Day><Suffix>` (e.g., "January 1st - January 5th").
  2. **Locate Week Columns**:
     - Scans the first two rows of the sheet for a header matching the formatted week range.
     - Identifies the starting column for the target week.
  3. **Data Processing**:
     - Iterates over rows starting from the third row (index `2`) to extract technician data.
     - Skips rows with missing or empty technician names.
     - Checks the cells in the target week’s columns for markers ('x', 'r', or 'v') indicating time off:
       - Each marker contributes 8 excused hours.
     - Sums up excused hours for each technician and stores the result in a dictionary.
  4. **Logging**:
     - Logs detailed debug messages for matching week ranges, detected markers, and processed hours.
     - Logs warnings for missing week ranges or errors during cell processing.

- **Error Handling**:
  - Catches and logs exceptions during file reading or processing.
  - Returns an empty dictionary if an error occurs.

- **Returns**:
  - A dictionary mapping technician names to their total excused hours for the target week.

- **Notes**:
  - Assumes valid markers for excused hours are `'x'`, `'r'`, and `'v'`.
  - Skips cells with invalid or missing data.
  - Handles missing week ranges gracefully by returning an empty dictionary.

[calculate_box_metrics](Payroll_Plus.py#L923):
This function calculates key metrics for a technician’s performance over a specified week:

- **Parameters**:
  - `data`: A DataFrame containing job-related information.
  - `tech_name`: The technician’s name.
  - `base_date`: A date within the target week.

- **Functionality**:
  1. **Week Range**:
     - Calculates the start (Monday) and end (Sunday) of the week containing `base_date`.
  2. **Relevant Jobs**:
     - Filters jobs where the technician is the `Primary Technician` or `Sold By`.
     - Logs the total number of relevant jobs found.
  3. **Box A (CJR - Completed Job Revenue)**:
     - Includes jobs where the technician is the `Primary Technician`, within the week range, and marked as an `Opportunity`.
     - Sums `Jobs Total Revenue` for these jobs.
     - Breaks down the revenue by subdepartment and logs details for each job.
  4. **Box B (TSIS - Technician Sold Installations and Services)**:
     - Includes jobs where the technician is the `Sold By` but not the `Primary Technician`, within the week range.
     - Sums `Jobs Total Revenue` for these jobs.
     - Breaks down the revenue by subdepartment and logs details for each job.
  5. **Box C (Total)**:
     - Computes as the sum of Box A and Box B.
  6. **Subdepartment Breakdown**:
     - Tracks revenue for completed, sold, and total jobs across predefined subdepartments.
  7. **Job Summary**:
     - Logs the number of jobs included and skipped, along with reasons for skipping.
     - Provides details of skipped jobs for transparency.

- **Returns**:
  - `box_a`: Total revenue for completed jobs (CJR).
  - `box_b`: Total revenue for sold jobs (TSIS).
  - `box_c`: Combined total revenue (CJR + TSIS).
  - `subdept_breakdown`: A dictionary with revenue breakdown by subdepartment for completed, sold, and total jobs.

- **Error Handling**:
  - Ensures `Invoice Date` is in datetime format.
  - Handles missing or invalid job fields gracefully with appropriate warnings.

- **Notes**:
  - Only jobs with valid opportunities and within the specified week are included in Box A.
  - Skipped jobs are logged to help understand why they were excluded.

[calculate_percentages](Payroll_Plus.py#L1049):
This function calculates the Service Completion Percentage (SCP) and Install Contribution Percentage (ICP) based on Box A (CJR) and Box C (Total) metrics:

- **Parameters**:
  - `box_a`: The completed job revenue (CJR).
  - `box_c`: The total revenue (CJR + TSIS).

- **Functionality**:
  1. **Zero Check**:
     - Returns `(0, 0)` if `box_c` is `0` to avoid division errors.
  2. **Raw Percentages**:
     - Calculates raw percentages for SCP (`box_a / box_c * 100`) and ICP (`100 - SCP`).
     - Ensures both percentages are non-negative using `max(0, value)`.
  3. **Normalization**:
     - Ensures that the sum of SCP and ICP equals 100 by proportionally adjusting the raw percentages.
  4. **Rounding**:
     - Rounds SCP and ICP to the nearest multiple of 10.
  5. **Final Adjustment**:
     - Ensures the sum of SCP and ICP equals exactly 100:
       - If SCP is greater, sets SCP to 100 and ICP to 0.
       - Otherwise, sets SCP to 0 and ICP to 100.

- **Returns**:
  - A tuple `(scp, icp)` containing the rounded SCP and ICP percentages as integers.

- **Notes**:
  - The function avoids invalid percentages (e.g., negatives or totals not summing to 100).
  - Designed to handle edge cases such as `box_c = 0` or rounding discrepancies.

[calculate_average_ticket_value](Payroll_Plus.py#L1078):
  This function calculates the average ticket value for a technician, based on total revenue divided by the number of opportunity jobs completed during a specific week.

- **Parameters**:
  - `data`: A DataFrame containing job-related information.
  - `tech_name`: The name of the technician.
  - `box_a`: The completed job revenue (CJR).
  - `box_b`: The technician-sold revenue (TSIS).
  - `base_date`: A date within the target week.
  - `logger`: A logging instance for debugging and progress tracking.

- **Functionality**:
  1. **Week Range**:
     - Calculates the start (Monday) and end (Sunday) of the week containing `base_date`.
  2. **Data Preparation**:
     - Ensures the `Invoice Date` column is in datetime format.
  3. **Job Filtering**:
     - Filters jobs completed by the technician (`Primary Technician`) within the week range.
     - Identifies opportunity jobs (`Opportunity` = `True`) from the completed jobs.
  4. **Revenue and Count Calculation**:
     - Calculates total revenue as `box_a + box_b`.
     - Tracks opportunity count and revenue for each department (HVAC, Plumbing, Electric).
  5. **Average Ticket Calculation**:
     - Computes the overall average ticket value as `total_revenue / opportunity_count`.
     - Calculates department-specific averages where applicable.
  6. **Logging**:
     - Logs detailed job breakdowns for both opportunity and non-opportunity jobs.
     - Provides a summary of department totals and overall metrics.
  7. **Handling Non-Opportunity Jobs**:
     - Logs details of non-opportunity jobs for transparency but excludes them from average calculations.

- **Returns**:
  - A dictionary containing the overall average ticket value and, optionally, department-specific averages.

- **Notes**:
  - Non-opportunity jobs are excluded from the average ticket calculation.
  - Departments are determined by mapping business unit codes using helper functions like `extract_department_number` and `get_department_from_number`.
  - Returns an overall average of `0.0` if no opportunity jobs are found.

[calculate_department_revenue](Payroll_Plus.py#L1231):
This function calculates department revenue for a technician during a specified week, categorizing revenue into completed, sales, and combined totals.

- **Parameters**:
  - `data`: A DataFrame containing job-related information.
  - `tech_name`: The name of the technician whose revenue is being calculated.
  - `base_date`: A date within the target week.

- **Functionality**:
  1. **Week Range**:
     - Calculates the start (Monday) and end (Sunday) of the week containing `base_date`.
  2. **Data Preparation**:
     - Ensures the `Invoice Date` column is in datetime format.
  3. **Completed Jobs**:
     - Filters jobs where the technician is the `Primary Technician`, within the week range, and marked as `Opportunity`.
     - Categorizes completed revenue into departments (`HVAC`, `Plumbing`, `Electric`) based on the `Business Unit`.
     - Updates both `completed` and `combined` revenue for the corresponding department.
  4. **Sold Jobs**:
     - Filters jobs where the technician is the `Sold By`, within the week range, and not the `Primary Technician`.
     - Categorizes sales revenue into departments based on the `Business Unit`.
     - Updates both `sales` and `combined` revenue for the corresponding department.
  5. **Unknown Department**:
     - Any jobs with unrecognized `Business Unit` values are categorized under `'Unknown'`.

- **Returns**:
  - A dictionary with three keys: `completed`, `sales`, and `combined`. Each contains sub-dictionaries mapping departments (`HVAC`, `Plumbing`, `Electric`, `Unknown`) to their respective revenue totals.

- **Notes**:
  - Uses helper functions like `extract_department_number` and `get_department_from_number` to map `Business Unit` values to departments.
  - Assumes revenue from `'Jobs Total Revenue'` is used for calculations.
  - Ensures unrecognized `Business Unit` values do not disrupt calculations by grouping them under `'Unknown'`.

[get_commission_rate](Payroll_Plus.py#L1296):
This function calculates the commission rate for a technician based on their revenue, department, and other factors, applying various adjustments to the thresholds.

- **Parameters**:
  - `total_revenue`: Total revenue generated by the technician.
  - `flipped_percent`: The Install Contribution Percentage (ICP), rounded to the nearest 10%.
  - `department`: The department (`Electric`, `Plumbing`, or `HVAC`) for which the thresholds are calculated.
  - `excused_hours`: Number of excused hours for the technician (affects thresholds).
  - `tgl_reduction`: Total TGL credit amount applied as a reduction to thresholds.
  - `avg_ticket_value`: The average ticket value for calculating TGL counts.

- **Functionality**:
  1. **Flipped Percent Adjustment**:
     - Rounds `flipped_percent` to the nearest 10% and limits it to a range of 0% to 100%.
  2. **Base Thresholds**:
     - Selects the appropriate threshold table (`HVAC` or `Plumbing/Electrical`) based on the department and `flipped_percent`.
  3. **Time Off Adjustment**:
     - Reduces thresholds based on the number of excused hours (up to 5 days) with a reduction factor of 20% per day off.
  4. **TGL Reduction**:
     - Applies a TGL reduction based on the number of valid TGLs (`tgl_reduction / avg_ticket_value`) and reduces each threshold accordingly.
  5. **Determine Commission Rate**:
     - Compares `total_revenue` against adjusted thresholds to determine the highest tier met (2%, 3%, 4%, or 5%).
     - If no threshold is met, the commission rate is 0%.
  6. **Logging**:
     - Logs detailed calculations for debugging, including adjustments for time off, TGL reductions, and final threshold values.

- **Returns**:
  - `rate`: The calculated commission rate as a decimal (e.g., 0.05 for 5%).
  - `adjusted_thresholds`: The thresholds after all reductions.
  - `tier_thresholds`: The original thresholds before adjustments.

- **Notes**:
  - Ensures all adjusted thresholds and reductions are non-negative.
  - TGL counts are derived using the average ticket value, avoiding division errors when `avg_ticket_value` is 0.
  - The function is robust to handle edge cases such as minimal revenue or extensive time off.

[read_tech_department_data](Payroll_Plus.py#L1491):
This function reads and processes technician department data from an Excel file.

- **Parameters**:
  - `file_path`: The path to the Excel file containing technician data.
  - `logger`: A logging instance for logging progress and errors.

- **Functionality**:
  1. **File Reading**:
     - Reads the sheet `'Sheet1_Tech'` from the Excel file.
     - Ensures the `Payroll ID` column is read as a string.
  2. **Data Cleaning**:
     - Removes rows where the `Name` field contains numeric values.
     - Excludes technicians listed in the `EXCLUDED_TECHS` constant.
  3. **Badge ID Formatting**:
     - Applies the `format_badge_id` function to the `Payroll ID` column to format badge IDs correctly.
  4. **Logging**:
     - Logs the number of technician records successfully loaded.
     - Logs errors encountered during the reading or processing of data.

- **Returns**:
  - A cleaned and processed DataFrame containing technician department data.

- **Raises**:
  - Re-raises exceptions encountered during file reading or processing after logging the error.

- **Notes**:
  - Assumes `EXCLUDED_TECHS` and `format_badge_id` are defined elsewhere in the program.
  - Ensures robustness by handling unexpected errors gracefully and logging them for debugging.

[sum_spiffs_for_dept](Payroll_Plus.py#L1526):
This function calculates the total SPIFFs (Sales Performance Incentive Funds) for a specific technician and department code.

- **Parameters**:
  - `spiffs_df`: A DataFrame containing SPIFF data, including columns like `Technician`, `Memo`, and `Amount`.
  - `tech_name`: The name of the technician whose SPIFFs are being summed.
  - `dept_code`: The department code to filter SPIFFs (e.g., "20" for HVAC).

- **Functionality**:
  1. Filters the SPIFF DataFrame to include only:
     - Rows where the `Technician` matches `tech_name`.
     - Rows where the `Memo` field contains the specified `dept_code` (extracted using `extract_dept_code`).
     - Rows where the `Amount` is a positive number.
  2. If no matching rows are found, returns `0.0`.
  3. Converts valid SPIFF amounts to floats and sums them.

- **Returns**:
  - The total SPIFF amount for the specified technician and department as a float.

- **Notes**:
  - Assumes that `extract_dept_code` is a function that extracts department codes from the `Memo` field.
  - Handles missing or invalid `Amount` values gracefully by treating them as `0`.
  - Filters out SPIFF entries with negative or zero amounts.

[process_paystats](Payroll_Plus.py#L1546):
This function processes payroll entries for service technicians from a paystats file and related data sources.

- **Parameters**:
  - `output_dir`: Path to the directory containing output files (e.g., combined data).
  - `paystats_file`: Path to the paystats Excel file.
  - `tech_data`: A DataFrame containing technician details, including names and business units.
  - `base_date`: A date within the target week for which payroll entries are being processed.
  - `logger`: A logging instance for progress and error tracking.

- **Functionality**:
  1. **Initialize Date Range**:
     - Calculates the start (Monday) and end (Sunday) of the week containing `base_date`.
     - Formats the week-end date for use in payroll entries.
  2. **Load Data**:
     - Reads technician data from `paystats_file` and SPIFF adjustments from `combined_data.xlsx`.
     - Filters out non-service technicians and excluded technicians (`EXCLUDED_TECHS`).
  3. **Iterate Over Technicians**:
     - For each service technician:
       - Retrieves the commission rate and skips processing if it's 0%.
       - Finds the technician's badge ID using `Badge ID` or `Payroll ID`.
       - Skips technicians without a valid ID.
  4. **Process Department Entries**:
     - Iterates through predefined department codes:
       - Parses revenue, sales, and total amounts for the department.
       - Retrieves SPIFF adjustments for the department and subtracts positive SPIFFs from the total.
       - Applies the technician's commission rate to the adjusted amount.
       - Skips entries where the adjusted amount or final commission is 0 or less.
       - Creates a `PayrollEntry` for the adjusted amount and adds it to the results.
  5. **Error Handling**:
     - Logs warnings for individual processing errors (e.g., missing data or invalid entries).
     - Raises exceptions for critical issues during file reading or data processing.

- **Returns**:
  - A list of `PayrollEntry` objects representing payroll adjustments for the specified week.

- **Notes**:
  - Adjusted amounts consider positive SPIFFs but exclude negatives, which are handled separately.
  - The `get_service_department_code` function maps department codes to their main service codes.
  - Handles missing or invalid data gracefully, logging issues for debugging while continuing processing.

[process_gp_entries](Payroll_Plus.py#L1677):
This function processes payroll entries for service technicians from a paystats file and related data sources.

- **Parameters**:
  - `output_dir`: Path to the directory containing output files (e.g., combined data).
  - `paystats_file`: Path to the paystats Excel file.
  - `tech_data`: A DataFrame containing technician details, including names and business units.
  - `base_date`: A date within the target week for which payroll entries are being processed.
  - `logger`: A logging instance for progress and error tracking.

- **Functionality**:
  1. **Initialize Date Range**:
     - Calculates the start (Monday) and end (Sunday) of the week containing `base_date`.
     - Formats the week-end date for use in payroll entries.
  2. **Load Data**:
     - Reads technician data from `paystats_file` and SPIFF adjustments from `combined_data.xlsx`.
     - Filters out non-service technicians and excluded technicians (`EXCLUDED_TECHS`).
  3. **Iterate Over Technicians**:
     - For each service technician:
       - Retrieves the commission rate and skips processing if it's 0%.
       - Finds the technician's badge ID using `Badge ID` or `Payroll ID`.
       - Skips technicians without a valid ID.
  4. **Process Department Entries**:
     - Iterates through predefined department codes:
       - Parses revenue, sales, and total amounts for the department.
       - Retrieves SPIFF adjustments for the department and subtracts positive SPIFFs from the total.
       - Applies the technician's commission rate to the adjusted amount.
       - Skips entries where the adjusted amount or final commission is 0 or less.
       - Creates a `PayrollEntry` for the adjusted amount and adds it to the results.
  5. **Error Handling**:
     - Logs warnings for individual processing errors (e.g., missing data or invalid entries).
     - Raises exceptions for critical issues during file reading or data processing.

- **Returns**:
  - A list of `PayrollEntry` objects representing payroll adjustments for the specified week.

- **Notes**:
  - Adjusted amounts consider positive SPIFFs but exclude negatives, which are handled separately.
  - The `get_service_department_code` function maps department codes to their main service codes.
  - Handles missing or invalid data gracefully, logging issues for debugging while continuing processing.

[match_department_spiffs](Payroll_Plus.py#L1815):
This function matches positive and negative SPIFFs (Sales Performance Incentive Funds) within departments, processes TGL (Technician Generated Leads) entries, and tracks calculations for reporting.

- **Parameters**:
  - `adj_df`: A DataFrame containing adjustment data, including SPIFF amounts, technicians, and memos.
  - `logger`: A logging instance for tracking progress and debugging.

- **Functionality**:
  1. **Filter Data**:
     - Excludes technicians listed in `EXCLUDED_TECHS` from processing.
  2. **Process TGL Entries**:
     - Identifies rows where the `Memo` contains "tgl".
     - Extracts the technician, department code, memo, and amount for each TGL entry.
     - Tracks these entries in a separate list (`tgl_entries`).
  3. **Match Positive and Negative SPIFFs**:
     - Groups the adjustment data by technician.
     - For each technician:
       - Processes SPIFFs by department, summing positive and negative amounts separately.
       - Calculates the net amount as the sum of positives and negatives for each department.
       - Logs the totals and calculations for each department.
     - Tracks matched entries (`matched_entries`) with detailed calculations and statuses.
     - Identifies departments with significant negative balances and tracks them separately (`unmatched_negatives`).
  4. **Detailed Logging**:
     - Logs department-level calculations, including total positives, negatives, and net amounts.
     - Tracks entries with negative balances for additional reporting or follow-up.

- **Returns**:
  - A tuple containing three DataFrames:
    1. `tgl_entries`: Contains TGL-related entries.
    2. `matched_entries`: Contains matched SPIFF entries with detailed calculations and net amounts.
    3. `unmatched_negatives`: Contains entries where net balances are negative.

- **Notes**:
  - Uses helper functions (e.g., `process_department_entries`) to process department-level SPIFF data.
  - Ensures all calculations are logged for traceability.
  - Tracks unmatched negative balances for additional reporting or reconciliation.

[process_adjustments](Payroll_Plus.py#L1884):
This function processes payroll adjustments, separating TGL (Technician Generated Leads), positive SPIFFs, and consolidated negative SPIFFs for technicians.

- **Parameters**:
  - `combined_file`: Path to the Excel file containing adjustment and technician data.
  - `logger`: A logging instance for progress and error tracking.

- **Functionality**:
  1. **Load Data**:
     - Reads the `Direct Payroll Adjustments` and `Sheet1_Tech` sheets from the combined Excel file.
     - Filters out excluded technicians using `EXCLUDED_TECHS`.
  2. **Create Technician Lookup**:
     - Builds a dictionary mapping technician names to their badge ID, business unit, and home department.
  3. **Filter Adjustments**:
     - Excludes rows with invalid or excluded technicians and rows containing "Totals."
  4. **Process Adjustments**:
     - Separates TGL and SPIFF entries:
       - **TGL Entries**:
         - Identifies rows where the memo contains "tgl."
         - Extracts subdepartment codes and maps them to full department codes.
         - Appends TGL details to `tgl_entries`.
       - **Positive SPIFFs**:
         - Processes positive amounts, maps them to department codes, and appends them to `spiff_entries`.
       - **Negative SPIFFs**:
         - Aggregates negative amounts by technician and stores them in `tech_negatives`.
  5. **Consolidate Negative SPIFFs**:
     - Creates entries for each technician with a net negative SPIFF total, mapped to their home department.
  6. **Logging**:
     - Logs detailed counts of processed entries for TGLs, positive SPIFFs, and negative SPIFFs.

- **Returns**:
  - `tgl_df`: DataFrame of TGL entries.
  - `spiff_df`: DataFrame of positive SPIFF entries.
  - `spiff_df` (duplicate, potentially intended for further processing).
  - `neg_df`: DataFrame of consolidated negative SPIFF entries.

- **Notes**:
  - Uses helper functions like `get_tech_home_department` and `DEPARTMENT_CODES` for data mapping.
  - Negative SPIFFs are consolidated for each technician and included as separate entries.
  - Handles missing or invalid data gracefully, logging warnings for problematic rows.

[save_payroll_file](Payroll_Plus.py#L2003):
This function saves payroll entries to an Excel file with specific formatting, validation, and consolidation.

- **Parameters**:
  - `entries`: A list of `PayrollEntry` objects containing payroll data.
  - `output_file`: The path to the output Excel file.
  - `logger`: A logging instance for progress and error tracking.

- **Functionality**:
  1. **Convert to DataFrame**:
     - Converts `PayrollEntry` objects to a DataFrame with columns:
       - `Company Code`, `Badge ID`, `Date`, `Amount`, `Pay Code`, `Dept`, `Location ID`.
  2. **Consolidate Duplicates**:
     - Groups entries by `Badge ID`, `Dept`, `Pay Code`, `Company Code`, `Date`, and `Location ID`.
     - Sums `Amount` for duplicate entries and warns about consolidation.
  3. **Validation**:
     - Validates:
       - Positive `Amount` values.
       - Valid `Pay Code` (must be one of `PCM`, `ICM`, or `SPF`).
       - Valid `Dept` codes (matches codes in `DEPARTMENT_CODES`).
     - Logs warnings or errors for invalid rows.
  4. **Write to Excel**:
     - Saves the validated DataFrame to an Excel file.
     - Formats columns:
       - Adjusts column widths based on content.
       - Boldens headers.
       - Formats the `Amount` column with right alignment and a currency format (`#,##0.00`).
       - Center-aligns other columns.
  5. **Logging**:
     - Logs the total number of saved entries.
     - Provides a breakdown of entries by `Pay Code` (`PCM`, `ICM`, `SPF`).

- **Returns**:
  - Saves the file to `output_file`. Does not return a value.

- **Notes**:
  - Handles missing or invalid data gracefully by skipping problematic rows and logging warnings.
  - Consolidates entries with identical keys to prevent duplication in payroll records.
  - Uses `openpyxl` for Excel writing and formatting.
  - Ensures all entries are validated before saving.

  [save_adjustment_files](Payroll_Plus.py#L2094):
  This function processes and saves adjustment files, including payroll entries, positive and negative SPIFF adjustments, and TGL entries.

- **Parameters**:
  - `tgl_df`: DataFrame containing TGL (Technician Generated Leads) entries.
  - `matched_df`: DataFrame containing matched positive SPIFF entries.
  - `pos_df`: DataFrame for positive SPIFF references.
  - `neg_df`: DataFrame for negative SPIFF references.
  - `matched_file`: Path to save matched payroll entries.
  - `pos_file`: Path to save positive SPIFF reference entries.
  - `neg_file`: Path to save negative SPIFF reference entries.
  - `tech_data`: DataFrame containing technician details.
  - `base_date`: A date within the target week.
  - `logger`: A logging instance for tracking progress and errors.

- **Functionality**:
  1. **Calculate Week Dates**:
     - Determines the start (Monday) and end (Sunday) of the week based on `base_date`.
  2. **Process Payroll Entries**:
     - Reads the payroll file and filters for PCM entries.
     - Updates PCM entries with negative SPIFFs, ensuring proper department and badge ID matching.
     - Tracks unprocessed negative adjustments for reporting.
  3. **Process SPIFFs**:
     - Consolidates TGL and positive SPIFF amounts by badge ID and department.
     - Creates payroll entries for SPIFFs.
  4. **Save Updated Payroll Entries**:
     - Updates the payroll file with modified PCM entries and saves new SPIFF entries.
  5. **Save Adjustment Files**:
     - Saves matched payroll entries, positive reference entries, and negative adjustments as Excel files with proper formatting.
  6. **Logging**:
     - Logs the counts of saved entries for payroll, positive SPIFFs, and negative adjustments.

- **Error Handling**:
  - Logs errors and warnings during processing.
  - Catches exceptions to ensure processing continues for valid entries.

- **Excel Formatting**:
  - Adjusts column widths based on content.
  - Formats headers with bold text.
  - Formats the `Amount` column with currency formatting.

- **Returns**:
  - Saves files to the specified paths. Does not return a value.

- **Notes**:
  - Handles empty DataFrames gracefully, ensuring valid output files.
  - Updates payroll files to reflect consolidated negative SPIFFs against PCM entries.

  [validate_required_files](Payroll_Plus.py#L2343):
  This function ensures that all required files are present in the specified directory before processing.

- **Parameters**:
  - `directory`: The path to the directory containing the required files.

- **Functionality**:
  1. **UUID File Check**:
     - Uses `glob.glob` to search for a file matching the pattern of a UUID-like name (e.g., "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.xlsx").
     - Raises a `FileNotFoundError` if no such file is found.
  2. **Jobs Report File Check**:
     - Searches for a file with the pattern `Copy of Jobs Report for Performance -DE2_Dated *.xlsx`.
     - Raises a `FileNotFoundError` if no matching file is found.
  3. **Tech Department File Check**:
     - Searches for a file with the pattern `Technician Department_Dated *.xlsx`.
     - Raises a `FileNotFoundError` if no matching file is found.
  4. **Time Off File Check**:
     - Validates the presence of a file named `Approved_Time_Off 2023.xlsx`.
     - Raises a `FileNotFoundError` if the file is not found.
  5. **Exit on Missing File**:
     - If any required file is missing, prints an error message and exits the program with status code `1`.

- **Returns**:
  - `True` if all required files are present.

- **Error Handling**:
  - Catches `FileNotFoundError` for missing files.
  - Exits the program with an error message if any required file is not found.

- **Usage**:
  - Call this function before starting the main process to ensure all necessary files are in place.

- **Notes**:
  - This function assumes the program cannot proceed without all required files.
  - Wildcard matching is used for filenames that include dynamic date components.
  - The program exits if validation fails, ensuring no further processing occurs with missing inputs.

[find_latest_files](Payroll_Plus.py#L2370):
This function finds the latest version of each required file in a specified directory.

- **Parameters**:
  - `directory`: The directory path where the required files are located.

- **Functionality**:
  1. **UUID File**:
     - Searches for a file matching the pattern of a UUID-like name (`xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.xlsx`).
     - Raises a `FileNotFoundError` if no such file is found.
  2. **Jobs Report File**:
     - Searches for files matching the pattern `Copy of Jobs Report for Performance -DE2_Dated *.xlsx`.
     - Selects the latest file by timestamp using `max`.
     - Raises a `FileNotFoundError` if no matching files are found.
  3. **Tech Department File**:
     - Searches for files matching the pattern `Technician Department_Dated *.xlsx`.
     - Selects the latest file by timestamp using `max`.
     - Raises a `FileNotFoundError` if no matching files are found.
  4. **Time Off File**:
     - Validates the existence of a file named `Approved_Time_Off 2023.xlsx`.
     - Raises a `FileNotFoundError` if the file is not found.
  5. **TGL File (Optional)**:
     - Searches for files matching the pattern `TGLs Set _Dated *.xlsx`.
     - Selects the latest file by timestamp using `max` if such files exist.
     - Returns `None` if no matching files are found.

- **Returns**:
  - A dictionary containing the paths to the latest version of each required file:
    - `'uuid'`: Path to the UUID file.
    - `'jobs'`: Path to the latest Jobs Report file.
    - `'tech'`: Path to the latest Tech Department file.
    - `'time_off'`: Path to the Time Off file.
    - `'tgl'`: Path to the latest TGL file, or `None` if no TGL files exist.

- **Error Handling**:
  - Raises `FileNotFoundError` if any mandatory file is missing.
  - Handles optional files (like TGLs) by returning `None` if no matching files are found.

- **Constants**:
  - `NAME_COLUMNS_MAP`: A dictionary mapping sheet names to the expected column names for processing specific data files.

- **Notes**:
  - Ensures robustness by selecting the latest version of files with dynamic date components.
  - Optional files are handled gracefully without disrupting the function's flow.
  - Can be used in conjunction with validation functions to ensure all necessary files are present and up-to-date.

[combine_workbooks](Payroll_Plus.py#L2411):
This function combines multiple Excel workbooks into a single consolidated file, cleaning name columns and formatting sheets as needed.

- **Parameters**:
  - `directory`: The directory containing the source files.
  - `output_file`: The path where the combined workbook will be saved.
  - `files`: A dictionary mapping required file keys (`uuid`, `jobs`, `tech`, `tgl`) to their respective file paths.

- **Functionality**:
  1. **Initialize with UUID File**:
     - Copies the UUID file to the output path as the base for the combined workbook.
     - Cleans the `Technician` column in the UUID file by trimming whitespace and logs changes.
     - Formats columns using `autofit_columns`.
  2. **Process Additional Workbooks**:
     - Iterates over a configuration of source files (`jobs`, `tech`, and optionally `tgl`).
     - For each workbook:
       - Loads the source workbook and selects the relevant source sheet.
       - Creates or replaces the corresponding target sheet in the combined workbook.
       - Cleans name columns specified in the configuration by trimming whitespace.
       - Copies all data from the source sheet to the target sheet.
       - Formats columns in the target sheet.
  3. **Save the Combined Workbook**:
     - Saves the consolidated workbook to `output_file`.

- **Helper Function**:
  - `clean_name_column(worksheet, col_idx)`:
    - Trims whitespace from cells in a specified column.
    - Logs changes when names are cleaned.

- **Configurations**:
  - Processes the following files:
    - `jobs` → Copies data from `Sheet1` to `Sheet1` with cleaning in `Sold By`, `Primary Technician`, and `Technician` columns.
    - `tech` → Copies data from `Sheet1` to `Sheet1_Tech` with cleaning in the `Name` column.
    - `tgl` (optional) → Copies data from `Sheet1` to `Sheet1_TGL` with cleaning in the `Lead Generated By` column.

- **Notes**:
  - Assumes all required files (`uuid`, `jobs`, `tech`) are present, with `tgl` being optional.
  - Uses `shutil.copy2` to preserve file metadata when copying the UUID file.
  - Trims whitespace from name columns to ensure consistency across sheets.
  - Relies on `autofit_columns` to adjust column widths for readability.

[process_department_entries](Payroll_Plus.py#L2495):
This function processes all SPIFF entries for a technician, grouping them by department and summing positive and negative amounts separately.

- **Parameters**:
  - `tech_group`: A DataFrame containing adjustment entries for a single technician, including columns for `Amount` and `Memo`.

- **Functionality**:
  1. **Initialize Totals**:
     - Creates a `defaultdict` where each department has a dictionary with `positives` and `negatives` initialized to 0.0.
  2. **Iterate Through Entries**:
     - For each entry:
       - Converts the `Amount` column to a float after stripping currency formatting.
       - Extracts the department code from the first two characters of the `Memo` column.
       - Skips entries where:
         - The `Memo` contains "tgl" (indicating a Technician Generated Lead).
         - The `Memo` does not contain a valid numeric department code.
       - Accumulates positive amounts into the `positives` key.
       - Accumulates negative amounts (already negative) into the `negatives` key.
  3. **Handle Errors**:
     - Ignores rows with invalid data in the `Amount` or `Memo` columns, logging no errors but silently continuing.

- **Returns**:
  - A dictionary (`dept_totals`) where keys are department codes (as strings), and values are dictionaries containing:
    - `'positives'`: The sum of all positive amounts for the department.
    - `'negatives'`: The sum of all negative amounts for the department.

- **Notes**:
  - Designed to handle numeric and string formats for the `Amount` column, including currency symbols and commas.
  - Ignores invalid department codes or malformed `Memo` entries.
  - Can be integrated into higher-level processes to calculate net totals or generate reports by department.

[process_calculations](Payroll_Plus.py#L2524):
This function processes all calculations for technician revenue, commissions, and adjustments, then generates the required output files.

- **Parameters**:
  - `base_path`: Path to the directory containing source files, such as the time-off file.
  - `output_dir`: Path to the directory where output files will be saved.
  - `logger`: A logging instance for tracking progress and errors.
  - `start_of_week`: The start date (Monday) of the target week.
  - `end_of_week`: The end date (Sunday) of the target week.

- **Functionality**:
  1. **Initialize File Paths**:
     - Constructs paths for the combined data file (`combined_data.xlsx`) and the paystats output file (`paystats.xlsx`).
  2. **Read Data**:
     - Reads technician and job data from the combined file (`Sheet1`).
     - Filters technician data to include only service technicians using the `determine_tech_type` function.
  3. **Calculate Excused Hours**:
     - Reads excused hours for technicians from the original time-off file (`Approved_Time_Off 2023.xlsx`).
     - Maps technicians to their respective excused hours in a dictionary.
  4. **Process Commission Calculations**:
     - Calls `process_commission_calculations` to calculate revenue, commissions, and adjustments for service technicians.
     - Utilizes data, service technician data, combined file, week start date, and excused hours.
  5. **Save Results**:
     - Saves the calculated results to the paystats file (`paystats.xlsx`) in a sheet named `Technician Revenue Totals`.
     - Formats the columns for readability using `autofit_columns`.
  6. **Logging**:
     - Logs the successful completion of commission calculations.
  7. **Error Handling**:
     - Catches exceptions, logs the error message, and raises the exception to ensure visibility.

- **Notes**:
  - Ensures only service technicians are processed for commissions and adjustments.
  - Handles time-off data from the original file to maintain accurate calculations.
  - Uses `openpyxl` to save and format the paystats file.

[process_payroll](Payroll_Plus.py#L2562):
This function processes payroll and adjustments, separating the handling of service technicians and installers, and generates the required output files.

- **Parameters**:
  - `base_path`: Path to the directory containing source files.
  - `output_dir`: Path to the directory where output files will be saved.
  - `base_date`: A date within the target payroll week.
  - `logger`: A logging instance for tracking progress and errors.
  - `tech_data`: A DataFrame containing technician information, including business units.

- **Functionality**:
  1. **Define File Paths**:
     - Constructs paths for output files, including combined data, paystats, payroll, and adjustment files.
  2. **Split Technicians by Type**:
     - Filters `tech_data` into service technicians and installers using the `determine_tech_type` function.
     - Logs the count of service technicians and installers being processed.
  3. **Process Payroll Entries**:
     - Calls `process_paystats` to calculate and process service technician commissions, generating payroll entries.
     - Calls `process_gp_entries` to handle installer GP entries separately.
     - Combines payroll entries from both groups into a single list.
  4. **Process Adjustments**:
     - Calls `process_adjustments` to process TGLs and SPIFF adjustments for eligible technicians (both service and installers).
     - Generates DataFrames for TGLs, matched entries, positive adjustments, and negative adjustments.
  5. **Save Output Files**:
     - Saves payroll entries to an Excel file using `save_payroll_file`.
     - Saves adjustment data (TGLs, positive, and negative adjustments) to separate Excel files using `save_adjustment_files`.
  6. **Logging**:
     - Logs the number of entries processed for service technician commissions, installer GP, and total payroll entries.
     - Confirms successful completion of payroll processing.
  7. **Error Handling**:
     - Catches exceptions, logs the error message, and raises the exception for visibility.

- **Generated Files**:
  - `combined_data.xlsx`: Consolidated data for processing.
  - `paystats.xlsx`: Technician revenue and commission calculations.
  - `payroll.xlsx`: Final payroll entries for submission.
  - `Spiffs.xlsx`: Consolidated SPIFFs and TGL adjustments.
  - `positive_adjustments.xlsx`: Positive SPIFF adjustments.
  - `negative_adjustments.xlsx`: Negative SPIFF adjustments.

- **Notes**:
  - Ensures service technicians and installers are processed separately, given their distinct payroll calculations.
  - Consolidates all payroll entries and adjustments into well-structured output files.
  - Logs detailed progress and results for transparency and debugging.

[main](Payroll_Plus.py#L2613):
This is the main entry point for the Commission Processing System, which coordinates the processing of payroll and adjustments for both service technicians and installers.

- **Functionality**:
  1. **User Interaction**:
     - Displays a welcome message with a checklist of required files.
     - Prompts the user to confirm the presence of all required files in the Downloads folder.
     - Exits if the files are not ready.
  2. **Initialization**:
     - Sets `base_path` to the user's Downloads directory.
     - Calls `get_validated_user_date` to select a valid date and identify required files.
     - Calculates the week range based on the selected date.
     - Configures logging with `setup_logging`.
     - Creates an output directory for processed files.
  3. **Workbook Combination**:
     - Combines data from multiple workbooks into a single file (`combined_data.xlsx`) using `combine_workbooks`.
  4. **Technician Categorization**:
     - Reads technician data with `read_tech_department_data`.
     - Splits technicians into service technicians and installers based on their business units.
  5. **Service Technician Calculations**:
     - Calls `process_calculations` to compute service technician commissions and generate paystats.
  6. **Payroll and Adjustments**:
     - Processes service technician commissions using `process_paystats`.
     - Processes installer GP entries using `process_gp_entries`.
     - Handles TGLs and SPIFFs for all eligible technicians using `process_adjustments`.
  7. **File Saving**:
     - Saves payroll entries to `payroll.xlsx` using `save_payroll_file`.
     - Saves adjustment data to separate files for TGLs, positive adjustments, and negative adjustments using `save_adjustment_files`.
  8. **Logging**:
     - Logs the number of entries processed for service technicians, installers, and total payroll.
     - Logs successful completion of all steps.
  9. **Error Handling**:
     - Logs fatal errors and exits with status code `1` if an exception occurs.

- **Required Files**:
  - UUID file.
  - Jobs Report file.
  - Tech Department file.
  - Time Off file.
  - TGL file.

- **Generated Files**:
  - `combined_data.xlsx`: Consolidated data for processing.
  - `paystats.xlsx`: Technician revenue and commission calculations.
  - `payroll.xlsx`: Final payroll entries for submission.
  - `Spiffs.xlsx`: Consolidated SPIFFs and TGL adjustments.
  - `positive_adjustments.xlsx`: Positive SPIFF adjustments.
  - `negative_adjustments.xlsx`: Negative SPIFF adjustments.

- **Exit Conditions**:
  - Exits if required files are missing.
  - Exits with a logged error message if a fatal error occurs during processing.

- **Notes**:
  - Handles both service technicians and installers with distinct processing pipelines.
  - Relies on user confirmation before starting the process to ensure all required files are ready.
  - Modular design allows each processing step to be reused or extended.
