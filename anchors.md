# Function Anchors
[parse_filename_date_range](Payroll_Plus.py#L106):
Determine technician type based on business unit.

Returns:
    str: One of three values:
    - 'ADMIN' for administrative/sales staff (to be disregarded)
    - 'SERVICE' for service technicians (PCM, eligible for paystats and spiffs/TGL)
    - 'INSTALL' for installers (ICM, eligible for GP and spiffs/TGL)


[get_week_range](Payroll_Plus.py#L122):
Get the Monday-Sunday dates for the week containing the input date.
Any date entered will map to its corresponding week's Monday-Sunday range.

[validate_files_for_date](Payroll_Plus.py#L2162):
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


[validate_files_for_date_with_uuid](Payroll_Plus.py#L2212):
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


[get_validated_user_date](Payroll_Plus.py#L2282):
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

[determine_tech_type](Payroll_Plus.py#L387):
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

[get_main_department_code](Payroll_Plus.py#L411):
Get the main department code (20/30/40) from a subdepartment code.

Args:
    subdept_code (str): Two-digit subdepartment code (e.g., '21', '31', '42')
    
Returns:
    str: Main department code ('20', '30', or '40')

[get_full_department_code](Payroll_Plus.py#L428):
Get the full 7-digit department code based on main department.

Args:
    subdept_code (str): Two-digit subdepartment code
    
Returns:
    str: Seven-digit department code (e.g., '2000000', '3000000', '4000000')

[get_tech_home_department](Payroll_Plus.py#L445):
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

[get_service_department_code](Payroll_Plus.py#L462):
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

[consolidate_negative_spiffs](Payroll_Plus.py#L1500):
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
