# How to Use the Payroll Processing System

## Pre-Processing Steps

### 1. File Preparation

First, ensure all required files are downloaded to your Downloads folder:

a) **UUID File**

- From Sage: Export Direct Payroll Adjustments report
- Filename will be a UUID (e.g., `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.xlsx`)
- Contains spiffs and TGL information

b) **Jobs Report**

- From ConnectWise: Export Jobs Report for Performance
- Filename format: `Copy of Jobs Report for Performance -DE2_Dated MM_DD_YY - MM_DD_YY.xlsx`
- Contains job completion and revenue data

c) **Tech Department File**

- Export technician department assignments
- Filename format: `Technician Department_Dated MM_DD_YY - MM_DD_YY.xlsx`
- Contains technician information and department assignments

d) **Time Off File**

- Named exactly: `Approved_Time_Off 2023.xlsx`
- Contains excused hours information

e) **TGL File**

- Export TGL set data
- Filename format: `TGLs Set _Dated MM_DD_YY - MM_DD_YY.xlsx`
- Contains TGL validation information

### 2. File Date Verification

- Ensure all dated files (Jobs Report, Tech Department, TGL) are for the same week
- Files should cover Monday through Sunday of the target week
- Date ranges in filenames should match

## Running the Program

### 1. Initial Launch

- Run the program (`Payroll_Plus.py`)
- You'll see a welcome message and file checklist
- Confirm all files are ready by entering 'Y'

### 2. Date Entry

- Enter the target date in MM/DD/YY format
- The program will calculate the full week (Monday-Sunday) containing your date
- Example: Entering "01/15/24" will process the week of 01/15/24 - 01/21/24

### 3. UUID File Selection

If multiple UUID files exist:

- The program will suggest using the most recent file
- Enter 'Y' to accept or 'N' to choose a different file
- If choosing different file, enter the exact filename when prompted

## Processing Steps

The program will automatically:

1. Validate all required files
2. Create an output directory for the week
3. Combine data from all sources
4. Calculate service technician commissions
5. Process installer GP entries
6. Handle TGLs and spiffs
7. Generate all output files

## Output Review

### 1. Location

Find outputs in a new folder in your Downloads directory:

- Folder name format: `Commission Output MM_DD_YY-MM_DD_YY`

### 2. Key Files to Review

a) **paystats.xlsx**

- Review commission calculations
- Check threshold calculations and adjustments
- Verify revenue totals and percentages

b) **payroll.xlsx**

- Review final payroll entries
- Verify commission rates and amounts
- Check department codes

c) **Spiffs.xlsx**

- Review processed spiff entries
- Verify TGL applications
- Check department assignments

## Troubleshooting

### Common Issues and Solutions

1. **Missing Files**
   - Error: "UUID file not found"
   - Solution: Ensure file is in Downloads folder with correct name
   - Check for hidden file extensions

2. **Date Range Mismatch**
   - Error: "Found file for different week"
   - Solution: Download correct file for target week
   - Verify all file dates match

3. **Invalid Data**
   - Error: "Error processing entry"
   - Solution: Check source files for:
     - Missing values
     - Incorrect formats
     - Invalid department codes

4. **Processing Errors**
   - Check the generated log file in the output directory
   - Review error messages for specific issues
   - Verify source data integrity

### Validation Checks

Before submitting results, verify:

1. All technicians are processed
2. Commission rates match expectations
3. Department codes are correct
4. Spiffs are properly applied
5. TGL counts match records

## Best Practices

1. **File Management**
   - Keep source files organized
   - Don't modify original files
   - Back up output files

2. **Data Verification**
   - Review logs for warnings
   - Cross-check critical calculations
   - Verify department assignments

3. **Regular Maintenance**
   - Archive old output folders
   - Keep time off records updated
   - Monitor for system updates

4. **Error Prevention**
   - Double-check date entries
   - Verify file names exactly match
   - Review source data before processing

## Support and Updates

If you encounter issues:

1. Check the log file for specific errors
2. Verify all file requirements
3. Ensure dates and formats match
4. Review troubleshooting steps above
5. Contact system administrator for additional support
