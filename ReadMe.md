# Payroll Processing System

## Overview

This system automates the processing of payroll data, including commission calculations for service technicians and GP (Gross Profit) entries for installers. It handles various types of compensation including PCM (Service Tech Commission), ICM (Installer Commission), and SPF (Spiffs/TGL).

## Required Files

All files must be in your Downloads folder:

1. UUID file (format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.xlsx`)
2. Jobs Report (format: `Copy of Jobs Report for Performance -DE2_Dated MM_DD_YY - MM_DD_YY.xlsx`)
3. Tech Department file (format: `Technician Department_Dated MM_DD_YY - MM_DD_YY.xlsx`)
4. Time Off file (name: `Approved_Time_Off 2023.xlsx`)
5. TGL file (format: `TGLs Set _Dated MM_DD_YY - MM_DD_YY.xlsx`)

## Features

- Separate processing for service technicians and installers
- Commission calculations based on department thresholds
- Automatic handling of TGLs and spiffs
- Adjustment processing for positive and negative entries
- Comprehensive logging and error handling
- Data validation and cleaning
- Automated file organization and output generation

## Department Codes

- HVAC: 20-29
  - 20: HVAC SERVICE
  - 21: HVAC INSTALL
  - 22: MAINTENANCE MVP
  - 24: OIL SERVICE
  - 25: OIL MAINTENANCE MVP
  - 27: HVAC DUCT CLEANING
- Plumbing: 30-39
  - 30: PLUMBING SERVICE
  - 31: PLUMBING INSTALL
  - 33: PLUMBING DRAIN CLEANING
  - 34: PLUMBING EXCAVATION
- Electric: 40-49
  - 40: ELECTRICAL SERVICE
  - 41: ELECTRICAL INSTALL
  - 42: GENERATOR MAINTENANCE

## Output Files

The system generates several output files in a dated directory:

- `combined_data.xlsx`: Consolidated data from all input files
- `paystats.xlsx`: Commission calculations and metrics
- `payroll.xlsx`: Final payroll entries
- `Spiffs.xlsx`: Processed spiff entries
- `positive_adjustments.xlsx`: Reference file for positive adjustments
- `negative_adjustments.xlsx`: Reference file for negative adjustments

## Commission Calculation Details

### Service Technicians (PCM)

- Commission rates: 2%, 3%, 4%, or 5% based on thresholds
- Thresholds adjusted by:
  - Install Contribution Percentage (ICP)
  - Excused time off
  - Valid TGLs
- Revenue components:
  - Completed Job Revenue (CJR)
  - Tech-Sourced Install Sales (TSIS)

### Installers (ICM)

- GP-based compensation
- Processed by department and job
- Automatic consolidation of multiple entries

### Spiffs and TGLs

- Automatic processing of positive and negative spiffs
- Department-specific TGL validation
- Consolidated negative spiffs applied to PCM entries

## Usage

1. Ensure all required files are in the Downloads folder
2. Run the script
3. Enter the target date when prompted (format: MM/DD/YY)
4. Select UUID file if multiple versions exist
5. Review output files in the generated directory

## Error Handling

- File validation before processing
- Comprehensive logging of all operations
- Clear error messages for missing or invalid files
- Automatic data cleaning and validation

## Technical Notes

- Uses Python with pandas for data processing
- Excel manipulation via openpyxl
- Configured for J6P company code and L100 location ID
- Handles multiple file formats and data structures
- Implements robust error checking and validation

## Dependencies

- pandas
- openpyxl
- logging
- datetime
- pathlib
- re (regular expressions)
