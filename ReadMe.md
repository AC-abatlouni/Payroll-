# Technician Commission Policy

## Table of Contents

- [Overview](#overview)

1. [Revenue Sources and Definitions](#1-revenue-sources-and-definitions)
2. [Calculating SCP and ICP](#2-calculating-scp-and-icp)
3. [Determining Weekly Sales Thresholds Based on SCP and ICP](#3-determining-weekly-sales-thresholds-based-on-scp-and-icp)
4. [Adjusting Sales Thresholds for Paid Days Off](#4-adjusting-sales-thresholds-for-paid-days-off)
5. [Applying Threshold Reductions for Tech Generated Leads (TGLs)](#5-applying-threshold-reductions-for-tech-generated-leads-tgls)
6. [Understanding Spiffs and Their Impact](#6-understanding-spiffs-and-their-impact)
7. [Calculating Commission Rate and Commission Amount](#7-calculating-commission-rate-and-commission-amount)
8. [Handling Refunds and Adjustments](#8-handling-refunds-and-adjustments)
9. [Step-by-Step Calculation Summary](#9-step-by-step-calculation-summary)
10. [Additional Notes](#10-additional-notes)
11. [Key Takeaways](#11-key-takeaways)

---

## Overview

Technicians earn commissions based on their revenue contributions and performance. Several key components determine how the commission is calculated.

## 1. Revenue Sources and Definitions

### Scenario 1 - Completed Job Revenue (CJR)

- **Definition**: Revenue from jobs that the technician completes directly without involving other teams
- **Notation**: Box A
- **Recognition**: Revenue is counted in the week the work is completed
- **Example**: If a technician completes a repair job worth $500, this amount is added to Box A.

### Scenario 2 - Tech-Sourced Install Sales (TSIS)

- **Definition**: Revenue from installs initiated by the technician but completed by another team
- **Notation**: Box B
- **Recognition**: Revenue is counted in the week the installation is completed
- **Example**: If a technician convinces a customer to install a new appliance worth $500, this amount is added to Box B when the installation is complete.

### Total Revenue

- **Definition**: Sum of CJR and TSIS
- **Notation**: Box C
- **Formula**: `Box C = Box A + Box B`
- **Example**: If Box A is $5,790 and Box B is $13,860, then Box C is $19,650.

## 2. Calculating SCP and ICP

### Service Completion Percentage (SCP)

- **Formula**: `SCP = (Box A / Box C) × 100`
- **Definition**: Percentage of total revenue from directly completed jobs
- **Rounding**: To nearest 10% using standard rounding rules
  - 13% becomes 10%
  - 75% becomes 80%
- **Example**: `SCP = (5,790 / 19,650) × 100 = 29.49% → 30%`

### Install Contribution Percentage (ICP)

- **Formula**: `ICP = (Box B / Box C) × 100`
- **Definition**: Percentage of total revenue from tech-sourced installs
- **Rounding**: To nearest 10% using standard rounding rules
- **Example**: `ICP = (13,860 / 19,650) × 100 = 70.51% → 70%`

## 3. Determining Weekly Sales Thresholds Based on SCP and ICP

### Commission Structure

- **Rates**: Range from 2% to 5%
- **Thresholds**: Four levels per percentage tier
- **Default**: Below lowest threshold = 0% commission
- **Maximum**: Above highest threshold = 5% commission

### HVAC Department Thresholds

| Flipped % | Threshold 1 (2%) | Threshold 2 (3%) | Threshold 3 (4%) | Threshold 4 (5%) |
|-----------|-------------------|-------------------|-------------------|-------------------|
| 0         | $7,000           | $8,000           | $9,000           | $10,000          |
| 10        | $7,500           | $8,500           | $9,500           | $11,000          |
| 20        | $8,000           | $9,000           | $10,000          | $11,000          |
| 30        | $9,000           | $10,000          | $11,000          | $12,000          |
| 40        | $10,000          | $11,000          | $12,000          | $13,000          |
| 50        | $12,000          | $14,000          | $16,000          | $18,000          |
| 60        | $14,000          | $16,000          | $18,000          | $20,000          |
| 70        | $15,000          | $17,000          | $19,000          | $21,000          |
| 80        | $16,500          | $18,500          | $20,500          | $23,000          |
| 90        | $18,500          | $20,500          | $23,500          | $26,000          |
| 100       | $22,000          | $24,000          | $26,000          | $29,000          |

### Plumbing and Electrical Departments Thresholds

| Flipped % | Threshold 1 (2%) | Threshold 2 (3%) | Threshold 3 (4%) | Threshold 4 (5%) |
|-----------|-------------------|-------------------|-------------------|-------------------|
| 0         | $7,000           | $8,000           | $9,000           | $10,000          |
| 10        | $7,500           | $8,500           | $9,500           | $11,000          |
| 20        | $8,000           | $9,000           | $10,000          | $11,000          |
| 30        | $9,000           | $10,000          | $11,000          | $12,000          |
| 40        | $10,000          | $11,000          | $12,000          | $13,000          |
| 50        | $11,500          | $12,500          | $14,000          | $15,000          |
| 60        | $13,000          | $14,000          | $15,500          | $17,000          |
| 70        | $14,500          | $15,500          | $17,000          | $18,000          |
| 80        | $15,500          | $16,500          | $18,000          | $19,000          |
| 90        | $17,000          | $18,500          | $19,500          | $21,000          |
| 100       | $18,000          | $20,000          | $22,000          | $24,000          |

## 4. Adjusting Sales Thresholds for Paid Days Off

### Adjustment Rules

- **Rule**: 20% reduction per approved day off
- **Maximum Reduction**: Up to 5 days off (100% reduction)
- **Manager Override**: Discretion to approve reductions for unpaid time off
- **Documentation**: Must be recorded in "Approved_Time_Off" worksheet

### Calculation Formula

- **Reduction Factor**: `1 - (0.20 × Number of Days Off)`
- **Adjusted Threshold**: `Original Threshold × Reduction Factor`
- **Example**:
  - **Original Threshold**: $9,000
  - **Days Off**: 1 day
  - **Reduction Factor**: `1 - (0.20 × 1) = 0.80`
  - **Adjusted Threshold**: `9,000 × 0.80 = $7,200`

## 5. Applying Threshold Reductions for Tech Generated Leads (TGLs)

### TGL Threshold Reduction Rules

| Component          | Rule                                                    |
|--------------------|---------------------------------------------------------|
| Eligibility        | Only same-department TGLs qualify for threshold reduction |
| Calculation Base   | Uses average ticket value                                |
| Minimum Sale       | Must exceed $2,000 for spiff qualification               |
| Payment Timing     | Paid after installation completion                       |

### Department Business Units

| Department | Business Unit Range |
|------------|---------------------|
| HVAC       | 20-29              |
| Plumbing   | 30-39              |
| Electric   | 40-49              |

### Calculation Formulas

- **Average Ticket Value**: `Sum of Ticket Values / Number of Tickets`
- **Total Threshold Reduction**: `Average Ticket Value × Number of Same-Department TGLs`
- **Final Adjusted Threshold**: `max(0, Adjusted Threshold - Total Threshold Reduction)`

**Example**:

- **Tickets**: $1,000, $1,100, $1,050, $1,114.24
- **Average Ticket**: `(1,000 + 1,100 + 1,050 + 1,114.24) / 4 = $1,066.06`
- **Same-Department TGLs**: 2
- **Total Reduction**: `1,066.06 × 2 = $2,132.12`

## 6. Understanding Spiffs and Their Impact

### Spiffs Overview

- **Definition**: Additional bonuses/incentives earned by technician
- **Source**: "Direct Payroll Adjustments" sheet in Payroll Detail report
- **Timing**: Paid after item installation
- **Treatment**: Not added to threshold qualification; subtracted from commissionable revenue

### Impact on Calculations

- **Revenue for Threshold Qualification**: Total Revenue (Spiffs not included)
- **Commissionable Revenue**: `Total Revenue - Spiffs`

**Example**:

- **Total Revenue**: $9,900
- **Spiffs**: $200
- **Revenue for Thresholds**: $9,900
- **Commissionable Revenue**: `9,900 - 200 = $9,700`

## 7. Calculating Commission Rate and Commission Amount

### Commission Rate Determination

1. **Calculate Adjusted Revenue**: Use Total Revenue without Spiffs and TGL Spiffs
2. **Compare to Adjusted Thresholds**: Match against thresholds
3. **Select Highest Qualifying Rate**: Based on highest qualifying threshold met

### Revenue Calculations

- **Adjusted Revenue**: Total Revenue (excluding Spiffs and TGL Spiffs)
- **Commissionable Revenue**: `Total Revenue - Spiffs`
- **Commission**: `Commissionable Revenue × Commission Rate`

**Note**: TGL Spiffs are not deducted from commissionable revenue.

## 8. Handling Refunds and Adjustments

### Jobs Report Treatment

| Scenario         | Handling                               |
|------------------|----------------------------------------|
| Negative Totals  | Exclude from Total Sales calculations  |
| Adjustments      | Calculated manually                    |
| Documentation    | Processed in Payroll Detail report     |

### Direct Payroll Adjustments

| Type               | Treatment                                              |
|--------------------|--------------------------------------------------------|
| Negative Amounts   | Show as "Payroll Adjustments"                          |
| Commission Impact  | Can result in negative commission                      |
| Processing         | Handled manually by Accounting                         |
| Identification     | Look for "commission" or "pcom" in memo                |

### Spiff Refunds

- **Reversed**: When job is refunded
- **Appearance**: Negative amounts in Direct Payroll Adjustments
- **Timing**: Processed when refund occurs

## 9. Step-by-Step Calculation Summary

1. **Calculate Total Revenue**
   - Add CJR (Box A) and TSIS (Box B)
   - Count revenue in completion/installation week
2. **Calculate and Round Percentages**
   Service Completion Percentage (SCP): Calculate (Box A / Box C) × 100, rounding to the nearest 10%.
   Install Contribution Percentage (ICP): Calculate (Box B / Box C) × 100, rounding to the nearest 10%.
3. **Determine Base Thresholds**
   - Use department-specific tables
   - Match rounded percentages to tiers

4. **Apply Adjustments**
   - Time off reductions
   - TGL reductions
   - Maintain minimum of zero

5. **Calculate Adjusted Revenue**
   - Use total revenue
   - Exclude Spiffs and TGL Spiffs

6. **Determine Commission Rate**
   - Compare to adjusted thresholds
   - Select highest qualifying rate

7. **Calculate Final Commission**
   - Apply rate to commissionable revenue
   - Add Spiffs separately
   - Process adjustments

## 10. Additional Notes

### Source Documents

| Document                     | Purpose                               |
|------------------------------|---------------------------------------|
| Payroll Detail Report        | Spiffs, TGL Spiffs, Adjustments       |
| Jobs Report                  | Revenue tracking, Sales data          |
| Approved Time Off Worksheet  | Time off verification                 |

### Timing Guidelines

| Event         | Recognition                  |
|---------------|------------------------------|
| Revenue       | Week of completion           |
| Sales         | Week of installation         |
| Spiffs        | After installation           |
| Refunds       | When processed               |

## 11. Key Takeaways

### Critical Rules

- **Revenue Recognition**
  - All revenue counts in completion week
  - Sales count at installation

- **TGL Requirements**
  - $2,000 minimum for spiffs
  - Department matching required
  - Same-department for reductions

- **Spiffs Treatment**
  - Deduct from commissionable revenue
  - TGL Spiffs handled separately
  - Both paid post-installation

- **Threshold Adjustments**
  - Time off with manager discretion
  - Same-department TGL reductions
  - Cannot go negative

- **Refund Processing**
  - Exclude from Total Sales
  - Process as adjustments
  - May create negative commission
