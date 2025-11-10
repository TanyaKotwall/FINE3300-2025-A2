# FINE3300-2025-A2
This repository is for Assignment 2 in FINE 3300 course which contains code in Python using Visual Studio Code. The assignment extends the first project to include advanced mortgage amortization schedules in Part A and also include analysis on the Consumer Price Index (CPI) in Part B.

## Contents
### Loan Amortization
`LoanAmortization.py` expands the **MortgagePayment** class from Assignment 1 to generate complete loan amortization schedules for six different payment frequencies.  
It converts a quoted annual interest rate (semi-annually compounded) into an effective annual rate and calculates payments using the **Present Value of an Annuity (PVA)** formula. The program prompts the user for four key inputs and returns payment details, schedules, and a visualization of the loan balance decline across all payment options.  

**User Inputs:**
- Principal amount ($)
- Annual quoted interest rate (%)
- Amortization period (years)
- Term of the mortgage (years)

**Outputs:**
- Monthly
- Semi-Monthly
- Bi-Weekly
- Weekly
- Rapid Bi-Weekly
- Rapid Weekly  

Each option generates a full amortization table with:
- **Period**
- **Starting Balance**
- **Interest Amount**
- **Payment**
- **Ending Balance**

**Files Produced using Pandas, Matplotlib, Numpy, Openpyxl:**
- `Payment_Schedules.xlsx` — Excel file with six worksheets (one for each payment type)
- `Loan_Balance_Decline.png` — Matplotlib.pyplot and Matplotlib.ticker chart showing balance decline across all schedules  

### Consumer Price Index (CPI)
`CPI_Analysis.py` analyzes monthly Consumer Price Index (CPI) data for 2024 across Canada and its provinces using data from **Statistics Canada**.  
It estimates inflation, compares price changes across categories, and evaluates real wages by province.

**User Inputs:**
- 11 CPI CSV files (Canada + 10 provinces)
- `MinimumWages.csv` for nominal and real wage comparison

**Main Tasks:**
- Combine CPI files into one DataFrame: *Item, Month, Jurisdiction, CPI*  
- Show first 12 rows of combined data  
- Calculate average monthly % change for *Food*, *Shelter*, and *All-items excl. food & energy*  
- Identify provinces with the highest average change  
- Compute salary equivalents for $100,000 in Ontario (Dec 2024)  
- Compare nominal vs real minimum wages  
- Calculate annual % change in *Services CPI* and find region with highest inflation  

**Outputs:**
- Combined CPI summary  
- Monthly change table by province  
- Salary and wage comparison tables  
- Services inflation summary

**Files Produced using Pandas:**
- `CPI_Analysis_Results.xlsx` — Excel file with worksheets that answers all questions related to CPI