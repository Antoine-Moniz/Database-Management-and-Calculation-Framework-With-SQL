

# VBA Project: Pricing of Monetary Notes

This project is designed to automate the management of monetary note pricing using VBA, Access databases, and Excel macros. The application involves retrieving, structuring, and analyzing monetary and issuance data to calculate dynamic prices for financial instruments. The final output includes a client template for pricing analysis.

---

## Project Objectives

### **1. Data Retrieval and Preparation**
- **MONETARY_INDEX_DATA**:
  - Collect data for key monetary indices from 01/01/2016 to 31/12/2023:
    - EONIA, ESTER, EURIBOR 3M, LIBOR 3M USD, FED FUNDS, SONIA GBP.
- **NOTES_DATA**:
  - Compile key attributes of monetary notes, such as ISIN, issue date, maturity date, issue price, currency, reference rate, and spread.
- **NOTE_RATE_DATA**:
  - Gather rerating data for monetary notes, including ISIN, rerating date, reference rate, and spread.

### **2. Database Construction**
- Automate the creation of an Access database with the following tables:
  1. **MONETARY_INDEX**:
     - Stores monetary indices data with fields for dates and index values (EONIA, ESTER, etc.).
  2. **ISSUANCE_STATIC_DATA**:
     - Contains static data for monetary notes, including ISIN, issue date, maturity date, issue price, currency, reference rate, and spread.
  3. **ISSUANCE_RATE_DATA**:
     - Holds rerating data with fields for ISIN, rerating date, reference rate, and spread.
  4. **ISSUANCE_DYNAMIC_DATA**:
     - Stores dynamic pricing data for monetary notes with fields for pricing date, ISIN, and calculated price.
- Populate each table with data from the corresponding Excel files.

### **3. Pricing Model and Dynamic Data**
- Develop a VBA function to calculate the price of a monetary note based on:
  - Initial spread.
  - Reference rate at a specific pricing date.
  - Rerating spreads.
  - The duration between the issue date and the pricing date.
- Automatically populate the `ISSUANCE_DYNAMIC_DATA` table with calculated prices for all ISINs and pricing dates.

### **4. Client Template**
- Create an Excel template (`NOTE_TEMPLATE`) with the following features:
  1. **Clean Button**:
     - Clears data while retaining titles.
  2. **Print Button**:
     - Fetches and displays pricing data for a given ISIN and pricing date, adapting rerating data.
  3. - **Email Integration**:
       - A button to email the file with the ISIN in the filename and subject.
     - **Price Evolution Graph**:
       - Visualize the price history of a note in the client template and optionally in the email body.

---
