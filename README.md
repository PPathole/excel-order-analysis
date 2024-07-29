# Order Data Management in Excel

## Overview

This project involves analyzing order data using Excel. The dataset includes columns like Order ID, Order Date, Order Priority, Order Quantity, Order Type, Expanded Order Type, Sales, Discount, Sales with Discount, Sales with Free Shipping, Ship Mode, Shipping Cost, and SalesDeliveryTruck. The project includes creating data validations, conditional formatting, and utilizing advanced Excel formulas like VLOOKUP, XLOOKUP, and INDEX MATCH to derive insights and automate tasks.

## Features

1. **Data Validation:**
   - Created a drop-down list for the `Ship Mode` column using data validation to ensure consistent data entry.
   - Steps:
     1. Select the `Ship Mode` column.
     2. Go to `Data` > `Data Validation`.
     3. Set the validation criteria to `List` and enter the possible ship modes.

2. **Conditional Formatting:**
   - Applied conditional formatting to highlight high-priority orders.
   - Steps:
     1. Select the range.
     2. Go to `Home` > `Conditional Formatting`.
     3. Add a new rule and set the condition based on the `Order Priority` column.

3. **Formulas:**
   - **VLOOKUP:**
     - Used to find the `Shipping Cost` based on `Order ID`.
     - Formula: `=VLOOKUP(OrderID, TableRange, ColumnIndex, FALSE)`
   - **XLOOKUP:**
     - An advanced version of VLOOKUP that can search in both directions.
     - Formula: `=XLOOKUP(OrderID, LookupArray, ReturnArray)`
   - **INDEX MATCH:**
     - A more flexible alternative to VLOOKUP.
     - Formula: `=INDEX(ReturnRange, MATCH(LookupValue, LookupRange, 0))`

## How to Use

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/PPathole/excel-order-analysis.git
   cd excel-order-analysis
   ```

2. **Open the Excel File:**
   - Open the `OrderData.xlsx` file in Excel.

3. **Data Validation:**
   - Follow the steps under `Features > Data Validation` to set up the drop-down list for `Ship Mode`.

4. **Conditional Formatting:**
   - Apply conditional formatting as described under `Features > Conditional Formatting`.

5. **Using Formulas:**
   - Insert the provided formulas in the appropriate cells to automate data lookup and analysis.

## Example Formulas

- **VLOOKUP Example:**
  ```excel
  =VLOOKUP(A2, $A$2:$G$100, 4, FALSE)
  ```

- **XLOOKUP Example:**
  ```excel
  =XLOOKUP(A2, $A$2:$A$100, $G$2:$G$100)
  ```

- **INDEX MATCH Example:**
  ```excel
  =INDEX($G$2:$G$100, MATCH(A2, $A$2:$A$100, 0))
  ```
