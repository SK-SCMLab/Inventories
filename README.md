# 🧺 Inventory Data Analysis using Microsoft Excel (Fundamental)
This repository demonstrates an end-to-end inventory data analysis project undertaking a sample steel manufacturing data using Advanced Excel techniques. It showcases how Excel can be used not just as a spreadsheet tool, but as a powerful platform for structured data transformation, and analysis. 

---

## 🖊 Key Features & Concepts Covered 
### 1. **Data analysis using Logical and dynamic functions**
- IF()/IFS(): Used for basic conditional logic to categorize and flag data based on specific criteria
- SWITCH(): Simplifies multi-condition logic by replacing nested 'IF' statement with a cleaner, more readable approach
- LET(): Introduces named variables inside formulas for better performance and clarity
- LAMBDA(): Creates custom, reusable function without VBA. Helps encapsulate logic for repeatable calculations
- RANDBETWEEN(): Generates sample datasets by producing random integer values within a defined range

### 2. **Data transormation using Lookup, Text, String and Reference functions**
- CONCAT()/TEXTJOIN(): Combine data across multiple cells into a single string
- TRIM(): Remove extra spaces from text (useful for cleaning imported or copied data)
- SUBSTITUTE(): Replace specific characters or words within strings
- LEFT()/RIGHT(): Extract fixed-length substrings from start or end of a value
- TEXT(): Format numeric values into readable text
- VLOOKUP(): Retrieve data from a table vertically based on a matching key
- HLOOKUP(): Similar to VLOOKUP, but works across rows
- MATCH(): Return the relative position of a value in a row or column
- INDEX(): Return the value of a cell in a specified row or column, often combined with MATCH()
### 2.1. **Data transformation using Power Query editor**
- Launch PowerQuery Editor from Data tab in Excel
- Load the data from existing table/range
- Group and aggregate the values as mentioned in the tasklist

### 3. **Excel Tables & Structured references**
- Converted raw data into Excel Tables for better structure, flexibility, and dynamic referencing
- Used structred references for consistency and forumla clarity throughout the workbook

### 4. **Dynamic Arrays**
- FILTER(): Apply row filters in Power Query using custom logic
- SORT(): Sort row data by one or more columns using ascending/descending order; nested sorting
- UNIQUE(): To extract distinct values from a column or entire table
- SEQUENCE(): To generate custom sequences for simulations or lookups
- XLOOKUP(): To retrive multiple related fields from a single lookup key

### 5. **Pivot Tables**
- Custom calculations using calculated fields
- Grouping by date (month, quarter, year)
- Report filters
- Slicers for easy navigation
- Value and label filters for focused analysis

## 📜 Repository Structure
Inventories_dataset_analysis.xlsx
- Sheet1: Original (Hidden & Protected)
- Sheet2: About the project
- Sheet3: Tasklist
- Sheet4: Analyze data
- Sheet5: Transform data
- Sheet6: Excel tables
- Sheet7: Dynamic Arrays
- Sheet8: Pivot Tables
- Sheet9: Visualization

---

## 🔦 Use cases
This project is ideal for:
- Supply Chain analysts optimizing delivery performance
- Business users looking to automate repetitive Excel tasks
- Learners interested in mastering Excel for real-world data problems

---
## 📍 Requirements
- Microsoft Excel 2016 or later (Dynamic Arrays and Power Query support)
- Basic understanding of inventory concepts (OTIF, Gross Profit)

---

_"Without data, you're just another person with an opinion." - W. Edwards Deming_
