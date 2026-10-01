## 📌 Project Overview

This project focuses on cleaning, standardizing, validating, and preparing an e-commerce sales dataset using Microsoft Excel.
The objective was to identify data-quality issues in the raw dataset, apply appropriate Excel-based cleaning and validation techniques, and prepare a structured dataset suitable for further analysis.

## 🎯 Project Objectives
- Identify duplicate and potentially duplicated records
- Standardize date fields and verify date data types
- Clean customer, country, city, and product-category fields
- Standardize payment methods and order status
- Convert quantity and financial fields into usable numeric formats
- Validate product-category consistency
- Flag potentially invalid quantities, discounts, sales, ratings, and dates
- Identify missing values
- Prepare a clean dataset for subsequent analysis

## 🛠️ Tools & Excel Features Used
- Microsoft Excel
- Sorting & Filtering
- Conditional Formatting
- Remove Duplicates
- Find & Replace
- Flash Fill
- Text to Columns
- Paste Special → Values
- Excel formulas such as `TRIM`, `PROPER`, `TEXTSPLIT`, `LET`, `TEXT`, `ISNUMBER`, `SWITCH`, and `IF`

## 1. Initial Data Preparation

### Freeze Top Row

**Issue addressed:**  
Column headers can disappear while scrolling through a large dataset, making it difficult to identify variables.

**Action:**  
Used **View → Freeze Panes → Freeze Top Row**.

### Create a Working Copy

**Issue addressed:**  
Cleaning directly on the original dataset can result in accidental data loss or irreversible changes.

**Action:**  
Created a working copy of the original worksheet by holding **Ctrl** and dragging the worksheet tab.

### Apply Filters

**Issue addressed:**  
Large datasets are difficult to inspect manually without the ability to isolate specific values.

**Action:**  
Applied filters using **Data → Filter**.


## 2. Duplicate Data Identification

### Identify Duplicate Values

**Issue addressed:**  
Duplicate values can indicate repeated IDs, repeated attributes, or potential data-entry problems.

**Action:**  
Used **Home → Conditional Formatting → Highlight Cells Rules → Duplicate Values**.
> A duplicate value in one column does not necessarily mean that the entire row is duplicated.

### Identify Exact Duplicate Rows

**Issue addressed:**  
Rows with identical values across all columns may represent complete duplicate records.

**Action:**  
Used **Data → Remove Duplicates** with all columns selected.
This distinguishes complete duplicate records from records that only share one or more repeated values.

### Check Duplicate Order IDs

**Issue addressed:**  
Repeated Order IDs may indicate multiple records belonging to one order or possible duplication.

**Action:**  
Sorted by **Order ID**, reviewed repeated IDs, and investigated whether repeated IDs represented genuine multiple records or duplication.

## 3. Date Standardization

### Order Date

**Issue addressed:**  
The original Order Date values were in DMY format and were not consistently stored by Excel as proper date values.

### Step 1 — Verify Date Recognition

Used:

```excel
=ISNUMBER([@[Order Date]])
```

This checks whether Excel stores the value as a numerical date or as text.

### Step 2 — Identify the Date Convention

Used:

```excel
=LET(x,TEXTSPLIT(B2,"/"),a,--INDEX(x,1,1),b,--INDEX(x,1,2),IF(a>12,"DMY",IF(b>12,"MDY","Ambiguous")))
```

The analysis established that the original Order Date column followed **DMY** format.

### Step 3 — Convert Text Dates to Actual Excel Dates

Used:

**Data → Text to Columns → Delimited → Next → Next → Date: DMY → Finish**

This converted the values into valid Excel date values.

### Step 4 — Standardize Display Format

Used **Ctrl + 1 → Custom** and entered:

```text
mm/dd/yyyy
```

This standardized the display format to **MM/DD/YYYY**.

### Delivery Date

Applied the same date-conversion process to the Delivery Date field to ensure consistent date handling.

### Remove Temporary Verification Columns

Temporary columns created for date-type and date-format verification were removed after validation.

## 4. Customer Name Cleaning

**Issue addressed:**  
Customer names may contain unnecessary spaces and inconsistent capitalization.

Used:

```excel
=TRIM(PROPER(C2))
```

- `TRIM()` removes unnecessary spaces.
- `PROPER()` standardizes capitalization.

After validation, the cleaned column was copied and **Paste Special → Values** was used. The original improperly formatted Customer Name column was then removed.

## 5. Country Name Standardization

### Find & Replace

**Issue addressed:**  
Country values contained abbreviated/inconsistent representations such as `UK`.

Used:

```text
Find: UK
Replace: United Kingdom
```

### Flash Fill

A Clean Country Name column was also created and Flash Fill was used to apply the desired standardized pattern. The original column was removed after validation.

## 6. City Name Cleaning

**Issue addressed:**  
Leading, trailing, or unnecessary spaces can create incorrect grouping and duplicate city categories.

Used:

```excel
=TRIM(E2)
```

The cleaned City Name values were copied, pasted as values, and the original column was removed.

---

## 7. Product Category Cleaning

### Standardize Category Values

Used:

```excel
=TRIM(PROPER(H2))
```

The cleaned values were converted to values and the original Product Category column was removed.

### Correct Category Naming

Used Find & Replace to standardize variations beginning with `Electronic`:

```text
Find: Electronic*
Replace: Electronics
```

## 8. Extract Category Code

**Issue addressed:**  
Category codes embedded in product information were not separately structured.

Used **Flash Fill** to extract category codes.

Example:

```text
ELEC-1002-Y-Bluetooth Speaker
            ↓
           ELEC
```

A separate Category Code field was created for analysis and validation.

---

## 9. Derive Product Category from Category Code

A standardized category field was created using:

```excel
=SWITCH(H2,
"ELEC","Electronics",
"BEAU","Beauty",
"HOME","Home & Kitchen",
"FASH","Fashion")
```

## 10. Product Category Data-Quality Validation

A `Product_Category_Match` column was created to compare:

- The category recorded in the order
- The category derived from product/category information

The result identifies:

- `Match`
- `Mismatch`

A **Mismatch** was treated as a record requiring investigation rather than automatically assuming that one value was incorrect. Verification should be performed against the relevant business source or product master data.

## 11. Payment Method Cleaning

**Issue addressed:**  
Payment method values contained inconsistent descriptions, including variations beginning with `Credit`.

Used Find & Replace:

```text
Find: Credit*
Replace: Credit Card
```

This standardized the payment method representation to **Credit Card**.

## 12. Order Status Standardization

The Order Status field did not show significant spelling or spacing problems, but Flash Fill was used to standardize values based on the observed pattern.

After validation:

1. Copied the cleaned values.
2. Pasted as values.
3. Deleted the original Order Status column.

## 13. Quantity Field Cleaning

### Convert Quantity to Numeric Values

**Issue addressed:**  
Some Quantity values were not properly recognized as numeric values.

Used:

**Data → Text to Columns → Fixed Width → Next → Next → General → Finish**

This converted Quantity values into a consistent numeric format.

### Identify Zero and Negative Quantities

Used Conditional Formatting to highlight:

- `0`
- Negative values

These records were **flagged for business review rather than automatically deleted**, because they may represent returns, cancellations, exceptional transactions, or data-entry errors.

## 14. Amount / Currency Cleaning

### Remove Currency Symbol

**Issue addressed:**  
Currency symbols stored directly within values can cause numbers to be treated as text.

Used Find & Replace:

```text
Find: $
Replace: [blank]
```

### Apply Currency Formatting

Applied the appropriate currency format to the cleaned Amount column while retaining the values as numerical data.

## 15. Discount Field Cleaning

### Convert Discount Values to Percentage

**Issue addressed:**  
Discount values were not represented in the appropriate percentage format.

A separate cell containing `100` was used as the divisor, with an absolute cell reference.

Example:

```excel
=DiscountCell/$X$1
```

The resulting values were formatted as percentages.

For example:

```text
20 → 20%
```

### Identify Invalid Discount Values

Conditional Formatting was used to highlight values:

```text
>100%
```

These records were flagged for investigation.

---

## 16. Total Sales Validation

Applied US Dollar (`$`) currency formatting to the Total Sales column.

### Identify Negative Sales

Conditional Formatting was used to highlight negative Total Sales values.

Negative sales can represent refunds, returns, cancellations, credit adjustments, or data-entry issues, so these records were flagged for review rather than automatically removed.

## 17. Customer Rating Validation

Conditional Formatting was used to highlight Customer Rating values:

```text
>5
```

This identifies ratings outside the expected rating scale for review.

## 18. Order Date vs Delivery Date Validation

### Validate Date Sequence

A validation column was created to check whether Delivery Date occurs on or after Order Date.

Conceptually:

```excel
=IF(Delivery_Date>=Order_Date,"Valid","Check")
```

Records where Delivery Date occurred before Order Date were identified for investigation.

Conditional Formatting was also used to highlight:

```text
Delivery Date < Order Date
```

## 19. Missing Value Identification

**Issue addressed:**  
Blank values can affect calculations, filtering, aggregation, and analytical conclusions.

Used:

**Conditional Formatting → New Rule → Format only cells that contain → Blanks**

Blank values were identified rather than automatically deleted because the appropriate treatment depends on the business meaning of each field.

## 📊 Before and After Cleaning

### Before Cleaning

![Data Before Cleaning](Data%20before%20cleaning.png)

### After Cleaning

![Data After Cleaning](Data%20after%20cleaning.png)

## 📁 Project Files

| File | Description |


| `Ecommerce_Data_Cleaning_Dataset_Excel.xlsx` | Original e-commerce sales dataset |

| `Cleaned Ecommerce_Data.xlsx` | Cleaned dataset after applying the documented cleaning and validation steps |

| `Data before cleaning.png` | Screenshot of the dataset before cleaning |

| `Data after cleaning.png` | Screenshot of the dataset after cleaning |

## 📌 Key Data-Quality Practices Demonstrated

This project goes beyond basic formatting and demonstrates several important data-analytics practices:

- Duplicate detection and investigation
- Data-type verification
- Date standardization
- Text standardization
- Category normalization
- Derived-field creation
- Cross-field validation
- Exception identification
- Missing-value identification
- Preservation of potentially meaningful exceptional records
- Conversion of formula outputs to static values

## 💡 Skills Demonstrated

- Microsoft Excel
- Data Cleaning
- Data Preparation
- Data Standardization
- Data Validation
- Data Quality Assessment
- Conditional Formatting
- Excel Formulas
- Flash Fill
- Find & Replace
- Text to Columns
- Duplicate Detection
- Analytical Thinking
- Documentation

## ⚠️ Data Validation Note

Records flagged as potential errors were identified for investigation rather than automatically deleted where the business meaning was uncertain. 
