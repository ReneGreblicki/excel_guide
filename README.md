# Excel for Data Analytics

A concise guide to spreadsheet analysis, Power Query, Power Pivot, M and DAX.

[Download the complete lesson notebook](download_notes/Excel_Lesson_Notebook.docx)  
[Download the 18-page visual cheatsheet](download_notes/Excel_Analytics_Cheatsheet_18_Pages.zip)

## Contents

- [Excel foundations and formulas](#02-workbook-structure-and-navigation)
- [Logic, text, dates, lookups and dynamic arrays](#09-logical-formulas-and-categories)
- [Statistics, PivotTables, charts and dashboards](#18-measures-of-centre-and-spread)
- [Power Query and M](#26-power-query-workflow)
- [Power Pivot, the Data Model and DAX](#30-the-data-model-and-star-schema)
- [Complete analytical workflow](#35-one-complete-excel-analysis)

This guide explains the Excel skills used to prepare data, calculate trustworthy metrics, explore patterns and present analytical results. Each lesson focuses on one idea so the syntax and the reasoning behind it stay connected.

| **Lessons** | **Focus**                                                               |
|-------------|-------------------------------------------------------------------------|
| 02–08       | Workbook structure, Tables, formulas, references and errors             |
| 09–17       | Logic, text, dates, lookups, summaries, arrays and reusable formulas    |
| 18–25       | Statistics, validation, PivotTables, charts, dashboards and forecasting |
| 26–31       | Power Query, M, the Data Model and relationships                        |
| 32–36       | DAX, end-to-end analysis, reliability and quick reference               |

**Analytical workflow:** define the question and grain → inspect and clean the data → calculate and compare → visualise and explain → reconcile and refresh.

Examples use Microsoft 365 function names and comma separators. Your regional settings may use semicolons. Newer functions may be unavailable in older Excel versions. This notebook contains no practice questions.

> **Key idea:** Before writing a formula, define what one output row, PivotTable cell or chart mark should represent.

## 02 Workbook structure and navigation

### Know where data logic and reporting should live

| **Part**    | **Meaning**                            | **Good practice**                         |
|-------------|----------------------------------------|-------------------------------------------|
| Workbook    | The .xlsx file                         | Keep raw data, logic and outputs separate |
| Worksheet   | A tab inside the workbook              | Use short descriptive names               |
| Cell        | One value at a row-column intersection | Do not mix labels and numbers             |
| Range       | A group such as A2:D100                | Prefer Tables for expanding data          |
| Formula bar | Displays or edits the active cell      | Press F2 to edit                          |
| Name box    | Displays an address or defined name    | Type A1000 to jump there                  |

A dependable workbook normally has four layers: **Raw** for unchanged imports, **Clean** for typed and standardised data, **Model** for relationships and calculations, and **Report** for PivotTables, charts and conclusions.

- Ctrl + Arrow moves to the edge of a data region; Ctrl + Shift + Arrow selects to that edge.

- Ctrl + G opens Go To; Ctrl + F finds; Ctrl + H replaces; Ctrl + S saves.

- Use Formulas → Show Formulas to display formulas instead of results. F4 cycles reference locking while editing a formula.

> **Key idea:** A workbook is easier to audit when imported data, transformation logic, calculations and presentation are visibly separated.

## 03 Grain, tidy data and data types

### Define what one row means before calculating

The **grain** is the real-world meaning of one row. In a Sales table, one row might represent an order, an order line or a daily summary. Mixing grains creates believable but incorrect totals.

| **OrderID** | **OrderDate** | **Region** | **Product** | **Units** | **UnitPrice** | **Status** |
|-------------|---------------|------------|-------------|-----------|---------------|------------|
| 1001        | 2026-01-05    | North      | Laptop      | 2         | 950.00        | Paid       |
| 1002        | 2026-01-06    | West       | Monitor     | 3         | 240.00        | Paid       |
| 1003        | 2026-01-09    | North      | Keyboard    | 5         | 70.00         | Pending    |

- Use one header row and one field per column.

- Keep one observation per row and one data type per column.

- Avoid merged cells, totals inside raw data, decorative blank rows and repeated headers.

- Store IDs as text when leading zeros matter. Store quantities and money as numbers, not text.

- Treat zero, blank, error and unknown as different analytical states.

Before analysis, compare total rows with distinct business keys. If 10,000 rows contain only 8,000 distinct OrderIDs, determine whether repeated IDs are valid order lines or duplicate records.

> **Key idea:** A formula cannot correct an undefined grain. Resolve the row meaning first.

## 04 Excel Tables and defined names

### Create stable inputs that expand and remain readable

Select a cell in a range and press **Ctrl + T**. Confirm the header row, then assign a short Table name under Table Design, such as Sales.

| **Feature**           | **Why it matters**                                      |
|-----------------------|---------------------------------------------------------|
| Automatic expansion   | New rows join formulas, formats and source ranges       |
| Structured references | Formulas use field names instead of fragile coordinates |
| Built-in filters      | Sort, filter and search from each header                |
| Total Row             | Uses SUBTOTAL functions that respond to filters         |
| Stable source         | Supports PivotTables, charts and Power Query refresh    |

    =[@Units]*[@UnitPrice]
    =SUM(Sales[Revenue])
    =SUMIFS(Sales[Revenue],Sales[Region],H2)

**\[@Units\]** means Units on the current Table row. **Sales\[Revenue\]** means the complete Revenue column. The name describes the logic, and the reference expands when the Table grows.

Defined names are useful for assumptions such as TaxRate or StartDate. They make formulas easier to read, but names must be documented and scoped deliberately to the workbook or worksheet.

> **Key idea:** Use Tables for datasets and defined names for important assumptions. Do not use formatting as a substitute for structure.

## 05 Formula anatomy and operators

### Read a formula as a sequence of instructions

    =ROUND([@Units]*[@UnitPrice]*(1-$H$2),2)

| **Part**   | **Meaning**                           |
|------------|---------------------------------------|
| =          | Starts a formula                      |
| ROUND(...) | Function and its arguments            |
| \* and -   | Arithmetic operators                  |
| \[@Units\] | Value from the current Table row      |
| \$H\$2     | A reference locked in both directions |
| 2          | Return two decimal places             |

Excel evaluates parentheses first, followed by exponentiation, multiplication and division, addition and subtraction, then comparisons. Parentheses should show the intended logic even when the default order would produce the same result.

| **Operator**         | **Idea**                           | **Example**         |
|----------------------|------------------------------------|---------------------|
| \+ - \* /            | Arithmetic                         | Revenue-Cost        |
| ^                    | Exponent                           | Growth^Years        |
| &                    | Join text                          | Region&" - "&Year   |
| = \<\> \> \< \>= \<= | Comparison returning TRUE or FALSE | \[@Revenue\]\>=1000 |

> **Key idea:** A long formula becomes easier to verify when you identify its inputs, transformations and final output separately.

## 06 Cell and structured references

### Control what changes when a formula is copied

| **Reference** | **Copied right** | **Copied down** | **Use**              |
|---------------|------------------|-----------------|----------------------|
| A1            | B1               | A2              | Relative calculation |
| \$A\$1        | \$A\$1           | \$A\$1          | Fixed assumption     |
| \$A1          | \$A1             | \$A2            | Lock the column      |
| A\$1          | B\$1             | A\$1            | Lock the row         |

A **relative reference** moves with the formula. An **absolute reference** stays fixed. A mixed reference locks only the row or the column. Press F4 while editing a reference to cycle through these four forms.

    =$B2*C$1
    ='Sales Data'!A1
    =SUM(January:December!B5)

The first formula locks column B and row 1. A worksheet name containing spaces needs single quotes. A 3-D reference sums the same cell across a continuous range of sheets, but it is harder to audit than a single normalised Table.

Structured references are safer for row-level Table logic because field names remain meaningful after columns move. Coordinate references are still useful for compact calculation areas and assumption cells.

> **Key idea:** Copy a formula across several rows and inspect the first, middle and last references before trusting the result.

## 07 Core calculations

### Choose the simplest function that matches the analytical question

| **Function** | **Question answered**               | **Example**                     |
|--------------|-------------------------------------|---------------------------------|
| SUM          | What is the total?                  | =SUM(Sales\[Revenue\])          |
| AVERAGE      | What is the arithmetic mean?        | =AVERAGE(Sales\[Revenue\])      |
| MIN / MAX    | What are the extremes?              | =MAX(Sales\[Revenue\])          |
| COUNT        | How many numeric cells?             | =COUNT(Sales\[Revenue\])        |
| COUNTA       | How many non-empty cells?           | =COUNTA(Sales\[OrderID\])       |
| ROUND        | What precision should be returned?  | =ROUND(G2,2)                    |
| ABS          | What is the unsigned difference?    | =ABS(New-Old)                   |
| SUBTOTAL     | What is the visible filtered total? | =SUBTOTAL(109,Sales\[Revenue\]) |

    Share % = Part / Total
    Growth % = (New - Old) / Old
    Weighted average = SUMPRODUCT(Values,Weights) / SUM(Weights)

A percentage is a ratio, not a stored unit. Format the decimal result as a percentage. A weighted average gives observations different influence; confirm that weights use the intended scale and denominator.

> **Key idea:** Use one clear function before combining several. Complexity should come from the question, not from the formula.

## 08 Errors and formula debugging

### Find the cause before hiding the symptom

| **Error**          | **Meaning**               | **First check**                |
|--------------------|---------------------------|--------------------------------|
| \#DIV/0!           | Zero or blank denominator | Denominator and missing policy |
| \#N/A              | No result available       | Lookup key or missing data     |
| \#VALUE!           | Wrong value type          | Numbers stored as text         |
| \#REF!             | Invalid reference         | Deleted or moved cells         |
| \#NAME?            | Unknown name or function  | Spelling and Excel version     |
| \#SPILL!           | Array result is blocked   | Cells in the spill area        |
| Circular reference | Formula depends on itself | Direct or indirect loop        |

    =IF(B2=0,NA(),(C2-B2)/B2)
    =IFERROR(XLOOKUP(A2,Keys,Results),"Not found")

**IFERROR** replaces every error produced by an expression. Use it only when the fallback has a defined meaning; otherwise it can hide broken references, bad types and logic errors. **NA()** returns \#N/A, which charts normally treat as unavailable.

- Use Formulas → Evaluate Formula to step through nested logic.

- Use Trace Precedents and Trace Dependents to inspect inputs and downstream effects.

- Use Formulas → Show Formulas to reveal formulas and compare neighbouring rows for formula drift.

> **Key idea:** An error is information. Handle it only after you understand why it occurred.

## 09 Logical formulas and categories

### Turn business rules into explicit ordered tests

    =IF([@Revenue]>=1000,"Large","Standard")
    =IFS([@Revenue]>=1000,"Large",[@Revenue]>=500,"Medium",TRUE,"Small")
    =IF(AND([@Status]="Paid",[@Revenue]>0),"Include","Exclude")
    =SWITCH([@Region],"N","North","S","South","Unknown")

| **Function**           | **Idea**                            |
|------------------------|-------------------------------------|
| IF(test,true,false)    | Choose between two outcomes         |
| IFS(test1,result1,...) | Apply several ordered rules         |
| AND(...)               | Every condition must be TRUE        |
| OR(...)                | At least one condition must be TRUE |
| NOT(test)              | Reverse TRUE and FALSE              |
| SWITCH(value,...)      | Map exact values to outputs         |

IFS returns the result for the **first TRUE condition**, so test the highest or most restrictive threshold first. SWITCH is clearer than nested IF statements when mapping known codes to labels.

Logical tests return TRUE or FALSE. Double unary converts them to 1 and 0, which lets SUMPRODUCT aggregate Boolean conditions; for ordinary criteria totals, SUMIFS is usually easier to audit.

> **Key idea:** Write the business rule in words beside the formula, including thresholds, units and what happens at the boundary.

## 10 Text cleaning and extraction

### Standardise labels before grouping matching or joining

| **Function**           | **Purpose**                    | **Example**             |
|------------------------|--------------------------------|-------------------------|
| TRIM                   | Remove extra ordinary spaces   | =TRIM(A2)               |
| CLEAN                  | Remove non-printing characters | =CLEAN(A2)              |
| UPPER / LOWER / PROPER | Standardise case               | =UPPER(A2)              |
| SUBSTITUTE             | Replace matching text          | =SUBSTITUTE(A2,"-"," ") |
| TEXTBEFORE / TEXTAFTER | Extract around a delimiter     | =TEXTBEFORE(A2,"-")     |
| LEFT / RIGHT / MID     | Extract by position            | =MID(A2,5,3)            |
| LEN                    | Count characters               | =LEN(A2)                |

    =TEXTSPLIT(A2,",")
    =TEXTJOIN(" | ",TRUE,B2:D2)
    =VALUE("1250.50")
    =NUMBERVALUE(A2,".",",")

TEXTSPLIT returns several cells. TEXTJOIN combines values and can ignore empty cells. VALUE follows local separators; NUMBERVALUE lets you state the decimal and thousands separators explicitly.

Cleaning joins requires the same transformation on both key columns. Check case, whitespace, punctuation, leading zeros and hidden characters before deciding that a key is missing.

> **Key idea:** Do not use formatted text as input to numeric analysis. Convert the stored value, then apply a display format.

## 11 Date and time foundations

### Understand serial values boundaries and display formats

Excel stores a date as a whole-number serial value and time as a fraction of one day. A date-time combines both. Formatting changes how the number looks; it does not change the stored value.

| **Value or formula** | **Meaning**                           |
|----------------------|---------------------------------------|
| =DATE(2026,3,1)      | Build 1 March 2026 from numeric parts |
| =TODAY()             | Current date; volatile                |
| =NOW()               | Current date and time; volatile       |
| =YEAR(A2)            | Extract the year                      |
| =MONTH(A2)           | Extract the month number              |
| =DAY(A2)             | Extract the day number                |
| =MOD(A2,1)           | Extract the time portion              |

    =SUMIFS(Sales[Revenue],Sales[OrderDate],">="&DATE(2026,1,1),
     Sales[OrderDate],"<"&DATE(2026,4,1))

This **half-open interval** includes 1 January and excludes 1 April. It works for both dates and timestamps because it avoids relying on a final second such as 31 March 23:59:59.

Imported text such as 03/04/2026 is ambiguous across locales. Confirm the source convention before conversion, and confirm the time zone when timestamps cross systems.

> **Key idea:** A date is a value and a date format is presentation. Keep those ideas separate.

## 12 Periods, workdays and calendar logic

### Build consistent reporting periods instead of manual labels

| **Need**      | **Formula**                      | **Idea**                             |
|---------------|----------------------------------|--------------------------------------|
| Month end     | =EOMONTH(A2,0)                   | Last date of the current month       |
| Shift months  | =EDATE(A2,3)                     | Same day three months later          |
| Working days  | =NETWORKDAYS(Start,End,Holidays) | Exclude weekends and listed holidays |
| Add workdays  | =WORKDAY(Start,10,Holidays)      | Move by ten working days             |
| Year fraction | =YEARFRAC(Start,End)             | Part of a year between dates         |
| Month label   | =TEXT(A2,"yyyy-mm")              | Readable label; the result is text   |

A reporting calendar often needs Year, Quarter, Month, YearMonth, Week and working-day flags. Store a real MonthStart date for sorting, then display a friendly label. Text labels alone can sort alphabetically instead of chronologically.

    MonthStart = DATE(YEAR(A2),MONTH(A2),1)
    Quarter = "Q"&ROUNDUP(MONTH(A2)/3,0)
    YearMonth = TEXT(A2,"yyyy-mm")

NETWORKDAYS and WORKDAY use weekend rules that may differ by organisation. Use the .INTL versions when weekends are not Saturday and Sunday.

> **Key idea:** Use one controlled calendar definition across formulas, PivotTables, charts and DAX.

## 13 Exact lookups with XLOOKUP

### Return an attribute by matching a unique key

    =XLOOKUP([@CustomerID],Customers[CustomerID],Customers[CustomerName],"Not found")

| **Argument**     | **Meaning**                                         |
|------------------|-----------------------------------------------------|
| lookup_value     | The key to find                                     |
| lookup_array     | The column containing the key                       |
| return_array     | The column or columns to return                     |
| \[if_not_found\] | Optional result when no match exists                |
| \[match_mode\]   | 0 exact; -1 next smaller; 1 next larger; 2 wildcard |
| \[search_mode\]  | Search direction or binary-search option            |

XLOOKUP uses exact match by default and can return values from either side of the key column. If duplicate keys exist, a normal search returns the first match; that is not proof that the relationship is one-to-one.

    =IFNA(XLOOKUP(A2,Products[ProductID],Products[Category]),"Missing key")
    =COUNTIF(Customers[CustomerID],A2)

IFNA handles only the no-match case and therefore masks fewer error types than IFERROR. COUNTIF checks how many times the key appears. Expected cardinality should be defined before the lookup.

> **Key idea:** Clean both key columns, test duplicates and count unmatched keys before trusting lookup results.

## 14 Approximate and two-way lookups

### Match thresholds or find an intersection deliberately

    =XLOOKUP(Score,Thresholds[Minimum],Thresholds[Grade],,-1)

Match mode **-1** asks for an exact match or the next smaller threshold. If a score is 82 and minimums are 0, 60, 70, 80 and 90, the formula returns the label attached to 80. Keep threshold rows sorted from smallest to largest and test exact boundary values.

| **Minimum** | **Grade** |
|-------------|-----------|
| 0           | F         |
| 60          | D         |
| 70          | C         |
| 80          | B         |
| 90          | A         |

    =INDEX(DataRange,
     XMATCH(RowKey,RowKeys,0),
     XMATCH(Header,Headers,0))

The first XMATCH finds the row position. The second finds the column position. INDEX returns the value at their intersection. Separating position from return logic makes the formula easier to test.

Binary search modes can be faster on very large ranges, but they require correctly sorted data. Incorrect order can return a plausible wrong answer instead of an error.

> **Key idea:** Approximate matching encodes a classification rule. Store and document the thresholds as data, not inside a deeply nested formula.

## 15 Conditional summaries

### Aggregate only rows that meet explicit criteria

    =SUMIFS(Sales[Revenue],Sales[Region],H2,Sales[Status],"Paid")
    =COUNTIFS(Sales[Region],H2,Sales[Status],"Paid")
    =AVERAGEIFS(Sales[Revenue],Sales[Region],H2,Sales[Status],"Paid")
    =MAXIFS(Sales[Revenue],Sales[Region],H2)

| **Function**    | **Output**                              |
|-----------------|-----------------------------------------|
| SUMIFS          | Total of values meeting every criterion |
| COUNTIFS        | Number of rows meeting every criterion  |
| AVERAGEIFS      | Mean of qualifying values               |
| MINIFS / MAXIFS | Smallest or largest qualifying value    |

SUMIFS starts with the sum range, followed by pairs of criteria range and criterion. Every criteria range must have the same shape. Criteria can be exact text, a cell reference, a comparison such as "\>="&H2, a non-blank test "\<\>", or a wildcard such as "North\*".

    =SUMIFS(Sales[Revenue],Sales[Region],"North")
     +SUMIFS(Sales[Revenue],Sales[Region],"West")

IFS functions apply **AND logic** across criteria. For simple OR logic, add separate results. For many alternatives, use a PivotTable, FILTER or a maintained mapping table rather than a long formula.

> **Key idea:** Conditional summaries are reliable only when the metric range and every criteria range refer to the same rows.

## 16 Dynamic arrays

### Return a changing result set from one formula

    =FILTER(Sales,Sales[Status]="Paid","No rows")
    =SORTBY(UNIQUE(Sales[Product]),UNIQUE(Sales[Product]),1)
    =TAKE(SORTBY(Sales,Sales[Revenue],-1),10)
    =CHOOSECOLS(FILTER(Sales,Sales[Region]=H2),1,3,7)

| **Function**    | **Returns**                              |
|-----------------|------------------------------------------|
| FILTER          | Rows meeting a Boolean condition         |
| SORT / SORTBY   | A sorted array                           |
| UNIQUE          | Distinct values or rows                  |
| SEQUENCE        | Generated numbers                        |
| TAKE / DROP     | Rows or columns from an array edge       |
| CHOOSECOLS      | Selected columns                         |
| HSTACK / VSTACK | Arrays joined horizontally or vertically |

A dynamic-array formula is entered once and **spills** into neighbouring cells. The spill operator \# refers to the complete result, such as J2#. Keep the spill area empty.

A spilled formula cannot be entered inside an Excel Table, but it can use structured references to a Table as its source. The result automatically changes when the source or criteria change.

> **Key idea:** Think of a dynamic-array formula as a small query that returns a table, not as a formula copied into many cells.

## 17 LET, LAMBDA and array helpers

### Name intermediate logic and reuse tested calculations

    =LET(
     revenue,Sales[Units]*Sales[UnitPrice],
     paid,FILTER(revenue,Sales[Status]="Paid"),
     SUM(paid)
    )

LET creates name-value pairs inside one formula, followed by the final calculation. It improves readability and can improve performance because a repeated expression can be calculated once and reused.

    =LAMBDA(new,old,IF(old=0,NA(),(new-old)/old))(C2,B2)

LAMBDA defines parameters and the calculation that uses them. The example immediately calls the function with C2 and B2. After testing, store it in Name Manager with a clear name such as GrowthRate.

| **Helper** | **Idea**                             | **Example**                         |
|------------|--------------------------------------|-------------------------------------|
| BYROW      | Apply a LAMBDA to each row           | =BYROW(B2:E10,LAMBDA(r,SUM(r)))     |
| BYCOL      | Apply a LAMBDA to each column        | =BYCOL(B2:E10,LAMBDA(c,AVERAGE(c))) |
| MAP        | Transform corresponding array values | =MAP(B2:B10,LAMBDA(x,x\*1.2))       |
| REDUCE     | Accumulate values into one result    | =REDUCE(0,B2:B10,LAMBDA(a,v,a+v))   |

> **Key idea:** Use LET to explain a complex formula from the inside. Use LAMBDA only after the underlying logic has been tested with edge cases.

## 18 Measures of centre and spread

### Describe a distribution instead of reporting only the average

| **Measure**        | **Excel**                  | **Interpretation**                            |
|--------------------|----------------------------|-----------------------------------------------|
| Mean               | =AVERAGE(range)            | Arithmetic centre; sensitive to outliers      |
| Median             | =MEDIAN(range)             | Middle sorted value; more robust to extremes  |
| Mode               | =MODE.SNGL(range)          | Most frequent value                           |
| Range              | =MAX(range)-MIN(range)     | Full spread; driven by extremes               |
| Standard deviation | =STDEV.S(range)            | Typical sample variation around the mean      |
| Variance           | =VAR.S(range)              | Squared spread                                |
| Percentile         | =PERCENTILE.INC(range,0.9) | Value at or below roughly 90% of observations |
| Quartile           | =QUARTILE.INC(range,1)     | First quartile                                |

Use the mean when total balance matters and the distribution is not dominated by extremes. Use the median when a typical observation is more useful, such as salary or house-price analysis.

    Coefficient of variation = STDEV.S(range) / AVERAGE(range)

The coefficient of variation expresses spread relative to the mean, but it is not meaningful when the mean is near zero or when values can legitimately cross zero.

> **Key idea:** Always pair a measure of centre with sample size and a measure of spread.

## 19 Ranking, correlation and outliers

### Find position association and unusual observations carefully

    =RANK.EQ([@Revenue],Sales[Revenue],0)
    =LARGE(Sales[Revenue],3)
    =CORREL(RevenueRange,UnitsRange)

RANK.EQ assigns position, with 0 meaning descending order. Tied values receive the same rank. LARGE returns the nth largest value. CORREL ranges from -1 to 1 and measures **linear association**, not causation.

    Q1 = QUARTILE.INC(range,1)
    Q3 = QUARTILE.INC(range,3)
    IQR = Q3 - Q1
    Lower fence = Q1 - 1.5 * IQR
    Upper fence = Q3 + 1.5 * IQR

The IQR rule flags values far outside the middle 50% of the distribution. It is a screening method, not an automatic deletion rule. An extreme value may be a data error, a rare valid event or the most important business case.

- Check source records, units, duplicates and currency before excluding an outlier.

- Compare results with and without extreme observations and report the sensitivity.

- Inspect a scatter plot before interpreting correlation; non-linear patterns can produce a weak coefficient.

> **Key idea:** An anomaly is a request for investigation, not a conclusion.

## 20 Data validation and conditional formatting

### Prevent invalid inputs and highlight meaningful exceptions

| **Validation type**    | **Example**                           |
|------------------------|---------------------------------------|
| List                   | Paid, Pending or Cancelled            |
| Whole number / Decimal | Units \>= 0; discount between 0 and 1 |
| Date / Time            | Values inside the reporting period    |
| Text length            | Fixed-length IDs                      |
| Custom formula         | =COUNTIF(\$A:\$A,A2)=1 for unique IDs |

Validation should explain the rule with an input message and stop or warn about invalid entries with an error alert. It reduces manual-entry errors but does not replace source-system controls.

    =$G2="Cancelled"
    =AND($H2<0,$A2<>"")
    =COUNTIF($A:$A,$A2)>1

Conditional-format formulas are written from the top-left cell of the **Applies to** range. Lock rows and columns intentionally. Review rule order and Stop If True when multiple rules can apply.

- Use colour scales for ordered magnitude, not nominal categories.

- Use data bars for compact comparisons and icon sets only when thresholds have defined meaning.

- Do not use colour as the only signal; retain labels, values or symbols.

> **Key idea:** Validation controls what may be entered. Conditional formatting changes only how a value appears.

## 21 PivotTable foundations

### Summarise a large Table by rearranging fields

**1.** Select a cell in the source Table and choose Insert \> PivotTable.

**2.** Place descriptive categories in Rows or Columns.

**3.** Place numeric fields in Values and confirm the aggregation.

**4.** Add Filters, slicers or a timeline when interaction is useful.

**5.** Refresh after the source changes.

| **Area** | **Role**           | **Example**    |
|----------|--------------------|----------------|
| Rows     | Primary grouping   | Region         |
| Columns  | Secondary grouping | Year           |
| Values   | Aggregated metric  | Sum of Revenue |
| Filters  | Whole-Pivot filter | Status         |

A PivotTable groups source rows by the selected dimensions and recalculates the chosen aggregation for each intersection. It does not alter the source data.

| **Symptom**                  | **Likely cause**                                       |
|------------------------------|--------------------------------------------------------|
| Count appears instead of Sum | Numbers may be stored as text                          |
| New rows are missing         | The source range did not expand or Refresh was not run |
| Dates do not group           | Invalid dates or blanks may exist                      |
| Distinct Count unavailable   | The source was not added to the Data Model             |

> **Key idea:** Treat every PivotTable value field as a metric definition: aggregation, filters, units and grain must all be correct.

## 22 Pivot calculations, grouping and slicers

### Control how a PivotTable computes and displays results

| **Feature**          | **Use**                                                    |
|----------------------|------------------------------------------------------------|
| Value Field Settings | Change Sum, Count, Average, Min, Max or Distinct Count     |
| Show Values As       | Percent of total, difference, running total, rank or index |
| Date grouping        | Aggregate dates into Years, Quarters or Months             |
| Slicer               | Filter categories with visible buttons                     |
| Timeline             | Filter a valid date field by period                        |
| Report Connections   | Connect one slicer to several compatible PivotTables       |

**Show Values As** changes the displayed calculation without changing the stored source. For percentage of total, define which total is the denominator: grand total, row total, column total or parent category.

Date grouping creates convenient Pivot groups, but a dedicated calendar table is more consistent when the same fiscal definitions must be reused in several reports or DAX measures.

A standard Pivot calculated field has limited logic. Measures in the Data Model use DAX, relationships and filter context, which makes them preferable for reusable business metrics.

- Use tabular layout and repeat item labels when the Pivot must read like a table.

- Turn off unnecessary subtotals and grand totals; retain totals that answer a question.

- Sort by the analytical value, not alphabetically, when ranking categories.

> **Key idea:** Reconcile Pivot totals against an independent formula or source-system total after every material change.

## 23 Charts for analytical questions

### Choose the chart from the relationship you need to show

| **Question**         | **Good starting chart**           | **Avoid**                       |
|----------------------|-----------------------------------|---------------------------------|
| Compare categories   | Horizontal bar sorted high to low | 3-D columns                     |
| Trend over time      | Line                              | Unsorted or irregular dates     |
| Part to whole        | 100% stacked bar                  | Many pie slices                 |
| Distribution         | Histogram or box plot             | A bar for every observation     |
| Relationship         | Scatter                           | Connecting unrelated categories |
| Actual versus target | Bar or combo chart                | Decorative gauges               |

- State the conclusion or question in the title.

- Use the correct aggregation, unit and comparison period.

- Start bar axes at zero; explain justified exceptions for line charts.

- Label selectively and remove legends that repeat obvious information.

- Use one highlight colour and muted context.

- Keep time intervals regular and comparison periods equivalent.

A PivotChart follows its PivotTable and connected slicers. Keep calculation logic in the source, PivotTable or Data Model instead of manually editing chart series.

> **Key idea:** A chart should answer one question. Put the takeaway near it and keep the metric definition available.

## 24 Dashboard construction

### Turn calculations into a controlled reporting interface

| **Zone** | **Contents**                                                 |
|----------|--------------------------------------------------------------|
| Top      | Title, reporting period, refresh date and three to five KPIs |
| Middle   | Primary trend and comparisons explaining its drivers         |
| Controls | Slicers, timeline and clear reset instructions               |
| Bottom   | Details, definitions, sources and caveats                    |

    ="Revenue — "&TEXT(MIN(DateRange),"mmm yyyy")&" to "&TEXT(MAX(DateRange),"mmm yyyy")
    =IF(Current>=Target,"On target","Below target")

Dynamic titles help the reader see the filtered period and state. A dashboard should expose its scope instead of forcing the user to infer it from visual changes.

- Use consistent units, decimal places and date formats.

- Align chart edges and spacing into a clear reading path.

- Separate assumptions, calculations and outputs.

- Show refresh date, source and material metric caveats.

- Test every slicer against every intended visual and provide a reset method.

Protect report sheets carefully while preserving slicer use and refresh behaviour. Keep the underlying calculation layer or named measures available for audit.

> **Key idea:** A dashboard is an interface to analytical logic, not a substitute for documented calculations.

## 25 What-if analysis and forecasting

### Change assumptions without rewriting the model

| **Tool**                | **Use**                                 | **Output**             |
|-------------------------|-----------------------------------------|------------------------|
| Goal Seek               | Find one input that reaches a target    | One solved input       |
| Scenario Manager        | Save named sets of inputs               | Scenario summary       |
| One-variable Data Table | Test many values for one input          | Sensitivity table      |
| Two-variable Data Table | Test pairs of assumptions               | Sensitivity matrix     |
| Solver add-in           | Optimise an objective under constraints | Best feasible decision |

    =FORECAST.LINEAR(NewX,KnownY,KnownX)
    =TREND(KnownY,KnownX,NewX)
    =FORECAST.ETS(TargetDate,Values,Timeline)

FORECAST.LINEAR assumes a linear relationship. FORECAST.ETS models a regular time series with repeating seasonality. Both return estimates, not facts; hold out recent periods and compare predictions with actual values.

- Place assumptions in clearly labelled cells with realistic bounds and units.

- Separate base, upside and downside scenarios.

- Record which inputs change and which outputs are monitored.

- Test whether relationships remain valid outside the observed range.

> **Key idea:** A scenario explains sensitivity to assumptions. It does not create evidence that the assumed future will occur.

## 26 Power Query workflow

### Import clean and refresh data through recorded steps

Power Query is Excel's repeatable data-preparation layer. It connects to sources, applies transformations in order and loads the result to a worksheet Table, the Data Model, both, or as a connection only.

| **Stage** | **Action**                                                      |
|-----------|-----------------------------------------------------------------|
| Get Data  | Connect to a Table, file, folder, database or supported service |
| Transform | Apply steps in the Power Query Editor                           |
| Combine   | Merge matching columns or append compatible rows                |
| Load      | Choose a worksheet Table, Data Model, both or connection only   |
| Refresh   | Re-read the source and rerun every step                         |

**1.** Keep an unchanged source query when it improves traceability.

**2.** Promote the correct header row and remove irrelevant rows or columns.

**3.** Set data types early and inspect conversion errors.

**4.** Trim and standardise labels; define the key before removing duplicates.

**5.** Reshape, merge, append or group as required, then choose the load destination.

Refresh reruns the recorded logic; it does not automatically make the logic correct for a changed source. Column renames, moved files and new invalid values can break or alter results.

> **Key idea:** Power Query changes its output, not the original source. Compare row counts and totals before and after major steps.

## 27 Power Query cleaning and combining

### Use the transformation that matches the required shape

| **Operation**         | **Shape change**                    | **Typical use**                   |
|-----------------------|-------------------------------------|-----------------------------------|
| Filter rows           | Remove observations                 | Keep paid orders                  |
| Choose columns        | Remove fields                       | Reduce data size                  |
| Split / merge columns | Change field structure              | Parse or assemble labels          |
| Group By              | Many rows become summaries          | Revenue by region                 |
| Merge Queries         | Add matching columns                | Join Sales to Products            |
| Append Queries        | Stack compatible rows               | Combine monthly files             |
| Pivot                 | Category values become columns      | Create status columns             |
| Unpivot               | Columns become attribute-value rows | Convert wide months to tidy dates |

A merge is conceptually like a database join. A left outer merge keeps every row from the first query and adds matches from the second. Duplicate keys on the second side can multiply rows.

- Make key data types compatible before merging.

- Filter rows and remove unused columns early when safe.

- Preserve query folding for supported sources so work can be pushed back to the source system.

- Disable load for staging queries that only feed another query.

- Parameterise source paths and dates that must change predictably.

> **Key idea:** After every merge, compare row counts, unmatched keys and totals to detect missing or multiplied records.

## 28 M language structure

### Read the code generated by Power Query

    let
        Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
        Typed = Table.TransformColumnTypes(Source,{
            {"OrderDate", type date},
            {"Units", Int64.Type},
            {"UnitPrice", Currency.Type}}),
        AddedRevenue = Table.AddColumn(Typed,"Revenue",
            each [Units] * [UnitPrice], Currency.Type)
    in
        AddedRevenue

| **Element**       | **Meaning**                           |
|-------------------|---------------------------------------|
| let               | Starts a set of named expressions     |
| Step = expression | Binds a step name to a value          |
| comma             | Separates steps                       |
| in                | Returns the selected final expression |
| each              | Shorthand for a one-argument function |
| \[Column\]        | Current row field inside each         |
| {...}             | A list or positional item access      |
| \[...\]           | A record or field access              |

M is case-sensitive. Each step returns a new value; later steps refer to earlier results. Renaming steps improves readability. A step name containing spaces is written as \#"Step Name".

> **Key idea:** Read an M query from Source to in: each step describes one transformation of the previous result.

## 29 M transformations, functions and errors

### Write small explicit steps and treat null deliberately

    Paid = Table.SelectRows(Source, each [Status] = "Paid"),
    Selected = Table.SelectColumns(Paid,{"OrderID","Region","Revenue"}),
    Grouped = Table.Group(Selected,{"Region"},
        {{"Revenue", each List.Sum([Revenue]), Currency.Type}})

Table.SelectRows keeps rows for which the function after each returns true. Table.SelectColumns chooses fields. Table.Group creates one row per Region and aggregates the nested values for that group.

    SafeAmount = try Number.FromText([Amount]) otherwise null
    Custom = (units as number, price as number) as number => units * price

try ... otherwise returns a fallback when evaluation fails. Use null only when it correctly means no value. Null is different from an empty text string and from zero, so replacements must follow a defined missing-data policy.

| **Type**               | **Example** | **Why type matters**        |
|------------------------|-------------|-----------------------------|
| number / Currency.Type | 1250.50     | Arithmetic and aggregation  |
| Int64.Type             | 12          | Whole-number keys or counts |
| text                   | 00125       | Labels and leading zeros    |
| date / datetime        | 2026-03-01  | Calendar operations         |
| logical                | true        | Filters and flags           |

> **Key idea:** Handle errors near the step that creates them, and keep enough information to distinguish missing, invalid and corrected values.

## 30 The Data Model and star schema

### Relate tables instead of flattening everything into one sheet

The Excel Data Model stores several related tables inside the workbook. Power Query prepares the tables; relationships connect them; Power Pivot adds advanced modelling and DAX calculations; PivotTables and charts present the results.

| **Table**   | **Grain**              | **Typical fields**                               |
|-------------|------------------------|--------------------------------------------------|
| FactSales   | One row per order line | DateKey, ProductKey, CustomerKey, Units, Revenue |
| DimDate     | One row per date       | Date, Year, Quarter, Month                       |
| DimProduct  | One row per product    | ProductKey, Product, Category                    |
| DimCustomer | One row per customer   | CustomerKey, Customer, Region                    |

A **fact table** records events at one grain and contains numeric values plus foreign keys. A **dimension table** contains descriptive attributes used to group and filter those events. This arrangement is a star schema.

- Use stable keys instead of display names.

- Keep one consistent grain per table.

- Remove unused high-cardinality text columns from large fact tables.

- Hide technical keys from report users and format measures clearly.

> **Key idea:** The Data Model separates event storage from descriptive context so one clean dimension can filter many analyses.

## 31 Relationships and the calendar table

### Control how filters travel through the model

| **Side**            | **Requirement**                      | **Example**                              |
|---------------------|--------------------------------------|------------------------------------------|
| One side            | A unique key with one row per entity | DimProduct\[ProductKey\]                 |
| Many side           | A repeated foreign key               | FactSales\[ProductKey\]                  |
| Direction           | Normally dimension filters fact      | DimProduct → FactSales                   |
| Active relationship | Default path used by calculations    | DimDate\[Date\] → FactSales\[OrderDate\] |

A one-to-many relationship does not repair duplicate dimension keys or unmatched fact keys. Blank or unknown members can appear when fact rows do not have a corresponding dimension row.

A complete calendar table has one unique row per date across the required period, with no gaps. Add reporting attributes such as Year, Quarter, Month Number, Month Name, Month Start and fiscal fields. Sort Month Name by Month Number.

    DimDate = CALENDAR(DATE(2024,1,1),DATE(2026,12,31))
    Year = YEAR(DimDate[Date])
    Month Number = MONTH(DimDate[Date])
    Month = FORMAT(DimDate[Date],"mmm")

Use inactive relationships only when one fact table has several date roles, such as OrderDate and ShipDate. Activate the alternative path inside a measure with USERELATIONSHIP when required.

> **Key idea:** Most model problems are key, grain or filter-path problems. Inspect those before changing DAX.

## 32 DAX measures, columns and context

### Choose where a calculation should be evaluated

| **Calculated column**              | **Measure**                                 |
|------------------------------------|---------------------------------------------|
| Calculated for every stored row    | Calculated when the report requests a value |
| Uses row context naturally         | Responds to filter context                  |
| Consumes model memory              | Usually compact and reusable                |
| Useful for row categories and keys | Preferred for KPIs and aggregations         |

    Total Revenue :=
    SUMX(FactSales,FactSales[Units]*FactSales[UnitPrice])

    Order Count := DISTINCTCOUNT(FactSales[OrderID])

    Average Order Value := DIVIDE([Total Revenue],[Order Count])

**Row context** means the current row. It exists in calculated columns and iterator functions such as SUMX. **Filter context** is the set of filters coming from Pivot rows, columns, slicers, relationships and DAX expressions.

SUM adds one column. SUMX iterates a table, evaluates an expression for each row and then sums the results. DIVIDE handles zero or blank denominators more safely than /.

> **Key idea:** A measure is evaluated for the current question. The same measure changes when the PivotTable or slicers change.

## 33 CALCULATE and filter context

### Change the population behind a measure

    Paid Revenue :=
    CALCULATE([Total Revenue],FactSales[Status]="Paid")

    Revenue % Total :=
    DIVIDE(
        [Paid Revenue],
        CALCULATE([Paid Revenue],REMOVEFILTERS(DimCustomer[Region]))
    )

CALCULATE evaluates an expression under modified filter context. The first measure adds Status = Paid. The second removes only the Region filter from the denominator, allowing each region to be compared with the all-region total while other filters remain.

| **Pattern**       | **Use**                                          |
|-------------------|--------------------------------------------------|
| Boolean filter    | A simple condition on one table column           |
| FILTER(table,...) | A row-by-row table condition                     |
| REMOVEFILTERS     | Remove selected filters                          |
| KEEPFILTERS       | Intersect a new filter with the existing filter  |
| USERELATIONSHIP   | Use an inactive relationship for one calculation |
| TREATAS           | Apply values as filters to unrelated columns     |

New CALCULATE filters normally replace existing filters on the same columns. KEEPFILTERS changes that behaviour to intersection. Context transition occurs when CALCULATE turns an existing row context into filter context.

> **Key idea:** When a DAX number looks wrong, state the expected filter context in words before rewriting the formula.

## 34 DAX variables and time intelligence

### Reuse measures across periods without duplicating business logic

    Revenue PY :=
    CALCULATE([Paid Revenue],SAMEPERIODLASTYEAR(DimDate[Date]))

    Revenue YoY % :=
    VAR CurrentRevenue = [Paid Revenue]
    VAR PriorRevenue = [Revenue PY]
    RETURN DIVIDE(CurrentRevenue-PriorRevenue,PriorRevenue)

    Revenue YTD := TOTALYTD([Paid Revenue],DimDate[Date])

Time-intelligence functions shift or expand the date filter applied to an existing measure. Reusing \[Paid Revenue\] keeps its definition in one place. VAR names intermediate values and RETURN supplies the final result.

| **Function**       | **Idea**                                        |
|--------------------|-------------------------------------------------|
| SAMEPERIODLASTYEAR | The equivalent dates one year earlier           |
| DATEADD            | Shift the current date set by a chosen interval |
| TOTALYTD           | Accumulate from year start to the current date  |
| DATESINPERIOD      | Return a rolling period of dates                |

These patterns require a proper date table, an active relationship and comparable periods. An incomplete current year should normally be compared with the same elapsed period in the prior year, not the prior full year.

Grand totals evaluate the measure under total filter context. They are not always the arithmetic sum of visible rows, especially for ratios, distinct counts and non-additive metrics.

> **Key idea:** Build time measures from a trusted base measure and a trusted calendar; do not repeat the full business formula in every period calculation.

## 35 One complete Excel analysis

### Connect the question, data, calculation, visual and validation

**Question:** Which regions contributed paid revenue in the first quarter of 2026?

| **Decision**   | **Definition**                                        |
|----------------|-------------------------------------------------------|
| Population     | Paid order lines from 1 January through 31 March 2026 |
| Grain          | One output row per region                             |
| Metric         | Sum of Units × UnitPrice                              |
| Denominator    | Paid revenue across the same period                   |
| Missing policy | Keep blank or unmatched regions as Unknown            |

**1.** Import Sales with Power Query; set date and numeric types; standardise Region.

**2.** Replace missing Region with Unknown; add Revenue; filter Paid and the half-open date range.

**3.** Load the fact and dimensions to the Data Model and verify relationships.

**4.** Create Paid Revenue and Revenue % Total measures.

**5.** Build a PivotTable by Region, sort high to low and add a horizontal bar chart.

**6.** Explain the unknown share and reconcile totals, counts and dates to the filtered source.

    Paid Revenue := CALCULATE(SUM(FactSales[Revenue]),FactSales[Status]="Paid")
    Revenue % Total := DIVIDE([Paid Revenue],CALCULATE([Paid Revenue],REMOVEFILTERS(DimCustomer[Region])))

| **Region** | **Paid Revenue** | **Share**  |
|------------|------------------|------------|
| North      | 400,000          | 50.0%      |
| West       | 240,000          | 30.0%      |
| Unknown    | 160,000          | 20.0%      |
| **Total**  | **800,000**      | **100.0%** |

> **Key idea:** A complete analysis links every displayed number back to a defined population, grain, metric, denominator and source.

## 36 Reliability, performance and quick reference

### Build workbooks another analyst can refresh inspect and trust

| **Risk**                  | **Control**                                     |
|---------------------------|-------------------------------------------------|
| Wrong grain or duplicates | Define keys; compare row and distinct counts    |
| Missing values            | Count blanks and document their treatment       |
| Broken joins or lookups   | Check unmatched and duplicate keys              |
| Stale outputs             | Show refresh date and use Refresh All           |
| Hidden constants          | Use labelled assumptions or parameters          |
| Formula drift             | Use Tables; inspect first, middle and last rows |
| Wrong totals              | Reconcile with an independent calculation       |
| Version mismatch          | Document required Excel features                |

| **Need**                     | **Use**                                    |
|------------------------------|--------------------------------------------|
| Clean and reshape repeatably | Power Query                                |
| Relate large tables          | Data Model and Power Pivot                 |
| Reusable filter-aware metric | DAX measure                                |
| Quick grouped summary        | PivotTable                                 |
| Row-level worksheet logic    | Excel formula or Table column              |
| Interactive decision view    | PivotCharts, slicers and a dashboard sheet |

**Performance:** avoid whole-column array calculations when bounded ranges exist; remove unused model columns; prefer Power Query to thousands of copied cleaning formulas; limit volatile functions such as OFFSET, INDIRECT, NOW and TODAY when unnecessary.

### Official references

- [Excel help and learning](https://support.microsoft.com/en-us/excel/)
- [Power Query M function reference](https://learn.microsoft.com/en-us/powerquery-m/power-query-m-function-reference)
- [DAX function reference](https://learn.microsoft.com/en-us/dax/dax-function-reference)

> **Key idea:** Before sharing: refresh, recalculate, inspect errors, test filters, reconcile totals, record the reporting period and save a clean version.

---

## Visual cheatsheets

![Excel Cheatsheet 01](download_notes/Excel_Cheatsheet_01.png)

![Excel Cheatsheet 02](download_notes/Excel_Cheatsheet_02.png)

![Excel Cheatsheet 03](download_notes/Excel_Cheatsheet_03.png)

![Excel Cheatsheet 04](download_notes/Excel_Cheatsheet_04.png)

![Excel Cheatsheet 05](download_notes/Excel_Cheatsheet_05.png)

![Excel Cheatsheet 06](download_notes/Excel_Cheatsheet_06.png)

![Excel Cheatsheet 07](download_notes/Excel_Cheatsheet_07.png)

![Excel Cheatsheet 08](download_notes/Excel_Cheatsheet_08.png)

![Excel Cheatsheet 09](download_notes/Excel_Cheatsheet_09.png)

![Excel Cheatsheet 10](download_notes/Excel_Cheatsheet_10.png)

![Excel Cheatsheet 11](download_notes/Excel_Cheatsheet_11.png)

![Excel Cheatsheet 12](download_notes/Excel_Cheatsheet_12.png)

![Excel Cheatsheet 13](download_notes/Excel_Cheatsheet_13.png)

![Excel Cheatsheet 14](download_notes/Excel_Cheatsheet_14.png)

![Excel Cheatsheet 15](download_notes/Excel_Cheatsheet_15.png)

![Excel Cheatsheet 16](download_notes/Excel_Cheatsheet_16.png)

![Excel Cheatsheet 17](download_notes/Excel_Cheatsheet_17.png)

![Excel Cheatsheet 18](download_notes/Excel_Cheatsheet_18.png)
