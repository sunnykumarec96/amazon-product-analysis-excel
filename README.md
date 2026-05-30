# Amazon Product Analysis using Excel

## Project Overview

This project analyzes Amazon product data using Excel.

## Skills Used

* Data Cleaning
* COUNTIF
* SUMIF
* AVERAGEIF
* LEFT
* FIND
* SUBSTITUTE
* VALUE

## Data Cleaning Steps

1. Extracted Main Category from category hierarchy.
2. Converted discounted price from text to numeric format.
3. Prepared clean columns for analysis.

## Key Findings

### Electronics Category

* Product Count: 526
* Total Discounted Price: ₹31,38,057
* Average Discounted Price: ₹5,965.89

## Excel Formulas Used

### Main Category

=LEFT(C2,FIND("|",C2)-1)

### Clean Price

=VALUE(SUBSTITUTE(SUBSTITUTE(E2,"â‚¹",""),",",""))

### Product Count

=COUNTIF(D:D,"Electronics")

### Total Price

=SUMIF(D:D,"Electronics",F:F)

### Average Price

=AVERAGEIF(D:D,"Electronics",F:F)

## Tools

* Microsoft Excel
* GitHub
