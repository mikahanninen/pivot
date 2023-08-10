*** Comments ***
1. Create a pivot table
2. Remove duplicate values
3. Adding more columns to the pivot table
4. create a formula in that column
5. filter the pivot table
6. pretty complex pivot table usage


*** Settings ***
Library     ExtendedExcel    autoexit=${FALSE}    WITH NAME    excel
Library     utils.py


*** Tasks ***
Minimal task
    Create Test Excel File    ${CURDIR}${/}pivoting.xlsx    data
    excel.Open Application    visible=${TRUE}
    Open Workbook    ${CURDIR}${/}pivoting.xlsx
    Add New Sheet    test
    # Write To Cells    row=1    column=1    value=abc
    @{pt_rows}=    Create List    products
    @{pt_cols}=    Create List    expense    date    # products
    @{pt_filters}=    Create List    @{EMPTY}
    @{pt_fields}=    Create List    @{EMPTY}    # price
    Create Pivot Table    data    test    pivoting    ${pt_rows}    ${pt_cols}    ${pt_filters}    ${pt_fields}
    Save Excel
    ${tables}=    Get Pivot Tables
    Log To Console    ${tables}
    Log    Done.
