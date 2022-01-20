*** Settings ***
Resource    load_components.resource

*** Variables ***

*** Keywords ***


[Documentation]    _*SUMMARY:*_
    ...    Open, read and write files in excel format
    ...
    ...    _*PRECONDITION:*_
    ...   - Library: Excel.py
    ...   - Installation Procedure:
    ...   - "Put in the Libraries folder the Excel.py"
    ...   - Install: pip install openpyxl
    ...   - Install: pip install natsort
    ...   - Install: pip install xlrd
    ...   Use Example:
    ...
    ...    _*PRECONDITION:*_
    ...   Mantainer:

Leer archivo excel
    [Arguments]   ${TC_sheet}
    Check input data row count    ${RUTA_DATA_EXCEL}    ${TEST_NAME}    ${TC_sheet}
    FOR    ${index}    IN RANGE    1    ${UsedRow}
        Read excel test    ${RUTA_DATA_EXCEL}    ${TC_sheet}    ${index}
    END

Check input data row count
    [Arguments]    ${file}    ${testname}    ${sheet_name}
    Open Excel Document     ${file}     ${testname}    #Open Excel File
    ${UsedRow}=    Count Row    ${sheet_name}    #Get used row count
    Set Suite Variable    ${UsedRow}
    Log to Console    Input Data Row Count: ${UsedRow}    #Log row count to console
    Save Excel Document     ${file}    #Save Excel Document
    Close Current Excel Document

Read excel test
    [arguments]   ${filename}    ${sheet}   ${row}
    ${Data}   Load Data Table     ${filename}     ${sheet}     ${row}
    Set Test Variable    ${_DATA}
