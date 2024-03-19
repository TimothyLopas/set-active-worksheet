*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Excel.Application


*** Variables ***
${ActiveFilePath}=      ActiveWorksheetTest.xlsx


*** Tasks ***
Minimal task
    Open Application
    # This will open the workbook to the last saved worksheet.
    # In our case Sheet1
    Open Workbook    filename=${ActiveFilePath}
    FOR    ${counter}    IN RANGE    1    5
        ${activeSheet}=    Set Variable    Sheet${counter}
        Set Active Worksheet    ${activeSheet}
        ${CellValue}=    Read From Cells    row=1    column=1
        ${currentWorksheet}=    Active worksheet    ${CellValue}
    END


*** Keywords ***
Active worksheet
    [Arguments]    ${value}
    IF    "${value}" == "Test1"
        RETURN    Sheet1
    ELSE IF    "${value}" == "Test2"
        RETURN    Sheet2
    ELSE IF    "${value}" == "Test3"
        RETURN    Sheet3
    ELSE IF    "${value}" == "Test4"
        RETURN    Sheet4
    ELSE
        RETURN    ERROR
    END
