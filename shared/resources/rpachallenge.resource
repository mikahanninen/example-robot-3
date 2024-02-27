*** Settings ***
Documentation       Robot to solve the first challenge at rpachallenge.com, which
...                 consists of filling a form that randomly rearranges itself,
...                 with data taken from a provided Microsoft Excel file.

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
Library             RPA.HTTP


*** Keywords ***
Get the list of people from the Excel file
    [Arguments]    ${excel_file}=challenge.xlsx
    Open Workbook    ${excel_file}
    ${table}=    Read Worksheet As Table    header=True
    Close Workbook
    RETURN    ${table}

Set value by XPath
    [Arguments]    ${xpath}    ${value}
    ${result}=
    ...    Execute Javascript
    ...    document.evaluate('${xpath}',document.body,null,9,null).singleNodeValue.value='${value}';
    RETURN    ${result}

Fill and submit the form
    [Arguments]    ${person}
    Set Value By XPath    //input[@ng-reflect-name="labelFirstName"]    ${person}[First Name]
    Set Value By XPath    //input[@ng-reflect-name="labelLastName"]    ${person}[Last Name]
    Set Value By XPath    //input[@ng-reflect-name="labelCompanyName"]    ${person}[Company Name]
    Set Value By XPath    //input[@ng-reflect-name="labelRole"]    ${person}[Role in Company]
    Set Value By XPath    //input[@ng-reflect-name="labelAddress"]    ${person}[Address]
    Set Value By XPath    //input[@ng-reflect-name="labelEmail"]    ${person}[Email]
    Set Value By XPath    //input[@ng-reflect-name="labelPhone"]    ${person}[Phone Number]
    Click Button    Submit
