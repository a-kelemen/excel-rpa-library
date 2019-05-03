*** Settings ***
Library                      ExcelProcessLibrary

*** Test Cases ***

Open Workbook
    [Tags]                   Open Workbook
    Open Workbook            //sample_workbook.xlsx  sample

Save Workbook As
    [Tags]                   Save Workbook As
    Save Workbook As         test_wb.xlsx  sample

Close Workbook
    [Tags]                   Close Workbook
    Close Workbook           sample

Create Worksheet
    [Tags]                   Create Worksheet  Open Workbook
    Open Workbook            //test_wb.xlsx
    Create Worksheet         test_sheet_not_renamed

Delete Worksheet
    [Tags]                   Delete Worksheet  Create Worksheet  Set Active Worksheet
    Create Worksheet         delete_sheet  workbook=default_
    Delete Worksheet         delete_sheet  workbook=default_
    Run Keyword And Expect Error     WorksheetNotFoundException*
        ...  Set Active Worksheet  delete_sheet

Rename Worksheet
    [Tags]                   Rename Worksheet
    Rename Worksheet         test_sheet_not_renamed  test_sheet

Set Active Worksheet
    [Tags]                   Set Active Worksheet
    Set Active Worksheet     test_sheet

Set Cell Value
    [Tags]                   Set Cell Value
    Set Cell Value           A2  string:
    Set Cell Value           B2  spagetti
    Set Cell Value           A3  number:
    Set Cell Value           B3  62
    Set Cell Value           A4  number:
    Set Cell Value           B4  0.15
    Set Cell Value           A5  date:
    Set Cell Value           B5  2000.01.01

Set Cell Value With Types
    [Tags]                   Set Cell Value
    Set Cell Value           C2  spagetti  cell_type=string
    Set Cell Value           C3  62  cell_type=number
    Set Cell Value           C4  0.15  cell_type=number
    Set Cell Value           C5  2000.01.01  cell_type=date

Add Formula To Cell
    [Tags]                   Add Formula To Cell  Set Cell Value
    Set Cell Value           A6  formula:
    Add Formula To Cell      B6  =SUM(B3:D4)

Get Cell Value String
    [Tags]                   Get Cell Value
    ${val}=                  Get Cell Value  C2
    Should Be Equal          ${val}  spagetti

Get Cell Value Int
    [Tags]                   Get Cell Value
    ${val}=                  Get Cell Value  C3
    Should Be Equal          ${val}  ${62}

Get Cell Value Float
    [Tags]                   Get Cell Value
    ${val}=                  Get Cell Value  C4
    Should Be Equal          ${val}  ${0.15}

Get Cell Value Date
    [Tags]                   Get Cell Value
    ${val}=                  Get Cell Value  C5  selected_sheet=active_
    Should Be Equal As Strings  ${val}  2000-01-01 00:00:00

Get Cell Value Formula
    [Tags]                   Get Cell Value
    ${val}=                  Get Cell Value  B6
    Should Be Equal As Numbers  ${val}  ${124.3}

Add List To Column
    [Tags]                   Add List To Column  Get Cell Value
    #${list}=                 Create List  ${1}  ${2}  ${3}
    ${list}=                 Create List  a  b  c
    Add List To Column       A8  ${list}
    ${val}=                  Get Cell Value  A10
    Should Be Equal          ${val}  c

Add List To Column Numbers
    [Tags]                   Add List To Column  Get Cell Value
    ${list}=                 Create List  ${1}  ${2}  ${3}
    Add List To Column       B8  ${list}
    ${val}=                  Get Cell Value  B10
    Should Be Equal As Numbers  ${val}  ${3}

Add List To Row
    [Tags]                   Add List To Row  Get Cell Value
    ${list}=                 Create List  a  b  c
    Add List To Row          A12  ${list}
    ${val}=                  Get Cell Value  C12  workbook=default_
    Should Be Equal          ${val}  c

Add List To Row With Types
    [Tags]                   Add List To Row  Get Cell Value
    ${list}=                 Create List  1  2  3
    Add List To Row          A13  ${list}  cell_type=number
    ${val}=                  Get Cell Value  C13
    Should Be Equal As Numbers  ${val}  ${3}

Get Cell Value Integer
    [Tags]                   Get Cell Value  Set Cell Value
    Set Cell Value           E2  ${100}
    ${val}=                  Get Cell Value  E2
    Should Be Equal As Numbers  ${val}  100

Copy Cell
    [Tags]                   Copy Cell  Get Cell Value
    Copy Cell                E2  E3
    ${val}=                  Get Cell Value  E2
    Should Be Equal As Numbers  ${val}  100

Copy Cell From Worksheet
    [Tags]                   Copy Cell From Worksheet  Get Cell Value
    Copy Cell From Worksheet  A1  A1  sample  test_sheet
    ${val}=                  Get Cell Value  A1
    Should Be Equal          ${val}  test_sheet

Clear Cell
    [Tags]                   Clear Cell  Get Cell Value
    Clear Cell               E2
    ${val}=                  Get Cell Value  E2  selected_sheet=active_  workbook=default_
    Should Be Equal          ${val}  ${None}

Get First Empty Cell Below
    [Tags]                   Get First Empty Cell Below
    ${cell}=                 Get First Empty Cell Below  A1
    Should Be Equal          ${cell}  A7

Get First Empty Cell Right To
    [Tags]                   Get First Empty Cell Right To
    ${cell}=                 Get First Empty Cell Right To  A1
    Should Be Equal          ${cell}  B1

Get Cell Right To
    [Tags]                   Get Cell Right To
    ${cell}=                 Get Cell Right To  A1
    Should Be Equal          ${cell}  B1

Get Cell Left To
    [Tags]                   Get Cell Left To
    ${cell}=                 Get Cell Left To  B1
    Should Be Equal          ${cell}  A1

Get Cell Above
    [Tags]                   Get Cell Above
    ${cell}=                 Get Cell Above  A2
    Should Be Equal          ${cell}  A1

Get Cell Below
    [Tags]                   Get Cell Below
    ${cell}=                 Get Cell Below  A1
    Should Be Equal          ${cell}  A2

Set Cell Background
    [Tags]                   Set Cell Background  Set Cell Value
    Set Cell Value           F2  Blue Background
    Set Cell Background      F2  BLUE

Set Font Color
    [Tags]                   Set Font Color  Set Cell Value
    Set Cell Value           F3  Yellow Font
    Set Font Color           F3  YELLOW

Set Cell Style
    [Tags]                   Set Cell Style  Set Cell Value
    Set Cell Value           F4  italic
    Set Cell Style           F4  italic
    Set Cell Value           F5  bold
    Set Cell Style           F5  bold  selected_sheet=test_sheet
    Set Cell Value           F6  underline
    Set Cell Style           F6  underline  workbook=default_
    Set Cell Value           F7  normal
    Set Cell Style           F7  normal  selected_sheet=active_

Set Font Size
    [Tags]                   Set Font Size  Set Cell Value
    Set Cell Value           F8  8pt size
    Set Font Size            F8  8

Remove Style
    [Tags]                   Remove Style  Set Cell Value
    Set Cell Value           F9  removed style
    Set Cell Style           F9  italic
    Set Cell Style           F9  bold
    Remove Style             F9

Multiple Styles
    [Tags]                   Set Font Size  Set Cell Style  Set Font Color  Set Cell Background  Set Cell Value
    Set Cell Value           F10  8pt red yellow bold italic
    Set Font Color           F10  RED
    Set Cell Background      F10  YELLOW
    Set Cell Style           F10  italic
    Set Cell Style           F10  bold

Close All Workbooks
    [Tags]                   Close All Workbooks
    Close All Workbooks
    Run Keyword And Expect Error     WorkbookNotFoundException*
        ...  Get Cell Value  A1