*** Settings ***
Documentation       Template robot file to Process Excel files

Library             RPA.Email.ImapSmtp    smtp_server=smtp.gmail.com    smtp_port=587
Library             RPA.Archive
Library             DateTime
Library             BuiltIn
Library             RPA.FileSystem
Library             RPA.Excel.Application
Library             Collections


*** Variables ***
${FOLDER_PATH}          C:\\Users\\lateef_p\\Documents\\Robots\\Remove_Blank_RowsAndCol\\Src_Files\\
${KEYWORD}              Test Excel Files
${USERNAME}             testdemorpa@gmail.com
${PASSWORD}             btsarmyxctswfgvugmjbfbs
${RECIPIENT}      		lateef_p@trigent.com
${source_path}          C:\\Users\\lateef_p\\Documents\\Robots\\Remove_Blank_RowsAndCol\\Master\\Master_File.xlsm
${Email_Recipient}      lateef_p@trigent.com

*** Tasks ***
Performing Excel Operations On Rows and Columns
    Download Attachments From Gmail
    Check Files and Send Email for Non-Excels
    ${status_Of_Macro}=    Macro Execution
    Run Keyword If    '${status_Of_Macro}' == 'PASS'    Send Email Success
    ...    ELSE    Send Failure Email



*** Keywords ***
Check Files and Send Email for Non-Excels
    ${files} =    List Files In Directory    ${FOLDER_PATH}
    ${non_excel_files} =    Evaluate    [file.name for file in $files if not str(file).endswith(('.xlsx', '.xls'))]
    Run Keyword If    ${non_excel_files} != []
    ...    Send Email for other Extensions    ${non_excel_files}

*** Keywords ***
Send Email for other Extensions    [Arguments]    ${non_excel_files}
    Authorize    account=${USERNAME}    password=${PASSWORD}
    ${subject} =    Set Variable    Files found
    ${body} =    Set Variable    The following non-Excel files were found: ${non_excel_files}
    Send Message    sender=${USERNAME}    recipients=${RECIPIENT}    subject=${subject}    body=${body}


*** Keywords ***
Download Attachments From Gmail
    #${credentials}=    Get Secret    gmail
    Authorize   ${USERNAME}    ${PASSWORD}
    @{messages}=    List Messages    SUBJECT "${KEYWORD}"
    FOR    ${message}    IN    @{messages}
        Log    ${message}
        IF    ${message}[Has-Attachments] == True
            Save Attachment    ${message}    ${CURDIR}${/}Src_Files    True
        END
    END
*** Keywords ***
Macro Execution
    Open Application    ${TRUE}
    Open Workbook    ${source_path}
    Run Macro    RemoveBlankRowAndCol
    Save Excel
    Quit Application
    #Set Test Variable    ${macro_execution_status}    ${TRUE}
    RETURN    PASS
    
*** Keywords ***
Send Email Success
    Authorize    ${Email_Username}    ${Email_Password}
    Send Message
    ...    ${Email_Username}
    ...    ${Email_Recipient}
    ...    Files Found and Excel Modification SuccessFully Done..

