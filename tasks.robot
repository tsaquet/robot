*** Settings ***
Library     RPA.Browser.Selenium    auto_close=${FALSE}
Library     RPA.Email.ImapSmtp
Library     RPA.Robocorp.Vault
Library     RPA.FileSystem
Library     Collections
Library     RPA.Excel.Files
Library     RPA.Tables


*** Variables ***
${DOWNLOAD_DIR}     ${CURDIR}${/}downl${/}


*** Tasks ***
Open browser navigateur
    Set Download Directory    ${DOWNLOAD_DIR}
    ${provider}=    Get Secret    provider
    Open Available Browser    ${provider}[url1]
    Sleep    1s
    Click Button When Visible    //*[@id="tarteaucitronAlertBig"]/button[2]
    Sleep    1s
    Input Text    alias:Username    ${provider}[username]
    Input Text    alias:Password    ${provider}[pwd]
    Click Element    alias:Mb2sztypobutton8
    Wait Until Element Is Visible    alias:Dflexnthchild2linkicondlgnone    30 seconds
    Go To    ${provider}[url2]
    Wait Until Element Is Visible    alias:Containervolumechartheader
    Click Button    css:#menu-contract-selector
    Click Element    //*[contains(text(),'Eau Chatou')]
    Wait Until Element Is Visible    alias:Downloadlink
    Sleep    2s
    
    Click Element    //u[contains(text(), '${provider}[button_title]')]/..
    #Click Element
    Sleep    5s
    ${files}=    List Files In Directory    ${DOWNLOAD_DIR}
    ${lastModifiedFile}=    Get From List    ${files}    0
    Log    ${lastModifiedFile}
    ${time1}=    Get File Modified Date    ${lastModifiedFile}
    Log    ${time1}
    FOR    ${file}    IN    @{files}
        ${time}=    Get File Modified Date    ${file}
        ${lastModifiedFile}=    Set Variable If    ${time1} < ${time}    ${file}    ${lastModifiedfile}
        ${time1}=    Set Variable If    ${time1} < ${time}    ${time}    ${time1}
    END
    
    ${mail}=    Get Secret    mail
    
    Authorize Smtp    ${mail}[sender]    ${mail}[pwd]    smtp.gmail.com    587
    ${fileToRead}=    Convert To String    ${lastModifiedFile}
    Log    ${fileToRead}
    Send Message
    ...    ${mail}[sender]
    ...    ${mail}[receiver]
    ...    ${mail}[title]
    ...    Le fichier en PJ
    ...    ${fileToRead}
    ${result}=    Create List
    Open Workbook    ${fileToRead}
    ${worksheet}=    Read Worksheet As Table    header=${TRUE}
    FOR    ${values}    IN    @{worksheet}
        IF    ${values}[M3] is not None
            ${quantity}=    Convert To Number    ${values}[M3]
            IF    ${quantity} >= 5    Append To List    ${result}    ${values}
        END
    END
    ${resultLength}=    Get Length  ${result}
    IF    ${resultLength} > 0
        Send Message
               ${mail}[sender]
        ...    ${mail}[receiver]
        ...    ${mail}[title]
        ...    ${result}
    END