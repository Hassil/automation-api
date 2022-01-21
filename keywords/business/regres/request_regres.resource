*** Keywords ***

Get con parametros
    Create Session   ${SESSION_NAME_REGRES}   ${APP_BASE_URL_REGRES}
    &{param}=   Create Dictionary    page=2
    ${response}=   Get Request   ${SESSION_NAME_REGRES}   /api/users   params=${param}
    
    ${statuscode}=  Convert To String    ${response.status_code}
    Should Be Equal    ${statuscode}    200
    Log To Console   ${response.content}
    ${jsonresponse}=   to json   ${response.content}
    @{name_list}=   get value from json   ${jsonresponse}   data[0].first_name
    ${name}=   get from list   ${name_list}   0
    Should Be Equal    ${name}    Michael
