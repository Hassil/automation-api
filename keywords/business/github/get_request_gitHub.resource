*** Keywords ***
Valida usuario gitHub
    Create Session  ${SESSION_NAME_GIT_HUB}    ${URL_BASE_GIT_HUB}
    ${response} =  Get Request  ${SESSION_NAME_GIT_HUB}  ${END_POINT_GIT_HUB}
    Should Be Equal As Strings  ${response.status_code}  200
    ${json} =  Set Variable  ${response.json()}
    Should Be Equal As Strings  ${json['login']}  robotframeworktutorial
    Log  ${json}
