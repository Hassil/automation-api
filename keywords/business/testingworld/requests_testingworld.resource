*** Keywords ***

Crear sesión
   RequestsLibrary.Create Session   ${SESSION_NAME_TESTING}   ${APP_BASE_URL_TESTING}
   BuiltIn.Set Test Variable    ${_STEP}    Se creo la sesión con: ${APP_BASE_URL_TESTING}
   Step    ${_STEP}

Búsqueda de información por id
    ${response}=   RequestsLibrary.Get Request   ${SESSION_NAME_TESTING}   ${END_POINT_STUDENT_DETAILS}${STUDENT_ID}
    BuiltIn.Set Test Variable    ${_STEP}    Se realizó la petición: ${response}
    Step    ${_STEP}

    ${actual_code}=   BuiltIn.Convert To String   ${response.status_code}
    BuiltIn.Should Be Equal    ${actual_code}    200
    BuiltIn.Set Test Variable    ${_STEP}    El estatus_code fue el correcto:${actual_code}
    Step    ${_STEP}

    ${json_res}=   RequestsLibrary.To json   ${response.content}
    BuiltIn.Set Test Variable    ${_STEP}    Se pasó a formato json el response: ${json_res}
    Step    ${_STEP}

    @{first_name_list}=   JSONLibrary.Get Value From Json   ${json_res}   data.first_name
    BuiltIn.Set Test Variable    ${_STEP}    Se creó una lista con el valor de first_name: ${first_name_list}
    Step    ${_STEP}

    ${first_name}=   Collections.Get From List   ${first_name_list}   0
    BuiltIn.Set Test Variable    ${_STEP}    Se obtuvo el primer first_name de la lista: ${first_name}
    Step    ${_STEP}

    BuiltIn.Should Be Equal    ${first_name}    b
    BuiltIn.Set Test Variable    ${_STEP}    El first_name concordó con el esperado
    Step    ${_STEP}

    @{last_name_list}=   JSONLibrary.Get Value From Json   ${json_res}   data.last_name
    BuiltIn.Set Test Variable    ${_STEP}    Se creó una lista con el valor de last_name: ${last_name_list}
    Step    ${_STEP}

    ${last_name}=   Collections.Get From List   ${last_name_list}   0
    BuiltIn.Set Test Variable    ${_STEP}    Se obtuvo el primer last_name de la lista: ${last_name}
    Step    ${_STEP}

    BuiltIn.Should Be Equal    ${last_name}    b
    BuiltIn.Set Test Variable    ${_STEP}    El last_name concordó con el esperado
    Step    ${_STEP}



TC_003_Validate_Delete_Request
    Create Session   AppAccess   ${App_Base_URL_regres}
    ${response}=   Delete Request   AppAccess   /api/users/10
    ${code}=  Convert To String    ${response.status_code}
    Should Be Equal    ${code}    200

    ${jsonresponse}=   to json   ${response.content}
    @{status_list}=   get value from json   ${jsonresponse}   status
    ${status}=   get from list   ${status_list}   0
    Should Be Equal    ${status}    true

TC_04_Create_new_resource
    Create Session   addData   ${App_Base_URL_regres}
    &{body}=   Create Dictionary   first_name=Hassil   middle_name=Rosio   last_name=World   date_of_birth=12/12/1990
    &{header}=   Create Dictionary    Content-Type=application/json
    ${response}=   Post Request   addData  api/studentsDetails/   data=${body}   headers=${header}
    ${code}=  Convert To String    ${response.status_code}
    Should Be Equal    ${code}    201
    Log To Console   ${response.content}

TC_05_update_data
    Create Session   updateData   ${App_Base_URL_regres}
    &{body}=   Create Dictionary   first_name=Hassil   middle_name=maria   last_name=World   date_of_birth=12/12/1990
    &{header}=   Create Dictionary    Content-Type=application/json
    ${response}=   Put Request   updateData  api/studentsDetails/205   data=${body}   headers=${header}
    ${code}=  Convert To String    ${response.status_code}
    Should Be Equal    ${code}    200
    og To Console   ${response.content}
    #| ${testscores} | {"scores": [{"id": 1, "value": "100"}, {"id": 2, "value": "77"}, {"id": 3, "value": "96"} ] } |
    #&{people}=    Convert string to JSON   {"People": [{"Name": "Mark"}, {"Name": "Jane"}]}

TC_06_End_to_end_testCase
    Create Session   e2e   ${App_Base_URL_testingworld}
    &{header}=   Create Dictionary    Content-Type=application/json
    &{body}=   Create Dictionary   first_name=${original_first_name}    middle_name=Rosio   last_name=World   date_of_birth=12/12/1990
    ${post_response}=   Post Request   e2e  api/studentsDetails/   headers=${header}   data=${body}
    ${json_response}=  To Json   ${post_response.content}
    @{id_list}=   Get Value From Json   ${json_response}   id
    Log To Console    Este es el ID_LIST:${id_list}
    ${id}=   Get From List   ${id_list}   0
    Log To Console    Este es el ID:${id}

    &{body1}=   Create Dictionary   id=${id}   first_name=${update_first_name}  middle_name=Rosio   last_name=World   date_of_birth=12/12/1990
    ${put_response}=   Put Request   e2e   api/studentsDetails/${id}    headers=${header}   data=${body1}
    Log To Console   ${put_response.content}
    Log To Console   ${put_response.status_code}

Fetch details and validate   ${id}
      ${delete_response}=   Delete Request   e2e   api/studentsDetails/${id}
      ${json_delete}=   To Json   ${delete_response.content}
      @{Message}=   Get Value From Json   ${json_delete}   status
      ${statusM}=   Get From List    ${Message}   0
      Should Be Equal    ${statusM}    true

TC_07_ValidateStatus_Keyword
        Fetch and Validate Get status code   50   200
        ${reponse}    Fetch and return get response  50
        Log To Console    ${reponse.content}


Fetch details and validate
    Arguments]   ${student_id}
    ${get_request}=   Get Request   e2e   api/studentsDetails/${student_id}
    ${json_reponse_2}=   To Json   ${get_request.content}
    ${Lfirst_name}=    Get Value From Json   ${json_reponse_2}   data.first_name
    ${first_name}=   Get From List   ${Lfirst_name}   0
    Should Be Equal   ${first_name}   ${update_first_name}


Fetch and Validate Get status code
    [Arguments]   ${student_id}   ${expectedStatusCode}
    Create Session   sname   ${App_Base_URL_testingworld}
    ${response}=   Get Request   sname   api/studentsDetails/${student_id}
    ${statusC}=   Convert To String    ${response.status_code}
    Should Be Equal    ${expectedStatusCode}    ${statusC}


Fetch and return get response
    [Arguments]   ${student_id}
    reate Session   sname   ${App_Base_URL_testingworld}
    ${response}=   Get Request   sname   api/studentsDetails/${student_id}
    [Return]   ${response}
