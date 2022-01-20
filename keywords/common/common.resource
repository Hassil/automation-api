*** Settings ***
Resource      load_components.resource

*** Variables ***
${_step_number}    0

*** Keywords ***
Step
      [Documentation]    Permite llevar el control de los pasos que se estan ejecutando
      [Arguments]    ${message}
      ${_step_number}=    Evaluate    ${_step_number} + 1
      Log console and log    Step ${_step_number}: ${message}
      BuiltIn.Set Test Variable    ${_step_number}

Log console and log
      [Documentation]    Permite ingresar un mensaje en consola y log
      [Arguments]    ${message}
      ${_timestamp} =    BuiltIn.Get Time
      BuiltIn.Log  ${message}
      BuiltIn.Log To Console  ${_timestamp} -- ${message}
