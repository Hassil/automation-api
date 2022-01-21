*** Settings ***
Resource    ../keywords/common/load_components.resource

*** Test Cases ***
Validar información de estudiante
   Crear sesión
   Búsqueda de información por id
   Delete All Sessions
