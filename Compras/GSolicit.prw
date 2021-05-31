#include 'protheus.ch'
#include 'parmtype.ch'

user function GSolicit(cMsblql)


Local cSolicit :=""


  If cMsblql == "2"

    cSolicit := FwGetUserName(RetCodUsr())

  ENDIF




return cSolicit 
