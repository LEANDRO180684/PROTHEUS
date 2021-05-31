//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} SP97R001
Função de consulta padrao nome de emitente
@author LEANDRO SOUZA - KETRA
@since 02/10/2020
@version 1.0
    @example
    u_SP97G001()
/*/

User Function SP97G001()

Local cQuery := ""

cQuery := "SELECT "
cQuery += " DISTINCT "
cQuery += " ZV_NOMUSER " 
cQuery += " FROM "
cQuery += "SZV020 ZV"
cQuery += " WHERE "
cQuery += " ZV.D_E_L_E_T_ <> '*' "

TCQuery cQuery New Alias "QRYPRO"


Return 


