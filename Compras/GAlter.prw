#include 'parmtype.ch'
#INCLUDE "PROTHEUS.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "FONT.CH"
#INCLUDE "COLORS.CH"
#INCLUDE "FWMVCDEF.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "TOTVS.CH"
#Include "TopConn.ch"


user function GAlter(codProduto)


Local cUsrAlt :=""

  Local cQuery4 := ""

    cQuery4 := "     SELECT "
    cQuery4 += "    * "
    cQuery4 += " FROM "
    cQuery4 += "     "+RetSQLName('SB1')+" SB1 "

    cQuery4 += " WHERE "
    cQuery4 += "     SB1.D_E_L_E_T_ = ''  "
    cQuery4 += " AND B1_COD = '"+ cValToChar(codProduto) + "'"
    TCQuery cQuery4 New Alias "QRYPRO4"


   While !QRYPRO4->(EOF()) //Enquando não for fim de arquivo
      _cUser := UsrRetName( SubStr( Embaralha( QRYPRO4->B1_USERLGA, 1 ), 3, 6 ) )

      _cUser
      QRYPRO4->(dbSkip()) //Anda 1 registro pra frente
  EndDo
  QRYPRO4->(dbCloseArea()) //Fecha a área de trabalho 
  // DbSelectArea('SB1')
  // SB1->(DbSetOrder(1)) // Filial + Num
  // //Se conseguir posicionar no produto
  // If SB1->(DbSeek(FWxFilial('SB1') + codProduto ))
  // if  cUsrAlt  <>""
  //   RecLock('SB1', .F.)
  //     SB1->B1_ULTALT := cUsrAlt
  //     SB1->(MsUnlock())
  // endif
    
      
 // EndIf


return cUsrAlt   
