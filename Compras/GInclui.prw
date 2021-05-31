#include 'protheus.ch'
#include 'parmtype.ch'

user function GInclui(codProduto)


Local cUsrInclui :=""



 DbSelectArea('SB1')
  SB1->(DbSetOrder(1)) // Filial + Num
  //Se conseguir posicionar no produto
  If SB1->(DbSeek(FWxFilial('SB1') + codProduto ))
      cUsrInclui := FWLeUserLg("B1_USERLGI", 1)
      RecLock('SB1', .F.)
      SB1->B1_USERCAD := cUsrInclui
      SB1->(MsUnlock())
  EndIf


return cUsrInclui   
