#include 'protheus.ch'
#include 'parmtype.ch'

user function GInclui2(codFornecedor)


Local cUsrInclui2 :=""



 DbSelectArea('SA2')
  SA2->(DbSetOrder(1)) // Filial + Num
  //Se conseguir posicionar no produto
  If SA2->(DbSeek(FWxFilial('SA2') + codFornecedor ))
      cUsrInclui := FWLeUserLg("A2_USERLGI", 1)
      RecLock('SA2', .F.)
      SA2->A2_USERCAD := cUsrInclui2
      SA2->(MsUnlock())
  EndIf


return cUsrInclui   
