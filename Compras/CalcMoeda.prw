#INCLUDE "PROTHEUS.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "FONT.CH"
#INCLUDE "COLORS.CH"
#INCLUDE "FWMVCDEF.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "TOTVS.CH"

/*
+-----------+------------+----------------+-------------------+-------+---------------+
| Programa  | FSFORMED   | Desenvolvedor  | Silvio C. Stecca  | Data  | 12/08/2020    |
+-----------+------------+----------------+-------------------+-------+---------------+
| Descricao | Retorna a descrição do campo Nome Fornec. Mant. Medicao.                |
+-----------+-------------------------------------------------------------------------+
| Modulos   | SIGACom                                                                 |
+-----------+-------------------------------------------------------------------------+
| Processos |                                                                         |
+-----------+-------------------------------------------------------------------------+
|                  Modificacoes desde a construcao inicial                            |
+----------+-------------+------------------------------------------------------------+
| DATA     | PROGRAMADOR | MOTIVO                                                     |
|          |             |                                                            |
+----------+-------------+------------------------------------------------------------+
*/
User function CalcMoeda(cNumPed)

    Local cTipoMoeda := ""
    Local cVlrTotalPedido  := ""
    Local cVlrTotal  := ""
    Local cValorMoeda := ""
    Local dDataB := Date() -1
    Local dDataRef := Date()
    Local dData 



    If Dow(dDataB) == 1    // Se for domingo
        dData := DTOS(dDataRef - 3)
    Else                                   // Se for dia normal
        dData := DTOS(dDataRef -1)
    EndIf

    aAreaSCR := SCR->(GetArea())	
	dbselectarea("SCR")
	SCR->(DbSetOrder(4))
	dbseek(xfilial("SCR")+cNumPed)	//uma linha 
        cVlrTotalPedido := SCR->CR_TOTAL
        cTipoMoeda      := SCR->CR_MOEDA
	RestArea(aAreaSCR)

    aAreaSM2 := SM2->(GetArea())	
	dbselectarea("SM2")
	SM2->(DbSetOrder(1))
	dbseek(dData)	
        if cTipoMoeda = 1
            cValorMoeda := 1
        Elseif cTipoMoeda = 2
            cValorMoeda := SM2->M2_MOEDA2
        Elseif cTipoMoeda = 3
            cValorMoeda := SM2->M2_MOEDA3
        Elseif cTipoMoeda = 4
            cValorMoeda := SM2->M2_MOEDA4
        Elseif cTipoMoeda = 5
            cValorMoeda := SM2->M2_MOEDA5   
        endif
	RestArea(aAreaSM2)

    
    cVlrTotal = (cVlrTotalPedidos * cValorMoeda)


Return cVlrTotal
