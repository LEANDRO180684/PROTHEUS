User function DescMoeda(cNumPed)

    Local cMoeda1 := SuperGetMv("MV_MOEDA1", .F.,"Real" )
    Local cMoeda2 := SuperGetMv("MV_MOEDA2", .F.,"Dolar" )
    Local cMoeda3 := SuperGetMv("MV_MOEDA3", .F.,"Ufir" )
    Local cMoeda4 := SuperGetMv("MV_MOEDA4", .F.,"Euro" )
    Local cMoeda5 := SuperGetMv("MV_MOEDA5", .F.,"Iene" )
    Local cDescMoeda := ""
    Local cTipoMoeda := ""


    aAreaSCR := SCR->(GetArea())	
	dbselectarea("SCR")
	SCR->(DbSetOrder(4))
	dbseek(xfilial("SCR")+cNumPed)	//uma linha 
        cTipoMoeda      := SCR->CR_MOEDA
	RestArea(aAreaSCR)

        if cTipoMoeda = 1
            cDescMoeda := cMoeda1 
        Elseif cTipoMoeda = 2
            cDescMoeda := cMoeda2
        Elseif cTipoMoeda = 3
            cDescMoeda := cMoeda3
        Elseif cTipoMoeda = 4
            cDescMoeda := cMoeda4
        Elseif cTipoMoeda = 5
            cDescMoeda := cMoeda5  
        endif

Return cDescMoeda
