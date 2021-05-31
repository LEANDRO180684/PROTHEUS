//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} ComR001
Função que cria um exemplo de FWMsExcel
@author Atilio
@since 06/08/2016
@version 1.0
    @example
    u_ComR001()
/*/

User Function ComR001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'ComR001.xml'
//************************************************************************************************
    cPerg    := "COMR001"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "De Pedido?",       "MV_PAR01", "MV_CH0", "C", TamSX3('C7_NUM')[01], 0, "G", cValid,       "SC7", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Pedido inicial")
    u_zPutSX1(cPerg, "02", "At� Pedido?",      "MV_PAR02", "MV_CH1", "C", TamSX3('C7_NUM')[01], 0, "G", cValid,       "SC7", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Pedido final")
    u_zPutSX1(cPerg, "03", "De Data?",         "MV_PAR03", "MV_CH2", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data inicial a ser considerada")
    u_zPutSX1(cPerg, "04", "Ate Data?",        "MV_PAR04", "MV_CH2", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data Final a ser considerada")


    If ! Pergunte("COMR001",.T.)
        Return
    EndIf

    //Pegando os dados
    cQuery := " SELECT "
    cQuery += "     SA2.A2_NOME,"   
    cQuery += "     SA2.A2_COD,"    
    cQuery += "     SA2.A2_LOJA,"   
    cQuery += "     SA2.A2_END,"    
    cQuery += "     SA2.A2_BAIRRO," 
    cQuery += "     SA2.A2_CEP,"    
    cQuery += "     SA2.A2_MUN,"    
    cQuery += "     SA2.A2_EST,"    
    cQuery += "     SA2.A2_CGC,"    
    cQuery += "     SA2.A2_INSCR," 
    cQuery += "     SA2.A2_TEL,"    
    cQuery += "     SA2.A2_FAX,"        
    cQuery += "     SC7.C7_NUM, "
    cQuery += "     SC7.C7_PRODUTO, " 
    cQuery += "		SC7.C7_DESCRI," 
    cQuery += "		SC7.C7_ITEM,"  
    cQuery += "		SC7.C7_UM,"    
    cQuery += "		SC7.C7_QUANT,"  
    cQuery += "		SC7.C7_SEGUM," 
    cQuery += "		SC7.C7_QTSEGUM,"
    cQuery += "		SC7.C7_PRECO,"     
    cQuery += "		SC7.C7_IPI,"  
    cQuery += "		SC7.C7_TOTAL,"    
    cQuery += "		SC7.C7_DATPRF,"
    cQuery += "		SC7.C7_CC,"    
    cQuery += "		SC7.C7_NUMSC,"        
    cQuery += "		SC7.C7_USER,"   
    cQuery += "		SC7.C7_COND,"   
    cQuery += "		SC7.C7_ITEMCTA,"
    cQuery += "		SC7.C7_CONTA,"  
    cQuery += "		SC7.C7_OBS"  
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SC7')+" SC7 "
    cQuery += "INNER JOIN " +RetSQLName('SA2')+" SA2  ON SA2.A2_COD = SC7.C7_FORNECE"
    cQuery += " WHERE "
    cQuery += "     SC7.D_E_L_E_T_ = '' "
    cQuery += " AND C7_NUM >= '" + MV_PAR01 + "' "
    cQuery += " AND C7_NUM <= '" + MV_PAR02 + "' "
    cQuery += " AND C7_EMISSAO >= '" + MV_PAR03 + "' "
    cQuery += " AND C7_EMISSAO <= '" + MV_PAR04 + "' "
    cQuery += " ORDER BY "
    cQuery += "     SC7.C7_NUM "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New("COMR001")

   
    oFWMsExcel:AddworkSheet("Pedidos")
    //Criando a Tabela
    oFWMsExcel:AddTable("Pedidos","Pedidos")
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Cod Fornece" ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Loja"		   ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","End"		   ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Bairro"      ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","CEP"    	   ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Mun"         ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Est"         ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","CGC"         ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Inscr"       ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Tel"         ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Fax"         ,1)  
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Numero Ped"  ,1)
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Produto"  ,1)
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Descricao",1)                                                       
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Item"     ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","UM"       ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Qntd"     ,1)    
	oFWMsExcel:AddColumn("Pedidos","Pedidos","Seg UM"   ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Qt Seg UM",1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Preco"    ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Ipi"      ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Total"    ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Dat Prf"  ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","CC"       ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Num SC"   ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","User "    ,1) 
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Cond"     ,1)     
	oFWMsExcel:AddColumn("Pedidos","Pedidos","Item CTA" ,1)
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Conta"    ,1)         
    oFWMsExcel:AddColumn("Pedidos","Pedidos","Obs"      ,1) 


    //Criando as Linhas... Enquanto não for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Pedidos","Pedidos",{QRYPRO->A2_COD,;             
            QRYPRO->A2_LOJA,; 
            QRYPRO->A2_END,;
            QRYPRO->A2_BAIRRO,;
            QRYPRO->A2_CEP,;  
            QRYPRO->A2_MUN,;  
            QRYPRO->A2_EST,;  
            QRYPRO->A2_CGC,;  
            QRYPRO->A2_INSCR,;
            QRYPRO->A2_TEL,;  
            QRYPRO->A2_FAX,; 
            QRYPRO->C7_NUM,;
            QRYPRO->C7_DESCRI,;
            QRYPRO->C7_PRODUTO,;
 		    QRYPRO->C7_ITEM  ,;
            QRYPRO->C7_UM    ,;
            QRYPRO->C7_QUANT,;
            QRYPRO->C7_SEGUM,;
            QRYPRO->C7_QTSEGUM,;
            QRYPRO->C7_PRECO,;
            QRYPRO->C7_IPI,;
            QRYPRO->C7_TOTAL,;
            QRYPRO->C7_DATPRF,;
            QRYPRO->C7_CC  ,;
            QRYPRO->C7_NUMSC,;
            QRYPRO->C7_USER,;
            QRYPRO->C7_COND,;
            QRYPRO->C7_ITEMCTA,;
            QRYPRO->C7_CONTA,;
            QRYPRO->C7_OBS})

        //Pulando Registro
        QRYPRO->(DbSkip())
    EndDo

    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)

    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
    oExcel:SetVisible(.T.)                 //Visualiza a planilha
    oExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas

    QRYPRO->(DbCloseArea())
    RestArea(aArea)
Return



/*/
