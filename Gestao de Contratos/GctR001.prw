//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} GctR001
FunÃ§Ã£o que cria um exemplo de FWMsExcel
@author Leandro Souza
@since 06/10/2020
@version 1.0
    @example
    Gestao de Contrato - MOD069 - SIGAGCT - Relatório de Contrato Detalhado
    u_GctR001()
/*/

User Function GctR001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'GctR001.xml'
//************************************************************************************************
    cPerg    := "GCTR003"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "De  Numero do Contrato?",         "MV_PAR01", "MV_CH0", "C", TamSX3('CN9_NUMERO')[01], 0, "G", cValid,       "CN9", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero do contrato")
    u_zPutSX1(cPerg, "02", "Até Numero do Contrato?",         "MV_PAR02", "MV_CH1", "C", TamSX3('CN9_NUMERO')[01], 0, "G", cValid,       "CN9", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero de contrato")
   
    If ! Pergunte("GCTR003",.T.)
        Return

    EndIf

    //Pegando os dados
    cQuery := " SELECT "
    cQuery += "     CN9.CN9_NUMERO,"   
    cQuery += "     CN9.CN9_DTINIC,"    
    cQuery += "     CN9.CN9_DTASSI,"     
    cQuery += "     CN9.CN9_VIGE,"    
    cQuery += "     CN9.CN9_DTFIM," 
    cQuery += "     CN9.CN9_VLINI,"    
    cQuery += "     CN9.CN9_VLATU,"    
    cQuery += "     CN9.CN9_FLGREJ,"
    cQuery += "     CN9.CN9_DTENCE,"
    cQuery += "     CN9.CN9_REVISA,"
    cQuery += "     CN9.CN9_SALDO,"
    cQuery += "     CN9.CN9_DTREV,"
    cQuery += "     CN9.CN9_VLADIT,"
    cQuery += "     CN9.CN9_DTULST,"
    cQuery += "     CN9.CN9_SITUAC,"
    cQuery += "     CN9.CN9_FILORI,"
    cQuery += "     CN9.CN9_NATURE,"
    cQuery += "     CN9.CN9_CODFOR,"
    cQuery += "     CN9.CN9_FORLOJ,"
    cQuery += "     CN9.CN9_FORDES,"
    cQuery += "     CN9.CN9_USER,"    
    cQuery += "     CNB.CNB_REVISA," 
    cQuery += "     CNB.CNB_NUMERO,"   
    cQuery += "     CNB.CNB_ITEM,"
    cQuery += "     CNB.CNB_PRODUT,"
    cQuery += "     CNB.CNB_DESCRI,"
    cQuery += "     CNB.CNB_QUANT,"
    cQuery += "     CNB_VLUNIT,"
    cQuery += "     CNB.CNB_VLTOT,"
    cQuery += "     CNB.CNB_DTANIV "
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('CN9')+" CN9 "
    cQuery += " INNER JOIN " +RetSQLName('CNB020')+" CNB  ON CN9.CN9_NUMERO = CNB.CNB_CONTRA "
    cQuery += " AND CN9.CN9_REVISA = CNB.CNB_REVISA "
    cQuery += " WHERE "
    cQuery += "     CN9.D_E_L_E_T_ = '' "
    cQuery += " AND    CNB.D_E_L_E_T_ = '' "
    cQuery += " AND CN9_NUMERO >= '" + MV_PAR01 + "' "
    cQuery += " AND CN9_NUMERO <= '" + MV_PAR02 + "' "

    cQuery += " ORDER BY "
    cQuery += "     CN9.CN9_NUMERO "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irÃ¡ gerar o conteÃºdo do Excel
    oFWMsExcel := FWMSExcel():New("GcTR001")

   
    oFWMsExcel:AddworkSheet("Contratos")
    //Criando a Tabela
    oFWMsExcel:AddTable("Contratos","Contratos")
    //criando colunas
    oFWMsExcel:AddColumn("Contratos","Contratos","Contrato"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Data de Início"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Data de Assinatura"       ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Vigencia"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Data Final"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Vlr inicial"  ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Vlr. Atual"  ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Reajuste", 1)   
    oFWMsExcel:AddColumn("Contratos","Contratos","Encerramento"     ,1) 
    oFWMsExcel:AddColumn("Contratos","Contratos","Revisao"     ,1) 
    oFWMsExcel:AddColumn("Contratos","Contratos","Saldo"     ,1) 
    oFWMsExcel:AddColumn("Contratos","Contratos","Dt Revisao"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Vlr. Aditivo"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Ultimo Status"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Situacao"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Filial"     ,1)                                                         
    oFWMsExcel:AddColumn("Contratos","Contratos","Natureza"     ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Cod Fornecedor"     ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Loja"     ,1)   
    oFWMsExcel:AddColumn("Contratos","Contratos","Desc Fornecedor"       ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Usuário"       ,1) 
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Revisão"       ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao numero"       ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao item"       ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Produto"       ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Descricao Prod"       ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Qtde Prod"       ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Vlr Unit"       ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Vlr Total"       ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Medicao Data"       ,1)  

    //Criando as Linhas... Enquanto nÃ£o for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Contratos","Contratos",{QRYPRO->CN9_NUMERO,; 
            QRYPRO->CN9_DTINIC,;
            QRYPRO->CN9_DTASSI,;  
            QRYPRO->CN9_VIGE,;  
            QRYPRO->CN9_DTFIM,;  
            QRYPRO->CN9_VLINI,;  
            QRYPRO->CN9_VLATU,;
            QRYPRO->CN9_FLGREJ,;  
            QRYPRO->CN9_DTENCE,;
            QRYPRO->CN9_REVISA,;
            QRYPRO->CN9_SALDO,;
            QRYPRO->CN9_DTREV,;
            QRYPRO->CN9_VLADIT,;
            QRYPRO->CN9_DTULST,;
            QRYPRO->CN9_SITUAC,;
            QRYPRO->CN9_FILORI,;
            QRYPRO->CN9_NATURE,;
            QRYPRO->CN9_CODFOR,; 
            QRYPRO->CN9_FORLOJ,;
            QRYPRO->CN9_FORDES,;
            QRYPRO->CN9_USER,;
            QRYPRO->CNB_REVISA,;
            QRYPRO->CNB_NUMERO,;
            QRYPRO->CNB_ITEM,;
            QRYPRO->CNB_PRODUT,;
            QRYPRO->CNB_DESCRI,;
            QRYPRO->CNB_QUANT,;
            QRYPRO->CNB_VLUNIT,;
            QRYPRO->CNB_VLTOT,;
            QRYPRO->CNB_DTANIV})

        //Pulando Registro
        QRYPRO->(DbSkip())
    EndDo

    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)

    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexÃ£o com Excel
    oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
    oExcel:SetVisible(.T.)                 //Visualiza a planilha
    oExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas

    QRYPRO->(DbCloseArea())
    RestArea(aArea)
Return
