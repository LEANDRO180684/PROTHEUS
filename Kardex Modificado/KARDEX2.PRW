//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} EST04R001
FunÃ§Ã£o que cria um exemplo de FWMsExcel
@author LEANDRO SOUZA - KETRA
@since 10/11/2020
@version 1.0
    @example
    u_EST04R001()
/*/

User Function EST04R001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'EST04R001.xml'


//************************************************************************************************
    cPerg    := "EST04R003"

    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""

    u_zPutSX1(cPerg, "01", "De Num Doc?",         "MV_PAR01", "MV_CH0", "C", TamSX3('CQ_NUMREQ')[01], 0, "G", cValid,       "SCQDOC", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero do Documento gerado pós baixa ")
    u_zPutSX1(cPerg, "02", "Ate Num Doc?",        "MV_PAR02", "MV_CH1", "C", TamSX3('CQ_NUMREQ')[01], 0, "G", cValid,       "SCQDOC", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero do Documento gerado pós baixa ")
    u_zPutSX1(cPerg, "03", "De Num SA?",         "MV_PAR03", "MV_CH2", "C", TamSX3('CQ_NUM')[01], 0, "G", cValid,       "SCQSA", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero da SA ")
    u_zPutSX1(cPerg, "04", "Ate Num SA?",        "MV_PAR04", "MV_CH3", "C", TamSX3('CQ_NUM')[01], 0, "G", cValid,       "SCQSA", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero da SA ")
    u_zPutSX1(cPerg, "05", "De Data inclusao?",  "MV_PAR05", "MV_CH4", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de inclusao da movimentacao ")
    u_zPutSX1(cPerg, "06", "Ate Data?",          "MV_PAR06", "MV_CH5", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de inclusao da movimentacao ")
    u_zPutSX1(cPerg, "07", "Filial 01 a 06?",    "MV_PAR07", "MV_CH6", "C", TamSX3('CQ_FILIAL')[01], 0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a filial ")
    u_zPutSX1(cPerg, "08", "Local 01 a 08?",    "MV_PAR08", "MV_CH7", "C", TamSX3('CQ_LOCAL')[01], 0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Local ") 
    u_zPutSX1(cPerg, "09", "Produto?",          "MV_PAR09", "MV_CH8", "C", TamSX3('CQ_PRODUTO')[01], 0, "G", cValid,       "SCQPRO",   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o produto ")
    
    If ! Pergunte("EST04R003",.T.)
        Return
    EndIf•

    //Pegando os dado•s
    
    cQuery := " SELECT DISTINCT"
    cQuery += "     CQ.CQ_FILIAL, "
    cQuery += "     CQ.CQ_LOCAL, "
    cQuery += "     CQ.CQ_PRODUTO, "
    cQuery += "     CQ.CQ_DESCRI, "
    cQuery += "     CQ.CQ_ITEM, "
    cQuery += "     CQ.CQ_NUM, "
    cQuery += "     CQ.CQ_NUMREQ, "
    cQuery += "     D3.D3_TM, "
    cQuery += "     CQ.CQ_QUANT, "
    cQuery += "     CQ.CQ_CC, "
    cQuery += "     CQ.CQ_DATPRF "
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SCQ')+" CQ "
    cQuery += " LEFT JOIN " +RetSQLName('SD3')+" D3  ON CQ.CQ_PRODUTO = D3.D3_COD AND D3.D_E_L_E_T_ = ''"
    cQuery += " WHERE "
    cQuery += "     CQ.D_E_L_E_T_ = '' "
    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND CQ_NUMREQ >= '" + MV_PAR01 + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND CQ_NUMREQ <= '" + MV_PAR02 + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND CQ_NUM >= '" + MV_PAR03 + "' ")
    cQuery += IF ((EMPTY (MV_PAR04)), "", " AND CQ_NUM <= '" + MV_PAR04 + "' ")
    cQuery += IF ((EMPTY (MV_PAR05)), "", " AND CQ_DATPRF >= '" + DTOS(MV_PAR05) + "' ")
    cQuery += IF ((EMPTY (MV_PAR06)), "", " AND CQ_DATPRF <= '" + DTOS(MV_PAR06) + "' ")
    cQuery += IF ((EMPTY (MV_PAR07)), "", " AND CQ_FILIAL = '" + MV_PAR07 + "' ")
    cQuery += IF ((EMPTY (MV_PAR08)), "", " AND CQ_LOCAL = '" + MV_PAR08 + "' ")
    cQuery += IF ((EMPTY (MV_PAR09)), "", " AND CQ_PRODUTO = '" + MV_PAR09 + "' ")
    cQuery += " ORDER BY " 
    cQuery += "   CQ.CQ_NUM "
    TCQuery cQuery New Alias "QRYPRO1"

    //Criando o objeto que irÃ¡ gerar o conteÃºdo do Excel
        oFWMsExcel := FWMSExcel():New("EST04R001")
        oFWMsExcel:AddworkSheet("EST04")
    //Criando a Tabela
        oFWMsExcel:AddTable("EST04","EST04")
    //criando colunas
        oFWMsExcel:AddColumn("EST04","EST04","Filial "         ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Local "         ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Cod Produto "       ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Descri Produto "         ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Item "         ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Num SA "  ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Num Documento "  ,1)
        oFWMsExcel:AddColumn("EST04","EST04","TES ",1)
        oFWMsExcel:AddColumn("EST04","EST04","Quantidade "     ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Centro de Custo "     ,1)
        oFWMsExcel:AddColumn("EST04","EST04","Data de movimentacao "     ,1)

    //Criando as Linhas... Enquanto nÃ£o for fim da query
        While !(QRYPRO1->(EoF()))
            oFWMsExcel:AddRow("EST04","EST04",{QRYPRO1->CQ_FILIAL,;
                QRYPRO1->CQ_LOCAL,;
                QRYPRO1->CQ_PRODUTO,;
                QRYPRO1->CQ_DESCRI,;
                QRYPRO1->CQ_ITEM,;
                QRYPRO1->CQ_NUM,;
                QRYPRO1->CQ_NUMREQ,;
                QRYPRO1->D3_TM,;
                QRYPRO1->CQ_QUANT,;
                QRYPRO1->CQ_CC,;
                sTod(QRYPRO1->CQ_DATPRF)})
        //Pulando Registro
            QRYPRO1->(DbSkip())
        EndDo

        //Ativando o arquivo e gerando o xml
        oFWMsExcel:Activate()
        oFWMsExcel:GetXMLFile(cArquivo)

        //Abrindo o excel e abrindo o arquivo xml
        oExcel := MsExcel():New()           //Abre uma nova conexÃ£o com Excel
        oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
        oExcel:SetVisible(.T.)              //Visualiza a planilha
        oExcel:Destroy()                    //Encerra o processo do gerenciador de tarefas

        QRYPRO1->(DbCloseArea())
        RestArea(aArea)
        Return
