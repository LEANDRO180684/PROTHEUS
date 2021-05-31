//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} SP97R001
FunÃ§Ã£o que cria um exemplo de FWMsExcel
@author LEANDRO SOUZA - KETRA
@since 02/10/2020
@version 1.0
    @example
    u_SP97R001()
/*/

User Function SP97R001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'SP97R001.xml'
    Local cNomeFornece := ""
    Local cNomeCliente := ""


//************************************************************************************************
    cPerg    := "SP97R019"

    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""

    u_zPutSX1(cPerg, "01", "De Num SP?",               "MV_PAR01", "MV_CH0",  "C", TamSX3('ZV_NUM')[01],     0, "G", cValid,       "SZV",      cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero da SP ")
    u_zPutSX1(cPerg, "02", "Ate Num SP?",              "MV_PAR02", "MV_CH1",  "C", TamSX3('ZV_NUM')[01],     0, "G", cValid,       "SZV",      cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero da SP ")
    u_zPutSX1(cPerg, "03", "De Data inclusao?",        "MV_PAR03", "MV_CH2",  "D", 08,                       0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de inclusao da SP ")
    u_zPutSX1(cPerg, "04", "Ate Data?",                "MV_PAR04", "MV_CH3",  "D", 08,                       0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de inclusao da SP ")
    u_zPutSX1(cPerg, "05", "Filial 01 a 06?",          "MV_PAR05", "MV_CH4",  "C", TamSX3('ZV_FILIAL')[01],  0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a filial ")
    u_zPutSX1(cPerg, "06", "Codigo do Fornecedor?",    "MV_PAR06", "MV_CH6",  "C", TamSX3('ZV_FORNECE')[01], 0, "G", cValid,       "SZVNFO",   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o codigo do fornecedor ")
    u_zPutSX1(cPerg, "07", "Nome do Fornecedor?",      "MV_PAR07", "MV_CH7",  "C", TamSX3('ZV_NOMFOR')[01],  0, "G", cValid,       "SZVNOF",   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o nome do fornecedor ")
    u_zPutSX1(cPerg, "08", "Nome do Emitente?",        "MV_PAR08", "MV_CH8",  "C", TamSX3('ZV_NOMUSER')[01], 0, "G", cValid,       "SZVEMI",   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o nome do Emitente ")
    u_zPutSX1(cPerg, "09", "De Data de vencimento?",   "MV_PAR09", "MV_CH9",  "D", 08,                       0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de vencimento da SP ")
    u_zPutSX1(cPerg, "10", "Ate Data de vencimento?",  "MV_PAR10", "MV_CH10", "D", 08,                       0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de vencimento da SP ")
    If ! Pergunte("SP97R019",.T.)
        Return
    EndIf•

    cNomeFornece := ALLTRIM(MV_PAR07)
    cNomeCliente := ALLTRIM(MV_PAR08)

    //Pegando os dado•s
    
    cQuery := " SELECT "
    cQuery += "     ZV.ZV_NUM, "
    cQuery += "     ZV.ZV_FILIAL, "
    cQuery += "     ZV.ZV_APROV, "
    cQuery += "     ZV.ZV_TIPOSP, "
    cQuery += "     ZV.ZV_TIPO, "
    cQuery += "     ZV.ZV_MOEDA, "
    cQuery += "     ZV.ZV_FORNECE, "
    cQuery += "     ZV.ZV_NOMFOR, "
    cQuery += "     ZV.ZV_LOJA, "
    cQuery += "     ZV.ZV_MOTIVO, "
    cQuery += "     ZV.ZV_EMISSAO, "
    cQuery += "     ZV.ZV_VENCTO, "
    cQuery += "     ZV.ZV_VALOR, "
    cQuery += "     ZV.ZV_VLRBRUT, "
    cQuery += "     ZV.ZV_NOMUSER, "
    cQuery += "     ZV.ZV_STATUS, "
    cQuery += "     ZV.ZV_DTINCL, "
    cQuery += "     ZV.ZV_HRINCL, "
    cQuery += "     ZV.ZV_DTLIBF, "
    cQuery += "     ZV.ZV_HRLIBF, "
    cQuery += "     ZV.ZV_USULIB, "
    cQuery += "     ZV.ZV_WF, "
    cQuery += "     ZV.ZV_WFID, "
    cQuery += "     ZV.ZV_WFOBS, "
    cQuery += "     ZV.ZV_XCONFIR, "
    cQuery += "     ZV.ZV_XCANCON "
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SZV')+" ZV "
    cQuery += " WHERE "
    cQuery += "     ZV.D_E_L_E_T_ = '' "
    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND ZV_NUM >= '" + MV_PAR01 + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND ZV_NUM <= '" + MV_PAR02 + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND ZV_DTINCL >= '" + DTOS(MV_PAR03) + "' ")
    cQuery += IF ((EMPTY (MV_PAR04)), "", " AND ZV_DTINCL <= '" + DTOS(MV_PAR04) + "' ")
    cQuery += IF ((EMPTY (MV_PAR05)), "", " AND ZV_FILIAL = '" + MV_PAR05 + "' ")
    cQuery += IF ((EMPTY (MV_PAR06)), "", " AND ZV_FORNECE = '" + MV_PAR06 + "' ")
    cQuery += IF ((EMPTY (MV_PAR07)), "", " AND ZV_NOMFOR LIKE '%" + cNomeFornece + "%' ")   
    cQuery += IF ((EMPTY (MV_PAR08)), "", " AND ZV_NOMUSER LIKE '%" + cNomeCliente + "%' ")
    cQuery += IF ((EMPTY (MV_PAR09)), "", " AND ZV_VENCTO >= '" + DTOS(MV_PAR09) + "' ")
    cQuery += IF ((EMPTY (MV_PAR10)), "", " AND ZV_VENCTO <= '" + DTOS(MV_PAR10) + "' ")
    cQuery += " ORDER BY " 
    cQuery += "   ZV.ZV_NUM "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irÃ¡ gerar o conteÃºdo do Excel
••        oFWMsExcel := FWMSExcel():New("SP97R001")
        oFWMsExcel:AddworkSheet("SP97")
    //Criando a Tabela
        oFWMsExcel:AddTable("SP97","SP97")
    //criando colunas
        oFWMsExcel:AddColumn("SP97","SP97","Numero "         ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Filial "         ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Gr. Aprov "       ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Tipo SP "         ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Tipo "         ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Moeda "  ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Num Fornec "  ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Nome Fornec ",1)
        oFWMsExcel:AddColumn("SP97","SP97","Loja "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Referente a "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Emissao "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Vencimento "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Vr. Titulo "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Vr. Bruto "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Emitente "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Status "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Dt Inclusao "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Hr inclusao "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Dt Lib Fisca "     ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Hr Lib Fisca "       ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Usuario Lib "       ,1)
        oFWMsExcel:AddColumn("SP97","SP97","WorkFlow "       ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Proc. ID WF "       ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Data Confirm "       ,1)
        oFWMsExcel:AddColumn("SP97","SP97","Dt Can Confi "       ,1)


    //Criando as Linhas... Enquanto nÃ£o for fim da query
        While !(QRYPRO->(EoF()))
            oFWMsExcel:AddRow("SP97","SP97",{QRYPRO->ZV_NUM,;
                QRYPRO->ZV_FILIAL,;
                QRYPRO->ZV_APROV,;
                QRYPRO->ZV_TIPOSP,;
                QRYPRO->ZV_TIPO,;
                QRYPRO->ZV_MOEDA,;
                QRYPRO->ZV_FORNECE,;
                QRYPRO->ZV_NOMFOR,;
                QRYPRO->ZV_LOJA,;
                QRYPRO->ZV_MOTIVO,;
                sTod(QRYPRO->ZV_EMISSAO),;
                sTod(QRYPRO->ZV_VENCTO),;
                QRYPRO->ZV_VALOR,;
                QRYPRO->ZV_VLRBRUT,;
                QRYPRO->ZV_NOMUSER,;
                QRYPRO->ZV_STATUS,;
                sTod(QRYPRO->ZV_DTINCL),;
                QRYPRO->ZV_HRINCL,;
                sTod(QRYPRO->ZV_DTLIBF),;
                QRYPRO->ZV_HRLIBF,;
                QRYPRO->ZV_USULIB,;
                QRYPRO->ZV_WF,;
                QRYPRO->ZV_WFID,;
                QRYPRO->ZV_XCONFIR,;
                QRYPRO->ZV_XCANCON})

        //Pulando Registro
            QRYPRO->(DbSkip())
        EndDo

        //Ativando o arquivo e gerando o xml
        oFWMsExcel:Activate()
        oFWMsExcel:GetXMLFile(cArquivo)

        //Abrindo o excel e abrindo o arquivo xml
        oExcel := MsExcel():New()           //Abre uma nova conexÃ£o com Excel
        oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
        oExcel:SetVisible(.T.)              //Visualiza a planilha
        oExcel:Destroy()                    //Encerra o processo do gerenciador de tarefas

        QRYPRO->(DbCloseArea())
        RestArea(aArea)
        Return
