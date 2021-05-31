//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} Mod69R001
FunÃ§Ã£o que cria um exemplo de FWMsExcel
@author Leandro Souza
@since 24/11/2020 Alterado em 14/04/2021
@version 1.0
    @example
    Gestao de Contrato - MOD069 - SIGAGCT - Relatório de Contrato Detalhado
    u_Mod69R001()
/*/

User Function Mod69R001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'Mod69R001.xml'
    Local cNomeFornece := ""
//************************************************************************************************
     cPerg    := "Mod69R012"
       
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
     u_zPutSX1(cPerg, "03", "Data de inicio - De:",            "MV_PAR03", "MV_CH2", "D", 08,                       0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a Data de Inicio")
     u_zPutSX1(cPerg, "04", "Data de inicio - Ate:",           "MV_PAR04", "MV_CH3", "D", 08,                       0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a Data de Inicio")
     u_zPutSX1(cPerg, "05", "Filial:",                         "MV_PAR05", "MV_CH4", "C", TamSX3('CN9_FILORI')[01], 0, "G", cValid,       cF3, cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a Filial")
     u_zPutSX1(cPerg, "06", "Fornecedor:",                     "MV_PAR06", "MV_CH5", "C", TamSX3('CN9_FORDES')[01], 0, "G", cValid,       "SA2", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Fornecedor")
     u_zPutSX1(cPerg, "07", "Situacao:",                       "MV_PAR07", "MV_CH6", "C", TamSX3('CN9_SITUAC')[01], 0, "G", cValid,       cF3, cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Situação 01-Cancelado.02-Elaboração.03-Emitido.04-Aprovação.05-Vigente.06-Paralisa.07-Sol. Finalização.08-Finali.09-Revisão.10-Revisado.")
     u_zPutSX1(cPerg, "08", "Adiantamento:",                       "MV_PAR08", "MV_CH7", "C", 1                       , 0, "G", cValid,       cF3, cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Adiantamento 1-Sim. 2-Não.")

   
   
    
    If !Pergunte("Mod69R012",.T.)
        Return
    EndIf
    

    cNomeFornece := ALLTRIM(MV_PAR06)
    
    //Pegando os dados
    cQuery := " SELECT DISTINCT"
    cQuery += "     CN9.CN9_FILORI,"
    cQuery += "     CN9.CN9_CODFOR,"
    cQuery += "     CN9.CN9_FORDES," 
    cQuery += "     CN9.CN9_NUMERO,"   
    cQuery += "     CN9.CN9_DTINIC," 
    cQuery += "     CN9.CN9_DTFIM,"   
    cQuery += "     CN9.CN9_DTASSI,"     
    cQuery += "     CN9.CN9_VIGE,"    

    cQuery += "     CASE CN9.CN9_UNVIGE WHEN '1' THEN 'Dias' "
    cQuery += "     WHEN '2' THEN 'Meses'"
    cQuery += "     WHEN '3' THEN 'Anos'"
    cQuery += "     WHEN '4' THEN 'Indeterminada'"
    cQuery += "     END AS 'CN9_UNVIGE',"
    cQuery += "     CN1.CN1_DESCRI," 
    
    cQuery += "     CASE CN9.CN9_SITUAC "
    cQuery += "     WHEN '01' THEN 'Cancelado' "
    cQuery += "     WHEN '02' THEN 'Elaboracao'"
    cQuery += "     WHEN '03' THEN 'Emitido'"
    cQuery += "     WHEN '04' THEN 'Aprovacao'"
    cQuery += "     WHEN '05' THEN 'Vigente'"
    cQuery += "     WHEN '06' THEN 'Paralisado'"
    cQuery += "     WHEN '07' THEN 'Sol. Finalizada'"
    cQuery += "     WHEN '08' THEN 'Finalizado'"
    cQuery += "     WHEN '09' THEN 'Revisao'"
    cQuery += "     WHEN '10' THEN 'Revisado'"
    cQuery += "     WHEN '11' THEN 'Sit. em branco'"
    cQuery += "     END AS 'CN9_SITUAC',"

    cQuery += "     CNB.CNB_REVISA,"    
    cQuery += "     CN9.CN9_DTREV,"

    cQuery += "     CN9.CN9_VLINI,"
    cQuery += "     CN9.CN9_VLATU,"
    cQuery += "     CN9.CN9_SALDO"

     if(MV_PAR08 = '1')
        cQuery += "     , CNX.CNX_VLADT,"
        cQuery += "     CNX.CNX_DTADT"
    endif
  

    cQuery += " FROM "
    cQuery += "     "+RetSQLName('CN9')+" CN9 "

    cQuery += " INNER JOIN " +RetSQLName('CNB020')+" CNB  ON CN9.CN9_NUMERO = CNB.CNB_CONTRA AND CN9.CN9_FILIAL = CNB.CNB_FILIAL"
    cQuery += " INNER JOIN " +RetSQLName('CN1020')+" CN1  ON CN9.CN9_TPCTO = CN1.CN1_CODIGO AND CN9.CN9_FILIAL = CN1.CN1_FILIAL"

    if(MV_PAR08 = '1')
        cQuery += " INNER JOIN " +RetSQLName('CNX020')+" CNX  ON CN9.CN9_NUMERO = CNX.CNX_CONTRA AND CN9.CN9_FILIAL = CNX.CNX_FILIAL"
    endif
    cQuery += " AND CN9.CN9_REVISA = CNB.CNB_REVISA "
    cQuery += " WHERE "
    cQuery += "     CN9.D_E_L_E_T_ = '' "
    cQuery += " AND    CNB.D_E_L_E_T_ = '' "
    cQuery += " AND    CNX.D_E_L_E_T_ = '' "
    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND CN9_NUMERO >= '" + MV_PAR01 + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND CN9_NUMERO <= '" + MV_PAR02 + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND CN9_DTINIC >= '" + DTOS(MV_PAR03) + "' ")
    cQuery += IF ((EMPTY (MV_PAR04)), "", " AND CN9_DTINIC <= '" + DTOS(MV_PAR04) + "' ")
    cQuery += IF ((EMPTY (MV_PAR05)), "", " AND CN9_FILORI = '" + MV_PAR05 + "' ")
    cQuery += IF ((EMPTY (MV_PAR06)), "", " AND CN9_CODFOR = '" + ALLTRIM(MV_PAR06) + "' ")  
    cQuery += IF ((EMPTY (MV_PAR07)), "", " AND CN9_SITUAC = '" + MV_PAR07 + "' ")

    cQuery += " ORDER BY "
    cQuery += "     CNB.CNB_REVISA DESC "
    
    MemoWrite("\system\Mod69R001.sql",cQuery)//MARCIO 10/02/2021
    
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irÃ¡ gerar o conteÃºdo do Excel
    oFWMsExcel := FWMSExcel():New("Mod69R001")

   
    oFWMsExcel:AddworkSheet("Contratos")
    //Criando a Tabela
    oFWMsExcel:AddTable("Contratos","Contratos")
    //criando colunas
    oFWMsExcel:AddColumn("Contratos","Contratos","Filial"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Cod. Fornecedor"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Fornecedor"       ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Num. Contrato"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Data Inicial"         ,1)  
    oFWMsExcel:AddColumn("Contratos","Contratos","Data Final"     ,1) 
    oFWMsExcel:AddColumn("Contratos","Contratos","Data Assinatura"  ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Vigencia"  ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Unidade Vigencia", 1)   
    oFWMsExcel:AddColumn("Contratos","Contratos","Descri Tipo de Contrato"     ,1) 
    oFWMsExcel:AddColumn("Contratos","Contratos","Situacao"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Num Revisao"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Dt Ult Revisao"     ,1)    
    oFWMsExcel:AddColumn("Contratos","Contratos","Valor inicial"     ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Valor Atual"     ,1)
    oFWMsExcel:AddColumn("Contratos","Contratos","Saldo"     ,1)
    
     if(MV_PAR08 = '1')
    
        oFWMsExcel:AddColumn("Contratos","Contratos","Vl Adiantamento"     ,1)   
        oFWMsExcel:AddColumn("Contratos","Contratos","Dt Adiantamento"     ,1)  
    endif
        

    //Criando as Linhas... Enquanto nÃ£o for fim da query
    While !(QRYPRO->(EoF()))
        IF (MV_PAR08 = '1') 
        oFWMsExcel:AddRow("Contratos","Contratos",{QRYPRO->CN9_FILORI,; 
            QRYPRO->CN9_CODFOR,; 
            QRYPRO->CN9_FORDES,;
            QRYPRO->CN9_NUMERO,;
            sTod(QRYPRO->CN9_DTINIC),;
            sTod(QRYPRO->CN9_DTFIM),; 
            sTod(QRYPRO->CN9_DTASSI),;  
            QRYPRO->CN9_VIGE,;  
            QRYPRO->CN9_UNVIGE,;
            QRYPRO->CN1_DESCRI,; 
            QRYPRO->CN9_SITUAC,;
            QRYPRO->CNB_REVISA,; 
            sTod(QRYPRO->CN9_DTREV),;
            "R$"+TRANSFORM(QRYPRO->CN9_VLINI, "@E 999,999,999.99"),;
            "R$"+TRANSFORM(QRYPRO->CN9_VLATU, "@E 999,999,999.99"),;
            "R$"+TRANSFORM(QRYPRO->CN9_SALDO,  "@E 999,999,999.99") ,;
            "R$"+TRANSFORM(QRYPRO->CNX_VLADT,  "@E 999,999,999.99"),; 
            sTod(QRYPRO->CNX_DTADT)})
        //Pulando Registro
        QRYPRO->(DbSkip())
        else

        oFWMsExcel:AddRow("Contratos","Contratos",{QRYPRO->CN9_FILORI,; 
            QRYPRO->CN9_CODFOR,; 
            QRYPRO->CN9_FORDES,;
            QRYPRO->CN9_NUMERO,;
            sTod(QRYPRO->CN9_DTINIC),;
            sTod(QRYPRO->CN9_DTFIM),; 
            sTod(QRYPRO->CN9_DTASSI),;  
            QRYPRO->CN9_VIGE,;  
            QRYPRO->CN9_UNVIGE,;
            QRYPRO->CN1_DESCRI,; 
            QRYPRO->CN9_SITUAC,;
            QRYPRO->CNB_REVISA,; 
            sTod(QRYPRO->CN9_DTREV),;
            "R$"+TRANSFORM(QRYPRO->CN9_VLINI, "@E 999,999,999.99"),;
            "R$"+TRANSFORM(QRYPRO->CN9_VLATU, "@E 999,999,999.99"),;
            "R$"+TRANSFORM(QRYPRO->CN9_SALDO,  "@E 999,999,999.99")})
        //Pulando Registro
        QRYPRO->(DbSkip())
        endif   
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
