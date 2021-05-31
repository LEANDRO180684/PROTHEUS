//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} EstR001
Função que cria um exemplo de FWMsExcel
@author Leandro Bittencourt
@since 22/10/2020
@version 2.0
    @example
    u_EstR001()
/*/

User Function EstR001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'EstR001.xml'
//************************************************************************************************
    cPerg    := "EstR004"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "De Produto?",         "MV_PAR01", "MV_CH0", "C", TamSX3('B1_COD')[01], 0, "G", cValid,       "SB1", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Produto inicial")
    u_zPutSX1(cPerg, "02", "Ate Produto?",        "MV_PAR02", "MV_CH1", "C", TamSX3('B1_COD')[01], 0, "G", cValid,       "SB1", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Produto final")
    u_zPutSX1(cPerg, "03", "De Data?",            "MV_PAR03", "MV_CH2", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data inicial a ser considerada")
    u_zPutSX1(cPerg, "04", "Ate Data?",           "MV_PAR04", "MV_CH3", "D", 08,                   0, "G", cValid,       cF3,   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data Final a ser considerada")
    
    If ! Pergunte("EstR004",.T.)
        Return
    EndIf

    //Pegando os dados
    cQuery := " SELECT DISTINCT"
    cQuery += "     B1.B1_COD,"   
    cQuery += "     B1.B1_DESC,"    
    cQuery += "     B1.B1_UM,"   
    cQuery += "     B1.B1_TIPO,"    
    cQuery += "     B1.B1_GRUPO," 
    cQuery += "     B2.B2_CM1,"    
    cQuery += "     B2.B2_QFIM,"    
    cQuery += "     B2.B2_VFIM1,"
    cQuery += "     B1.B1_POSIPI,"
    cQuery += "     B2.B2_LOCALIZ,"
    cQuery += "     B2.B2_DMOV,"
    cQuery += "     D1.D1_LOCAL,"
    cQuery += "     D1.D1_TES,"
    cQuery += "     D1.D1_CF,"
    cQuery += "     D3.D3_NUMSEQ,"
    cQuery += "     D3.D3_NUMSA,"
    cQuery += "     D1.D1_QUANT,"
    cQuery += "     D1.D1_TOTAL,"
    cQuery += "     D1.D1_VUNIT,"
    cQuery += "     D3.D3_QUANT,"
    cQuery += "     D3.D3_CUSTO1,"    
    cQuery += "     B2.B2_QATU," 
    cQuery += "     B2.B2_VATU1,"   
    cQuery += "     D3.D3_CC"
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SB1')+" B1 "
    cQuery += " INNER JOIN " +RetSQLName('SD3')+" D3  ON B1.B1_COD = D3.D3_COD AND D3.D_E_L_E_T_ = ''"
    cQuery += " INNER JOIN " +RetSQLName('SB2')+" B2  ON B1.B1_COD = B2.B2_COD AND B2.D_E_L_E_T_ = ''"
    cQuery += " INNER JOIN " +RetSQLName('SD1')+" D1  ON B1.B1_COD = D1.D1_COD AND D1.D_E_L_E_T_ = ''"
    cQuery += " WHERE "
    cQuery += "     B1.D_E_L_E_T_ = '' "
    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND B1_COD >= '"  + (MV_PAR01)  + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND B1_COD <= '"  + (MV_PAR02)  + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND B2_DMOV >= '" + DTOS(MV_PAR03) + "' ")
    cQuery += IF ((EMPTY (MV_PAR04)), "", " AND B2_DMOV <= '" + DTOS(MV_PAR04) + "' ")

    cQuery += " ORDER BY "
    cQuery += "     B1.B1_COD "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New("ESTR001")

   
    oFWMsExcel:AddworkSheet("Kardex")
    //Criando a Tabela
    oFWMsExcel:AddTable("Kardex","Kardex")
    //criando colunas
    oFWMsExcel:AddColumn("Kardex","Kardex","Codigo"         ,1)  
    oFWMsExcel:AddColumn("Kardex","Kardex","Descricao"         ,1)  
    oFWMsExcel:AddColumn("Kardex","Kardex","UM"       ,1)  
    oFWMsExcel:AddColumn("Kardex","Kardex","Tipo"         ,1)  
    oFWMsExcel:AddColumn("Kardex","Kardex","Grupo"         ,1)  
    oFWMsExcel:AddColumn("Kardex","Kardex","Custo Medio"  ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","Qtd. Saldo"  ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","Vlr. Total Saldo",1)   
    oFWMsExcel:AddColumn("Kardex","Kardex","Posicao IPI"     ,1) 
    oFWMsExcel:AddColumn("Kardex","Kardex","Endereco"     ,1) 
    oFWMsExcel:AddColumn("Kardex","Kardex","Operacao Data"     ,1) 
    oFWMsExcel:AddColumn("Kardex","Kardex","ARM"     ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","TES"     ,1)    
    oFWMsExcel:AddColumn("Kardex","Kardex","C.F"     ,1)    
    oFWMsExcel:AddColumn("Kardex","Kardex","DOC NUMERO"     ,1)    
    oFWMsExcel:AddColumn("Kardex","Kardex","SA"     ,1)                                                         
    oFWMsExcel:AddColumn("Kardex","Kardex","ENTRADAS QTDE"     ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","ENTRADAS CUSTO TOTAL"     ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","CUSTO MEDIO MOVIMENTO"     ,1)   
    oFWMsExcel:AddColumn("Kardex","Kardex","SAIDAS QUANTIDADE"       ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","SAIDAS CUSTO TOTAL"       ,1) 
    oFWMsExcel:AddColumn("Kardex","Kardex","SALDO QTDE"       ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","SALDO VALOR TOTAL"       ,1)
    oFWMsExcel:AddColumn("Kardex","Kardex","CLI, FOR, CC, PJ ou OP"       ,1)    


    //Criando as Linhas... Enquanto não for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Kardex","Kardex",{QRYPRO->B1_COD,;             
            QRYPRO->B1_DESC,; 
            QRYPRO->B1_UM,;
            QRYPRO->B1_TIPO,;  
            QRYPRO->B1_GRUPO,;  
            QRYPRO->B2_CM1,;  
            QRYPRO->B2_QFIM,;  
            QRYPRO->B2_VFIM1,;
            QRYPRO->B1_POSIPI,;  
            QRYPRO->B2_LOCALIZ,;
            sTod(QRYPRO->B2_DMOV),;
            QRYPRO->D1_LOCAL,;
            QRYPRO->D1_TES,;
            QRYPRO->D1_CF,;
            QRYPRO->D3_NUMSEQ,;
            QRYPRO->D3_NUMSA,;
            QRYPRO->D1_QUANT,;
            QRYPRO->D1_TOTAL,;
            QRYPRO->D1_VUNIT,; 
            QRYPRO->D3_QUANT,;
            QRYPRO->D3_CUSTO1,;
            QRYPRO->B2_QATU,;
            QRYPRO->B2_VATU1,;
            QRYPRO->D3_CC})

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
