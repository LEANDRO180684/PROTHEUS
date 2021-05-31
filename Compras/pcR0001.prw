//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} pcR001
Função que cria um exemplo de FWMsExcel
@author Leandro Souza
@since 12/01/2021
@version 1.0
    @example
    Compras - MOD02 - SIGACOM - Relat�rio de Pedido de Compras Modificado
    u_pcR001()
/*/

User Function pcR001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'pcR001.xml'
//************************************************************************************************
    cPerg    := "pcR014"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "De  Pedido?",          "MV_PAR01", "MV_CH0", "C", TamSX3('C7_NUM')[01],       0, "G", cValid,       "SC7",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero do Pedido")
    u_zPutSX1(cPerg, "02", "Ate Pedido?",          "MV_PAR02", "MV_CH1", "C", TamSX3('C7_NUM')[01],       0, "G", cValid,       "SC7",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero de Pedido")
    u_zPutSX1(cPerg, "03", "De  SC?",              "MV_PAR03", "MV_CH2", "C", TamSX3('C1_NUM')[01],       0, "G", cValid,       "SC1",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero do Pedido")
    u_zPutSX1(cPerg, "04", "Ate SC?",              "MV_PAR04", "MV_CH3", "C", TamSX3('C1_NUM')[01],       0, "G", cValid,       "SC1",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Numero de Pedido")
    u_zPutSX1(cPerg, "05", "De Data de Emissao:",  "MV_PAR05", "MV_CH4", "D", 08,                         0, "G", cValid,       cF3,      cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a Data de Emissao")
    u_zPutSX1(cPerg, "06", "Ate Data de Emissao:", "MV_PAR06", "MV_CH5", "D", 08,                         0, "G", cValid,       cF3,      cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a Data de Emissao")
    u_zPutSX1(cPerg, "07", "Filial:",              "MV_PAR07", "MV_CH6", "C", TamSX3('C7_FILIAL')[01],    0, "G", cValid,       cF3,      cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a Filial")
    u_zPutSX1(cPerg, "08", "Fornecedor:",          "MV_PAR08", "MV_CH7", "C", TamSX3('A2_COD')[01],       0, "G", cValid,       "SA2",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Fornecedor")
    u_zPutSX1(cPerg, "09", "Loja:",                "MV_PAR09", "MV_CH8", "C", TamSX3('A2_LOJA')[01],      0, "G", cValid,       cF3,      cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Fornecedor")
    u_zPutSX1(cPerg, "10", "C.C - De:",            "MV_PAR10", "MV_CH9", "C", TamSX3('C7_CC')[01],        0, "G", cValid,       "CTT",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Centro de Custo De:")
    u_zPutSX1(cPerg, "11", "C.C - Ate:",           "MV_PAR11", "MV_CH10", "C", TamSX3('C7_CC')[01],       0, "G", cValid,       "CTT",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Centro de Custo Ate:")
    u_zPutSX1(cPerg, "12", "Comprador:",           "MV_PAR12", "MV_CH11", "C", TamSX3('Y1_COD')[01],      0, "G", cValid,       "SY1",    cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Comprador:")

    If ! Pergunte("pcR014",.T.)

        Return

    EndIf
    
    
    
    //Pegando os dados
    cQuery := " SELECT DISTINCT"
    cQuery += "     C7.C7_FILIAL,"
    cQuery += "     C7.C7_NUM,"
    cQuery += "     C7.C7_EMISSAO," 
    cQuery += "     C7.C7_FORNECE," 
    cQuery += "     A2.A2_NOME,"  
    cQuery += "     C7.C7_ITEM," 
    cQuery += "     C7.C7_PRODUTO,"   
    cQuery += "     C7.C7_DESCRI,"         

    cQuery += "     CASE C7.C7_APLICA WHEN '1' THEN 'CAPEX' "
    cQuery += "     WHEN '2' THEN 'OPEX'"
    cQuery += "     WHEN '3' THEN 'SGA'"
    cQuery += "     WHEN '4' THEN 'ESTOQUE'"
    cQuery += "     END AS 'C7_APLICA',"
    
    cQuery += "     C7.C7_QUANT," 
    cQuery += "     C7.C7_UM,"    
    cQuery += "     C7.C7_PRECO,"
    cQuery += "     C7.C7_TOTAL,"
    
    cQuery += "     C7.C7_NUMSC,"
    cQuery += "     C1.C1_SOLICIT,"

    cQuery += "     C7.C7_ITEMSC," 
    cQuery += "     C7.C7_DATPRF,"    
    cQuery += "     C7.C7_QUJE,"
    cQuery += "     C7.C7_FILENT,"
    cQuery += "     C7.C7_LOCAL,"
    cQuery += "     C7.C7_OBS," 
    cQuery += "     C7.C7_CC,"    
    cQuery += "     C7.C7_CONTA,"

    cQuery += "     C7.C7_APROV,"
    cQuery += "     Y1.Y1_COD,"
    cQuery += "     Y1.Y1_NOME,"
    cQuery += "     C7.C7_CONTRA," 
    cQuery += "     C7.C7_MEDICAO,"    
    cQuery += "     C7.C7_ITEMED "
        
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SC7')+" C7 "

    cQuery += " LEFT JOIN " +RetSQLName('SC1')+" C1  ON C7.C7_NUM = C1.C1_PEDIDO "
    cQuery += " LEFT JOIN " +RetSQLName('SY1')+" Y1  ON C7.C7_COMPRA = Y1.Y1_COD "
    cQuery += " LEFT JOIN " +RetSQLName('SA2')+" A2  ON C7.C7_FORNECE = A2.A2_COD "

    cQuery += " WHERE "
    cQuery += "     C7.D_E_L_E_T_ = '' "

    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND C7_NUM >= '" + MV_PAR01 + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND C7_NUM <= '" + MV_PAR02 + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND C7_NUMSC >= '" + MV_PAR03 + "' ")
    cQuery += IF ((EMPTY (MV_PAR04)), "", " AND C7_NUMSC <= '" + MV_PAR04 + "' ")
    cQuery += IF ((EMPTY (MV_PAR05)), "", " AND C7_EMISSAO >= '" + DTOS(MV_PAR05) + "' ")
    cQuery += IF ((EMPTY (MV_PAR06)), "", " AND C7_EMISSAO <= '" + DTOS(MV_PAR06) + "' ")
    cQuery += IF ((EMPTY (MV_PAR07)), "", " AND C7_FILIAL = '" + MV_PAR07 + "' ")
    cQuery += IF ((EMPTY (MV_PAR08)), "", " AND C7_FORNECE = '" + MV_PAR08 + "' ") 
    cQuery += IF ((EMPTY (MV_PAR09)), "", " AND C7_LOJA = '" + MV_PAR09 + "' ") 
    cQuery += IF ((EMPTY (MV_PAR10)), "", " AND C7_CC >= '" + MV_PAR10 + "' ")
    cQuery += IF ((EMPTY (MV_PAR11)), "", " AND C7_CC <= '" + MV_PAR11 + "' ")
    cQuery += IF ((EMPTY (MV_PAR12)), "", " AND Y1_COD = '" + MV_PAR12 + "' ")

    cQuery += " ORDER BY "
    cQuery += "     C7.C7_FILIAL, C7.C7_NUM, C7.C7_ITEM, C7.C7_NUMSC, C7.C7_CC, Y1.Y1_NOME "

    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New("pcR001")

   
    oFWMsExcel:AddworkSheet("Relat.Pd.Compras")
    //Criando a Tabela
    oFWMsExcel:AddTable("Relat.Pd.Compras","Relat.Pd.Compras")
    //criando colunas
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Filial"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Num. Pedido"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Dt Emissao"       ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Cod. Fornecedor"         ,1) 
     oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Fornecedor"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Item Pedido"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Cod. Produto"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Descri Produto"       ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Aplicacao"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Qtd"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","UM"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Pre�o"       ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Total"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Num SC"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Solicitante"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Item SC"       ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Dt Entrega"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Qtd Entregue"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Filial Entrada"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Local"       ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Obs"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","CC"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Conta"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Gr. Aprov"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Cod Comprador"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Nome Comprador"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Num. Contrato"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Medicao"         ,1)  
    oFWMsExcel:AddColumn("Relat.Pd.Compras","Relat.Pd.Compras","Item Med"         ,1)  

    //Criando as Linhas... Enquanto não for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Relat.Pd.Compras","Relat.Pd.Compras",{QRYPRO->C7_FILIAL,; 
            QRYPRO->C7_NUM,; 
            sTod(QRYPRO->C7_EMISSAO),;
            QRYPRO->C7_FORNECE,; 
            QRYPRO->A2_NOME,;
            QRYPRO->C7_ITEM,; 
            QRYPRO->C7_PRODUTO,; 
            QRYPRO->C7_DESCRI,; 
            QRYPRO->C7_APLICA,; 
            QRYPRO->C7_QUANT,; 
            QRYPRO->C7_UM,; 
            QRYPRO->C7_PRECO,; 
            QRYPRO->C7_TOTAL,; 
            QRYPRO->C7_NUMSC,; 
            QRYPRO->C1_SOLICIT,; 
            QRYPRO->C7_ITEMSC,; 
            sTod(QRYPRO->C7_DATPRF),;
            QRYPRO->C7_QUJE,;
            QRYPRO->C7_FILENT,;
            QRYPRO->C7_LOCAL,;
            QRYPRO->C7_OBS,;
            QRYPRO->C7_CC,;
            QRYPRO->C7_CONTA,;
            QRYPRO->C7_APROV,;
            QRYPRO->Y1_COD,;
            QRYPRO->Y1_NOME,;
            QRYPRO->C7_CONTRA,;
            QRYPRO->C7_MEDICAO,;
            QRYPRO->C7_ITEMED})

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
