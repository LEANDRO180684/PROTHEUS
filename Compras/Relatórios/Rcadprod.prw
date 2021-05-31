//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} GctR001
FunÃ§Ã£o que cria um exemplo de FWMsExcel
@author Leandro Souza
@since 12/05/2021
@version 1.0
    @example
    Relatório Módulo de Compras - MOD02 - SIGACOM - Relatório de Usuários Responsáveis pelo Cadastro / Aprovação de Produtos
    u_Rcadprod()
/*/

User Function Rcadprod()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'RCadprod.xml'
//************************************************************************************************
    cPerg    := "Rcadprod04"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "De Data inclusao?",        "MV_PAR01", "MV_CH0",  "D", 08,                       0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de inclusao do Produto ")
    u_zPutSX1(cPerg, "02", "Ate Data?",                "MV_PAR02", "MV_CH1",  "D", 08,                       0, "G", cValid,       cF3,        cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe a data de inclusao da Produto ")
  
    u_zPutSX1(cPerg, "03", "Codigo do Produto?",    "MV_PAR03", "MV_CH2",  "C", TamSX3('B1_COD')[01], 0, "G", cValid,   "SB1",   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o codigo do produto ")
    
    
    If ! Pergunte("Rcadprod04",.T.)
        Return

    EndIf

    //Pegando os dados
    cQuery := " SELECT "
    cQuery += "     SB1.B1_COD,"   
    cQuery += "     SB1.B1_DESC,"    
    cQuery += "     SB1.B1_TIPO,"     
    cQuery += "     SB1.B1_UM,"    
    cQuery += "     SB1.B1_POSIPI,"
    cQuery += "     SB1.B1_MSBLQL," 
    cQuery += "     SB1.B1_USERCAD,"    
    cQuery += "     SB1.B1_LIBERA,"    
    cQuery += "     SB1.B1_DTCAD" 

    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SB1')+" SB1 "
    cQuery += " WHERE "
    cQuery += "     SB1.D_E_L_E_T_ = '' "

    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND B1_DTCAD >= '" + DTOS(MV_PAR01) + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND B1_DTCAD <= '" + DTOS(MV_PAR02) + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND B1_COD = '" + MV_PAR03 + "' ")
 
    
    cQuery += " ORDER BY "
    cQuery += "     SB1.B1_DTCAD, SB1.B1_USERCAD "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irÃ¡ gerar o conteÃºdo do Excel
    oFWMsExcel := FWMSExcel():New("Rcadpro01")

   
    oFWMsExcel:AddworkSheet("Cadastros")
    //Criando a Tabela
    oFWMsExcel:AddTable("Cadastros","Cadastros")
    //criando colunas
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Código"         ,1)  
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Descrição"         ,1)  
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Tipo"       ,1)  
    oFWMsExcel:AddColumn("Cadastros","Cadastros","UM"         ,1)  
    oFWMsExcel:AddColumn("Cadastros","Cadastros","POSIPI (NCM)"         ,1)
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Aprovação"         ,1)  
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Cadastrador"  ,1)
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Aprovador"  ,1)
    oFWMsExcel:AddColumn("Cadastros","Cadastros","Data do Cadastro"  ,1)
  

    //Criando as Linhas... Enquanto nÃ£o for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Cadastros","Cadastros",{QRYPRO->B1_COD,; 
            QRYPRO->B1_DESC,;
            QRYPRO->B1_TIPO,;  
            QRYPRO->B1_UM,;  
            QRYPRO->B1_POSIPI,;
            QRYPRO->B1_MSBLQL,;  
            QRYPRO->B1_USERCAD,;
            QRYPRO->B1_LIBERA,;
           stod( QRYPRO->B1_DTCAD)})

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
