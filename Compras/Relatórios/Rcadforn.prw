//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} Rcadforn01
FunÃ§Ã£o que cria um exemplo de FWMsExcel
@author Leandro Souza
@since 28/05/2021
@version 1.0
    @example
    Relatório Módulo de Compras - MOD02 - SIGACOM - Relatório de Usuários Responsáveis pelo Cadastro / Aprovação de FORNECEDOR
    u_Rcadforn()
/*/

User Function Rcadforn()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'RCadforn.xml'
//************************************************************************************************
    cPerg    := "Rcadforn01"
     
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
  
    u_zPutSX1(cPerg, "03", "Codigo do Produto?",    "MV_PAR03", "MV_CH2",  "C", TamSX3('A2_COD')[01], 0, "G", cValid,   "SA2",   cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o codigo do produto ")
    
    
    If ! Pergunte("Rcadforn01",.T.)
        Return

    EndIf

    //Pegando os dados
    cQuery := " SELECT "
    cQuery += "     SA2.A2_COD,"
    cQuery += "     SA2.A2_CGC,"
    cQuery += "     SA2.A2_LOJA,"
    cQuery += "     SA2.A2_NOME,"
    cQuery += "     SA2.A2_END,"
    cQuery += "     SA2.A2_COMPLEM,"
    cQuery += "     SA2.A2_ENDCOMP,"
    cQuery += "     SA2.A2_BAIRRO,"
    cQuery += "     SA2.A2_CEP,"
    cQuery += "     SA2.A2_EST,"
    cQuery += "     SA2.A2_MUN,"
    cQuery += "     SA2.A2_DDD,"
    cQuery += "     SA2.A2_TEL,"
    cQuery += "     SA2.A2_EMAIL,"
    cQuery += "     SA2.A2_BANCO,"
    cQuery += "     SA2.A2_AGENCIA,"
    cQuery += "     SA2.A2_DVAGE,"
    cQuery += "     SA2.A2_NUMCON,"
    cQuery += "     SA2.A2_DVCTA,"
    cQuery += "     F72.F72_TPCHV,"                                  
    cQuery += "     F72.F72_CHVPIX,"                                  
    cQuery += "     SA2.A2_USERCAD,"                                  
    cQuery += "     SA2.A2_DTCAD,"                                  
    cQuery += "     SA2.A2_LIBERA" 

    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SB1')+" SB1 "
    cQuery += " LEFT JOIN " +RetSQLName('F72')+" F72  ON A2.A2_COD = F72.F72_COD "
    cQuery += " WHERE "
    cQuery += "     SB1.D_E_L_E_T_ = '' "

    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND A2_DTCAD >= '" + DTOS(MV_PAR01) + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND A2_DTCAD <= '" + DTOS(MV_PAR02) + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND A2_COD = '" + MV_PAR03 + "' ")
 
    
    cQuery += " ORDER BY "
    cQuery += "     SA2.A2_DTCAD, SA2.A2_USERCAD "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irÃ¡ gerar o conteÃºdo do Excel
    oFWMsExcel := FWMSExcel():New("Rcadforn01")

   
    oFWMsExcel:AddworkSheet("Cad_Fornecedor")
    //Criando a Tabela
    oFWMsExcel:AddTable("Cad_Fornecedor","Cad_Fornecedor   ")
    //criando colunas
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Código"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","CNPJ"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Loja"       ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Nome"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","End"         ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Compl.1"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Compl.2"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Bairror"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","CEP"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Estado"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Municipio"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","DDD"       ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Tel"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","E-mail"         ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Bancoo"         ,1)  
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Agência"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","DV Ag"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Conta"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","DV Conta"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Tipo Chave Pix"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Chave Pix"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Cadastrador"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Data do Cadastro"  ,1)
    oFWMsExcel:AddColumn("Cad_Fornecedor   ","Cad_Fornecedor  ","Aprovador"  ,1)


    //Criando as Linhas... Enquanto nÃ£o for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Cad_Fornecedor  ","Cad_Fornecedor ",{QRYPRO->A2_COD,; 
            QRYPRO->A2_CGC,;
            QRYPRO->A2_LOJA,;
            QRYPRO->A2_NOME,;
            QRYPRO->A2_END,;
            QRYPRO->A2_COMPLEM,;
            QRYPRO->A2_ENDCOMP,;
            QRYPRO->A2_BAIRRO,;
            QRYPRO->A2_CEP,;
            QRYPRO->A2_EST,;
            QRYPRO->A2_MUN,;
            QRYPRO->A2_DDD,;
            QRYPRO->A2_TEL,;
            QRYPRO->A2_EMAIL,;
            QRYPRO->A2_BANCO,;
            QRYPRO->A2_AGENCIA,;
            QRYPRO->A2_DVAGE,;
            QRYPRO->A2_NUMCON,;
            QRYPRO->A2_DVCTA,;
            QRYPRO->F72_TPCHV,;
            QRYPRO->F72_CHVPIX,;
            QRYPRO->A2_USERCAD,;
            stod( QRYPRO->A2_DTCAD),;
            QRYPRO->A2_LIBERA})

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
