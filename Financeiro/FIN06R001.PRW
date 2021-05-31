//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"


/*/{Protheus.doc} FIN06R001
Função que cria um exemplo de FWMsExcel
@author Leandro Souza
@since 14/12/2020
@version 1.0
    @example
    Financeiro - MOD06 - SIGAFIN - Relat�rio de Fornecedores x Dados Bancarios
    u_FIN06R001()
/*/

User Function FIN06R001()

    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'FIN06R001.xml'
    
//************************************************************************************************
    cPerg    := "FIN06R002"
     
    cValid   := ""
    cF3      := ""
    cPicture := ""
    cDef01   := ""
    cDef02   := ""
    cDef03   := ""
    cDef04   := ""
    cDef05   := ""
     
    u_zPutSX1(cPerg, "01", "De Cod Fornecedor?",         "MV_PAR01", "MV_CH0", "C", TamSX3('A2_COD')[01],    0, "G", cValid,       "SA2", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Cod do Fornecedor")
    u_zPutSX1(cPerg, "02", "At� Cod Fornecedor?",        "MV_PAR02", "MV_CH1", "C", TamSX3('A2_COD')[01],    0, "G", cValid,       "SA2", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Cod do Fornecedor")
    u_zPutSX1(cPerg, "03", "Banco:",                     "MV_PAR03", "MV_CH2", "C", TamSX3('A2_BANCO')[01],  0, "G", cValid,       "SA2BAN", cPicture,        cDef01,  cDef02,        cDef03,        cDef04,    cDef05, "Informe o Cod do Banco")

    If ! Pergunte("FIN06R002",.T.)

        Return

    EndIf
    
     
    //Pegando os dados
    cQuery := " SELECT "
    cQuery += "     SA2.A2_COD, "
    cQuery += "     SA2.A2_NOME, "
    cQuery += "     SA2.A2_LOJA, " 
    cQuery += "     SA2.A2_CGC, "   
    cQuery += "     SA2.A2_END, " 
    cQuery += "     SA2.A2_CONTPRE, "   
    cQuery += "     SA2.A2_COMPLEM, "     
    cQuery += "     SA2.A2_BAIRRO, "    
    cQuery += "     SA2.A2_CEP, "    
    cQuery += "     SA2.A2_EST, "
    cQuery += "     SA2.A2_ESTADO, "
    cQuery += "     SA2.A2_COD_MUN, "
    cQuery += "     SA2.A2_MUN, "
    cQuery += "     SA2.A2_PAIS, "    
    cQuery += "     SA2.A2_DDD, "    
    cQuery += "     SA2.A2_TEL, "    
    cQuery += "     SA2.A2_EMAIL, "
    cQuery += "     SA2.A2_BANCO, "    
    cQuery += "     SA2.A2_AGENCIA, "
    cQuery += "     SA2.A2_DVAGE, "    
    cQuery += "     SA2.A2_NUMCON, "
    cQuery += "     SA2.A2_DVCTA, "
    cQuery += "     SA2.A2_CONTA"
    cQuery += " FROM "
    cQuery += "     "+RetSQLName('SA2')+" SA2 "
    cQuery += " WHERE "
    cQuery += "     SA2.D_E_L_E_T_ = ''  "
    cQuery += " AND SA2.A2_BANCO <> '' "
    cQuery += IF ((EMPTY (MV_PAR01)), "", " AND A2_COD >= '" + MV_PAR01 + "' ")
    cQuery += IF ((EMPTY (MV_PAR02)), "", " AND A2_COD <= '" + MV_PAR02 + "' ")
    cQuery += IF ((EMPTY (MV_PAR03)), "", " AND A2_BANCO = '" + MV_PAR03 + "' ")    
    
    cQuery += " ORDER BY "
    cQuery += "     SA2.A2_COD "
    TCQuery cQuery New Alias "QRYPRO"

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New("FIN06R001")

    oFWMsExcel:AddworkSheet("Fornecedor_x_Banco")
    //Criando a Tabela
    oFWMsExcel:AddTable("Fornecedor_x_Banco","Fornecedor_x_Banco")
    //criando colunas
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Cod "         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Nome "         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Loja "         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","CNPJ "         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Endereco "         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Cont "         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Complemento"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Bairro"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","CEP"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Est"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Estado"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Cod Mun"         ,1)
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Municipio"         ,1)
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Pais"         ,1)
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","DDD"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Tel"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Email"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Banco"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Agencia"         ,1) 
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","DV AG"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Conta"         ,1) 
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","DV Conta"         ,1)  
    oFWMsExcel:AddColumn("Fornecedor_x_Banco","Fornecedor_x_Banco","Comp Conta"         ,1)   
    //Criando as Linhas... Enquanto não for fim da query
    While !(QRYPRO->(EoF()))
        oFWMsExcel:AddRow("Fornecedor_x_Banco","Fornecedor_x_Banco",{QRYPRO->A2_COD,; 
            QRYPRO->A2_NOME,; 
            QRYPRO->A2_LOJA,;
            QRYPRO->A2_CGC,;
            QRYPRO->A2_END,;  
            QRYPRO->A2_CONTPRE,;
            QRYPRO->A2_COMPLEM,; 
            QRYPRO->A2_BAIRRO,;
            QRYPRO->A2_CEP,; 
            QRYPRO->A2_EST,;  
            QRYPRO->A2_ESTADO,;
            QRYPRO->A2_COD_MUN,; 
            QRYPRO->A2_MUN,;
            QRYPRO->A2_PAIS,; 
            QRYPRO->A2_DDD,; 
            QRYPRO->A2_TEL,; 
            QRYPRO->A2_EMAIL,; 
            QRYPRO->A2_BANCO,; 
            QRYPRO->A2_AGENCIA,; 
            QRYPRO->A2_DVAGE,; 
            QRYPRO->A2_NUMCON,; 
            QRYPRO->A2_DVCTA,; 
            QRYPRO->A2_CONTA})

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
