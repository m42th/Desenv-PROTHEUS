#include 'protheus.ch'
#include 'topconn.ch'

Static EOL := CHR(13) + CHR(10)

/*========================================================================
|Programa:| Relatório para impressão em excel com os dados recebidos do 
SQL Server, e totalizando valores de coluna.
--------------------------------------------------------------------------
|Data:    | 08/11/2020
--------------------------------------------------------------------------
|Autor:   | Matheus Santos de Oliveira / (61) 99812-9318 / Águas Lindas-GO
==========================================================================*/

User Function REL001()

    If MsgYesNo("Imprimir relatório?","Atenção!")
        MSAguarde({|| fProcessa()}, "Relatório Excel!", "Aguarde, gerando arquivo...")
    Else 
        Alert("Operação cancelada pelo usuário!")
        Return
    EndIf
    
Return

Static Function fProcessa()

    Local aArea := GetArea()
	Local cArq  := GetTempPath() + 'Excel.xml'
    Local oExcel := FWMSExcel():New()

	local cTab1 := "Tabela 1"
	Local cAba1 := "Aba 1"
	Local cTab2 := 'Tabela 2'
	Local cAba2 := "Aba 2"

    //Aba 1
	oExcel:AddWorkSheet(cAba1)
	oExcel:AddTable(cAba1,cTab1)

	oExcel:AddColumn(cAba1,cTab1,"COD FORNECEDOR",1,1,.F.)
	oExcel:AddColumn(cAba1,cTab1,"NOME FORNECEDOR",1,1,.F.)
	oExcel:AddColumn(cAba1,cTab1,"COD PRODUTO",1,1,.F.)
	oExcel:AddColumn(cAba1,cTab1,"NOME PRODUTO",1,1,.F.)
	oExcel:AddColumn(cAba1,cTab1,"PREÇO COMPRA",1,1,.T.)

	fDadosSQL()

	While !TMP1->(EOF())
		oExcel:AddRow(cAba1,cTab1,{TMP1->COD_FORNECEDOR,;
			TMP1->FORNECEDOR,;
			TMP1->COD_PRODUTO,;
			TMP1->NOME_PRODUTO,;
			TMP1->PRECO_COMPRA })
		TMP1->(DbSkip())
	End

	TMP1->(DbCloseArea())

    //Aba 2
    oExcel:AddWorkSheet(cAba2)
    oExcel:AddTable(cAba2,cTab2)

    oExcel:AddColumn(cAba2,cTab2,"Data hoje",1,4)
    oExcel:AddColumn(cAba2,cTab2,"Data + 7 Dias",1,4)

    oExcel:AddRow(cAba2,cTab2,{Date(),DaySum(Date(),7)})

    //Ativando a função 
    oExcel:Activate()
    oExcel:GetXMLFile(cArq)

    //ABRINDO O EXCEL E ABRINDO O ARQUIVO
    oExcel := MsExcel():New()            //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArq)         //Abre uma planilha
    oExcel:SetVisible(.T.)             //Visualiza a planilha
    oExcel:Destroy()                  //Encerra o processo do gerenciador de tarefas

    RestArea(aArea)
Return

Static Function fDadosSQL()

	Local cQry := ""
	Local TMP1 := ""

	TMP1 := GetNextAlias()
	cQry += EOL + "SELECT A5_FORNECE COD_FORNECEDOR, A5_NOMEFOR FORNECEDOR, B1_COD COD_PRODUTO, A5_NOMPROD NOME_PRODUTO, AIB_PRCCOM PRECO_COMPRA "
	cQry += EOL + "FROM SB1990 SB1 "
	cQry += EOL + "INNER JOIN SA5990 SA5 ON A5_PRODUTO = B1_COD AND SA5.D_E_L_E_T_='' "
	cQry += EOL + "INNER JOIN AIB990 AIB ON AIB_CODPRO = B1_COD AND AIB.D_E_L_E_T_='' "
	cQry += EOL + "WHERE SB1.D_E_L_E_T_='' "

    If Select("TMP1") > 0
        DbSelectArea("TMP1")
        DbCloseArea("TMP1")
    EndIf

	TCQuery cQry New Alias "TMP1"

Return
