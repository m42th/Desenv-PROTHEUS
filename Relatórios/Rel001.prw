#include 'protheus.ch'
#include 'totvs.ch'

/*========================================================================
|Programa:| Relatório para impressão em excel com os dados digitado 
diretamente dentro das colunas e linhas.             
--------------------------------------------------------------------------
|Data:    | 06/11/2020
--------------------------------------------------------------------------
|Autor:   | Matheus Santos de Oliveira / (61) 99812-9318 / Águas Lindas-GO
==========================================================================*/

User function REL001()

    MSAGUARDE({||fProcessa()},"Relatório Excel","Gerando Arquivo...")

Return

//FUNÇÃO ESTATICA ONDE É CRIADO O CORPO DO RELATÓRIO E ALIMENTADO COM OS DADOS QUE SAIRÃO NA IMPRESSÃO
Static Function fProcessa()

    //MÉTODO CONSTRUTOR 
    Local aArea  := GetArea()
    Local oExcel := FWMsExcel():New()
    Local cArq   := GetTempPath()+'Excel.xml'

    //ABA 1
    Local cAba1 := "Aba 1"
    Local cTab1 := "Planilha 1"
    //ABA 2
    Local cAba2 := "Aba 2"
    Local cTab2 := "Planilha 2"
    
    //CONSTRUINDO CORPO DO RELATÓRIO ABA 1
    oExcel:AddWorkSheet(cAba1)
    oExcel:AddTable(cAba1,cTab1)
    //COLUNAS
    oExcel:AddColumn(cAba1,cTab1,"COLUNA 1",1,1)
    oExcel:AddColumn(cAba1,cTab1,"COLUNA 2",1,1)
    oExcel:AddColumn(cAba1,cTab1,"COLUNA 3",1,1)
    oExcel:AddColumn(cAba1,cTab1,"COLUNA 4",1,1)
    oExcel:AddColumn(cAba1,cTab1,"COLUNA 5",1,1)
    //LINHAS
    oExcel:AddRow(cAba1,cTab1,{10,20,30,40,STOD("20201106")})
    oExcel:AddRow(cAba1,cTab1,{11,21,31,41,STOD("20201107")})
    oExcel:AddRow(cAba1,cTab1,{12,22,32,42,STOD("20201108")})
    oExcel:AddRow(cAba1,cTab1,{13,23,33,43,STOD("20201109")})
    oExcel:AddRow(cAba1,cTab1,{14,25,36,44,STOD("20201110")})

    //CONSTRUINDO CORPO DO RELATÓRIO ABA 2
    oExcel:AddWorkSheet(cAba2)
    oExcel:AddTable(cAba2,cTab2)
    //COLUNAS
    oExcel:AddColumn(cAba2,cTab2,"COLUNA 1",1,1) 
    oExcel:AddColumn(cAba2,cTab2,"COLUNA 2",1,1)  
    oExcel:AddColumn(cAba2,cTab2,"COLUNA 3",1,1)
    oExcel:AddColumn(cAba2,cTab2,"COLUNA 4",1,1)
    oExcel:AddColumn(cAba2,cTab2,"COLUNA 5",1,1)
    //LINHAS
    oExcel:AddRow(cAba2,cTab2,{"MATHEUS",25,0000000,987654321,"RUA 31 LT 32"})
    oExcel:AddRow(cAba2,cTab2,{"MURILLO",07,0000001,987654321,"RUA 32 LT 33"})

    //ATIVANDO E CRIANDO O ARQUIVO
    oExcel:Activate()
    oExcel:GetXMLFile(cArq)

    //ABRINDO O EXCEL E ABRINDO O ARQUIVO
    oExcel := MsExcel():New()            //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArq)         //Abre uma planilha
    oExcel:SetVisible(.T.)             //Visualiza a planilha
    oExcel:Destroy()                  //Encerra o processo do gerenciador de tarefas

    RestArea(aArea)

Return
