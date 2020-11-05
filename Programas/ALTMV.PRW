#include 'protheus.ch'
#include 'totvs.ch'

/*========================================================================
|Programa:  Programa escrito com a finalidade de procurar por um parâmetro
e poder alterar o valor do mesmo sem precisar logar no CFG.              |
--------------------------------------------------------------------------
|Data:    | 04/11/2020
--------------------------------------------------------------------------
|Autor:   | Matheus Santos de Oliveira / (61) 99812-9318
==========================================================================*/

User Function ALTMV()

    //Construção e funções da tela
    Local oDlg
    Local oTsay1
    Local oTsay2
    Local oButt1 
    Local oButt2

    //Variáveis para manipulação dos dados
    Local oGet1
    Local oGet2
    Local cParGet1 := space(10)
    Local cParGet2 := space(10)
    Local cParVlr1
    Local cParVlr2
   

    //CRIAÇÃO DA TELA PRINCIPAL   
    oDlg := MSDialog():New(0,0,160,200,'Alterar Parâmetros',,,,,CLR_WHITE,CLR_BLACK,,oDlg,.T.,,,,.T.)

    //GET PARA RECEPÇÃO DO PARÂMETRO
    oGet1 := TGet():New(10,05, {|u|"DIGITE O PARÂMETRO" + If(PCount()>0,cParGet1:=u,cParGet1)},oDlg,60,10,'',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","",,)
    oGet2 := TGet():New(40,05, {|u|"DIGITE O NOVO VALOR" + If(PCount()>0,cParGet2:=u,cParGet2)},oDlg,60,10,'',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","",,)
    
    //BOTÕES DE INTERAÇÃO
    oButt1 := TButton():New(13,65,"Procurar",oDlg,{||cParVlr1 := GetMV(cParGet1) },30,09,,,.F.,.T.,.F.,,.F.,,,.F.)
    oButt2 := TButton():New(43,65,"Alterar",oDlg, {||cParVlr2 := PutMV(cParGet1,cParGet2) },30,09,,,.F.,.T.,.F.,,.F.,,,.F.)
    
    //VISUALIZAÇÃO DO CONTEÚDO DO PARÂMETRO
    oTsay1 := TSay():New(25,05,{||"Valor atual: [ " + cValToChar(cParVlr1) + " ]"},oDlg,,,,,,.T.,CLR_BLACK,CLR_WHITE,200,20)
    oTsay2 := TSay():New(55,05,{||"Valor Alterado: [ " + cValToChar(GetMV(cParGet1)) + " ]"},oDlg,,,,,,.T.,CLR_BLACK,CLR_WHITE,200,20)
    oDlg:Activate()

Return
