#include 'protheus.ch'
#include 'totvs.ch'

/*========================================================================
|Programa:| Programa escrito com a finalidade de procurar por um parâmetro
e poder alterar o valor do mesmo sem precisar logar no CFG.              
--------------------------------------------------------------------------
|Data:    | 08/11/2020
--------------------------------------------------------------------------
|Autor:   | Matheus Santos de Oliveira / (61) 99812-9318
==========================================================================*/

User Function ALTMV()

    Local oDlg
    Local oGet, oGet2
    Local oBtt, oBtt2
    Local oSay, oSay2, oSay3 
    Local cGetVlr := space(10)
    Local cValPar
    Local cGetVlr2 := space(10)
    
    //Criação da tela
    oDlg := MSDialog():New(000,000,200,200,"Alterar Parâmetros",,,,,CLR_BLACK,CLR_WHITE,,oDlg,.T.,,,,.T.)
    
    //Criação das funções que carregam um valor e visualiza o valor atual
    oSay := TSay():New(11,05,{||"Pesquise o parâmetro"},oDlg,,oFont,,,,.T.,CLR_BLACK,CLR_WHITE,200,20)
    oGet := TGet():New(20,05,{|u| if(Pcount()>0,cGetVlr := u, cGetVlr)},oDlg,,,"@!",,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","cGetVlr",,)
    oBtt := TButton():New(20,70,"Procurar",oDlg,{|| cValPar := GetMV(cGetVlr)},30,10,,,,.T.,,,,,,,)
    oSay2 := TSay():New(33,06,{||"Valor atual: [ " + cValToChar(cValPar) + " ]"},oDlg,,oFont,,,,.T.,CLR_BLACK,CLR_WHITE,200,20)

    //Criação das funções que recebem um valor e alteram o valor do parâmetro mostrando seu resultado 
    oSay3 := TSay():New(55,05,{||"Digite o novo valor"},oDlg,,oFont,,,,.T.,CLR_BLACK,CLR_WHITE,200,20)
    oGet2 := TGet():New(65,05,{|u| if( Pcount()>0,cGetVlr2 := u, cGetVlr2)},oDlg,,,"@!",,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","cGetVlr2",,)
    oBtt2 := TButton():New(65,70,"Alterar",oDlg,{|| PutMV(cGetVlr,cGetVlr2)},30,10,,,,.T.,,,,,,,)
    oDlg:CtrlRefresh()
    
    oDlg:Activate(,,,.T.)

Return
