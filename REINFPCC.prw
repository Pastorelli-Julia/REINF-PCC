#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH
#INCLUDE "RWMAKE.CH"

//--------------------------------------------------------------------------------------------------------------//
//------------------|      Data       |---|      Função      |---|        Autor        |------------------------//
//------------------|     26/12/23    |---|      LRREINF     |---|   Julia Pastorelli  |------------------------//
//--------------------------------------------------------------------------------------------------------------//
// Descrição: Relatorio conferencia REINF - PCC                                                                                                         //
//--------------------------------------------------------------------------------------------------------------//

User Function REINFPCC() 

    Local cDesc1        := "Este programa tem como objetivo"
    Local cDesc2        := "imprimir relatorio dos registros do PCC"
    Local titulo        := "REINF PCC"
    Local cPerg         := "REINFPCC"
    Local aOrd          := {}
    Local aArea         := GetArea()
    Private tamanho     := "G"
    Private nomeprog    := "REINF PCC"
    Private nTipo       := 15
    Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1 }
    Private nLastKey    := 0 
    Private wnrel       := "REINF PCC"
    Private aRotina     := {}
    Private oApoio      := LibApoio():New()

    Pergunte(cPerg,.f.)

    wnrel := SetPrint("",NomeProg,cPerg,@titulo,cDesc1,cDesc2,,.F.,aOrd,.F.,Tamanho,,.F.)

    IF nLastKey == 27
        RETURN
    ENDIF

    SetDefault(aReturn,"")

    IF nLastKey(aReturn)
        RETURN
    ENDIF

    nTipo := IF(aReturn[4]==1,15,18)

    fGeraExcel()
    RestArea(aArea)

RETURN

//+--------------------------------------------------------------------------------------+
//| Gera EXCELL
//+--------------------------------------------------------------------------------------+

Static Function fGeraExcel()

    Local cQuery            := ""
    Local cPath             := ""
    Local cNameFile         := ""
    Local aRec              := ""
    Local aCab              := {}
    Local x                 := 0
    Private lMsErroAuto     := .F.
    Private aExcel          := FWMSEXCEL():New()

    cQuery := "SELECT" +CRLF
    cQuery += "F1_ESPECIE,F1_DOC,E2_PARCELA,F1_SERIE,F1_FORNECE,A2_NOME,A2_TIPO"

    cQuery += " FROM " + RetSqlName("SF1") + "  (NOBLOCK)" +CRLF
    cQuery += " INNER JOIN " + RetSqlName("SA2") + "  (NOBLOCK)" +CRLF
    cQuery += " A2_FILIAL = F1_FILIAL AND A2_COD = F1_FORNECE AND A2_LOJA = F1_LOJA AND" + RetSqlName("SA2") +  ".D_E_L_E_T_ = '' " +CRLF

    cQuery += " WHERE " +CRLF
    cQuery += " F1_ESPECIE = 'NFS', 'CTEOS', 'CTE'" +CRLF
    cQuery += " AND F1_EMISSAO BETWEEN '" +DTOS(mv_par01)+ "' AND '" +DTOS(mv_par02)+ "'" + CRLF
    cQuery += " AND " + RetSqlName("SF1") + ".D_E_L_E_T_=''"

    aRec := U_QryArr(cQuery)
        aadd( aCab, {""                 ,1 ,1 ,0})
        aadd( aCab, {""                 ,1 ,1 ,0})
        aadd( aCab, {""                 ,1 ,1 ,0})
        aadd( aCab, {""                 ,1 ,1 ,0})
        aadd( aCab, {""                 ,1 ,1 ,0})
        aadd( aCab, {""                 ,1 ,1 ,0})
        aadd( aCab, {""                 ,1 ,1 ,0})
        
        oExcel:AddworkSheet("REINFPCC")
        oExcel:AddTable ("REINFPCC", "REINFPCC")
        
        for x := 1 to  len(aCab)
            oExcel:AddColumn("REINFPCC", "REINFPCC", aCab[x][1],aCab[x][2],aCab[x][3])
        next x

        for x := 1 to len(aRec)
            oExcel:addRow("REINFPCC", "REINFPCC", {aRec[x][1],aRec[x][2],aRec[x][3],aRec[x][4],aRec[x][5],aRec[x][6],aRec[x][7]})
        next x

    oExcel:Activate()
    cPath := AllTrim(GetTempPath())
    cNameFile   := cPath + "REINFPCC.xls"

    oExcel:GetXMLFile(cNameFile)
    ShellExecute( "Open", cNameFile, '', '', 1 )

Return()
