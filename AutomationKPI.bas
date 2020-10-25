'Variáveis Publica-------------------------------------------'
Public locations(1 To 24) As Range
Public otifWorkbook As Workbook, databaseWorkbook As Workbook

Sub OTIF()

'*******************************************************
'                                                      '
' Created By: Fabio Juan Verdile                       '
' Version: 25.10.20v2.4                                '
' Note: Create a OTIF using database extrated at TMS   '
' OITF: On Time In Full (Key Performance Indicator)    '
'                                                      '      
'*******************************************************

'Variáveis------------------------------------------------ '
Dim month As String, months(0 To 12) As String
Dim nameColumn(1 To 45) As String
Dim otifPath As String
Dim otifPathFolder As String
Dim defaultOtifPath As String
Dim firstCheckRows As Long, lastCheckRows As Long
Dim arr(1 To 30) As String

'--Desativa tela e alertas
'Application.ScreenUpdating = False
'Application.DisplayAlerts = False

'--Atribui nome mês
month = Worksheets("VBA").Range("H2").Value

'--Chamando Funções
Call setColumns(nameColumn, months)

Call monthVerify(month, months)

Call setPath(month, otifPath, otifPathFolder, defaultOtifPath)

Call openOtif(month, months, otifPath, otifPathFolder, defaultOtifPath)

Call rowsVerify(month, firstCheckRows, lastCheckRows)

Call cleanColumns(month, nameColumn)

Call clients(month)

Call otifDatabase(month)

Call refreshDinamic

Call dateMonth

Call rowsVerify(month, firstCheckRows, lastCheckRows)

Call setLocations(month, months, locations)

Call inputValueClients(month)

Call inputValueStates(month)

Call message(firstCheckRows, lastCheckRows)

'--Zerando valores
Set dbWorkbook = Nothing
Set Workbook = Nothing
Set otifWorkbook = Nothing
Erase nameColumn
Erase locations
first = 0
last = 0
month = ""

'--Ativando tela e alertas
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Function monthVerify(month, months)

Dim count As Integer

'--Verifica se há informação
    If month = "" Then
        MsgBox "Favor insira um mês!"
        End
    End If

'--Verifica se o nome está correto
For count = 1 To 12
    If month = months(count) Then
        Exit Function
    End If
Next count

MsgBox "Favor verifique a ortografia"
End

End Function

Function setPath(month, otifPath, otifPathFolder, defaultOtifPath)

Dim automatePathFolder As String
Dim rootPathFolder As String
    
'--Caminho Pastas
automatePathFolder = Application.ActiveWorkbook.Path & "/"
Debug.Print automatePathFolder
rootPathFolder = Left(automatePathFolder, Len(automatePathFolder) - 16)
otifPathFolder = rootPathFolder & "Indicadores/OTIF/2020/"

'--Caminho Arquivos
Set databaseWorkbook = ActiveWorkbook
defaultOtifPath = automatePathFolder & "2020/BC.TRA-FO.052.01 - (OTIF FMCG).xlsm"
otifPath = rootPathFolder & "Indicadores/OTIF/2020/" & "OTIF - " & month & ".xlsm"

End Function

Function openOtif(month, months, otifPath, otifPathFolder, defaultOtifPath)

'--Abre OTIF
On Error GoTo stepError
'--Atribui caminho arquivo OTIF
Set otifWorkbook = Application.Workbooks.Open(otifPath)
 'Set otifWorkbook = Application.Workbooks.Open(otifPath)
 Exit Function
 
'--Caso não exista chama função
stepError: Call createOtif(month, months, otifPathFolder, defaultOtifPath)

End Function

Function createOtif(month, months, otifPathFolder, defaultOtifPath)

Dim lastMonth As String
Dim count As Long
Dim copyOtif As String
Dim newOtif As String

'--Cria OTIF
For count = 1 To 12
    If month = months(count) Then
        lastMonth = months(count - 1)
        copyOtif = otifPathFolder & "OTIF - " & lastMonth & ".xlsm"
        newOtif = otifPathFolder & "OTIF - " & month & ".xlsm"
        Exit For
    End If
Next

'--Criando Cópia OTIF
On Error GoTo defaultOtif
Application.Workbooks.Open (copyOtif)
ActiveWorkbook.SaveAs Filename:=newOtif, FileFormat:=xlOpenXMLWorkbookMacroEnabled
ActiveWorkbook.Close
'--Atrubuindo caminho novo OTIF criado
Set otifWorkbook = Application.Workbooks.Open(newOtif)
Exit Function

'--Caso não encontre nunhum arquivo cria um com o OTIF Modelo
defaultOtif:
copyOtif = defaultOtifPath
newOtif = otifPathFolder & "OTIF - " & month & ".xlsm"
Application.Workbooks.Open (copyOtif)
ActiveWorkbook.SaveAs Filename:=newOtif, FileFormat:=xlOpenXMLWorkbookMacroEnabled
ActiveWorkbook.Close
'--Atrubuindo caminho novo OTIF criado
Set otifWorkbook = Application.Workbooks.Open(newOtif)

End Function

Function rowsVerify(month, firstCheckRows, lastCheckRows)
    
Dim cRows As Long
    
'--Verifica a quantidade de linhas no inicio
If firstCheckRows = 0 Then
    databaseWorkbook.Activate
        cRows = 1
        While Worksheets(month).Cells(cRows, 3).Value <> ""
            cRows = cRows + 1
        Wend
    cRows = cRows - 1
    firstCheckRows = cRows
Else

'--Verifica novamente a quantidade de linhas no final do VBA
    cRows = 3
    otifWorkbook.Activate
        
'--On Error caso tenha #ND na base
    On Error Resume Next
        While Worksheets("Base Dados").Cells(cRows, 4).Value <> ""
            cRows = cRows + 1
        Wend
    On Error GoTo 0
    
    cRows = cRows - 2
    lastCheckRows = cRows

End If

End Function

Function cleanColumns(month, nameColumn)

Dim cColumns As Long
Dim saveColumn As String
Dim checkArr As Long

databaseWorkbook.Activate
Worksheets(month).Select

'--Remove os espaços no cabeçalho
Sheets(month).Rows(1).Replace What:=" ", Replacement:=""

'--Verifica as colunas setadas que irá ficar no relatório
For cColumns = 1 To 150
    For checkArr = 1 To 44
        If Worksheets(month).Cells(1, cColumns).Value = nameColumn(checkArr) Then
            saveColumn = nameColumn(checkArr)
        End If
    Next checkArr

'--Apaga informações das colunas que serão excluidas
    If Worksheets(month).Cells(1, cColumns).Value <> saveColumn Then
        Worksheets(month).Columns(cColumns).ClearContents
    End If
Next cColumns

'--Exclui as colunas em branco
For cColumns = 130 To 1 Step -1
    If Worksheets(month).Cells(1, cColumns).Value = "" Then
        Columns(cColumns).Delete
    End If
Next cColumns

Range("A1").Select

End Function

Function setColumns(nameColumn, months)

'--Atribui nome das colunas desejadas
nameColumn(1) = "NUMERO_CTRC"
nameColumn(2) = "NF"
nameColumn(3) = "VAL_NF"
nameColumn(4) = "DESCR_EMPRESA"
nameColumn(5) = "MODAL"
nameColumn(6) = "RAZ_CLI_PAGADOR"
nameColumn(7) = "CIDADE_ENTREGA"
nameColumn(8) = "ESTADO_ENTREGA"
nameColumn(9) = "REGIAO"
nameColumn(10) = "DATA_EMISSAO_CTRC"
nameColumn(11) = "DATA_EXPED_ULT_PLAN"
nameColumn(12) = "DATA_EMIS_MDFE_ULT_PLAN"
nameColumn(13) = "DATA_PREVISAO_ENTREGA"
nameColumn(14) = "DATA_AGENDADA"
nameColumn(15) = "ENTREGA_EFETIVADA"
nameColumn(16) = "DATA_NF"
nameColumn(17) = "NUM_BO"
nameColumn(18) = "DATA_INC_BO"
nameColumn(19) = "GRUPO_BO"
nameColumn(20) = "DESC_MOTIVO_NF"
nameColumn(21) = "DESC_CAUSA_NF"
nameColumn(22) = "STATUS_BO"
nameColumn(23) = "VAL_INDENIZ"
nameColumn(24) = "RESPONSABILIDADE"
nameColumn(25) = "DATA_CONF_BAIXA"
nameColumn(26) = "RAZ_TRANSP_RESP_BO"
nameColumn(27) = "PLACA_PLANILHA_PRIM"
nameColumn(28) = "NOME_MOT_PLANILHA_PRIM"
nameColumn(29) = "FINALIZADO_PLANILHA_PRIM"
nameColumn(30) = "TIPO_PLANILHA_PRIM"
nameColumn(31) = "TIPO_TRANSP_PRIM_PLAN"
nameColumn(32) = "RAZ_TRANSP_PRIM_PLAN"
nameColumn(33) = "PLACA_PLANILHA_ATU/ULT"
nameColumn(34) = "NOME_MOT_PLANILHA_ATU/ULT"
nameColumn(35) = "FINALIZADO_PLANILHA_ATU/ULT"
nameColumn(36) = "TIPO_PLANILHA_ATU/ULT"
nameColumn(37) = "TIPO_TRANSP_PLAN_ATU/ULT"
nameColumn(38) = "RAZ_TRANSP_ATUAL/ULTIMO"
nameColumn(39) = "DESCRICAO_CENTRO_CUSTO"
nameColumn(40) = "RAZ_TRANSP_MUNICIP"
nameColumn(41) = "RAZ_TRANSP_REDESP"
nameColumn(42) = "B.U."
nameColumn(43) = "EMPRESA_CTRC"

'--Atribui nome dos meses
months(0) = "Janeiro"
months(1) = "Janeiro"
months(2) = "Fevereiro"
months(3) = "Março"
months(4) = "Abril"
months(5) = "Maio"
months(6) = "Junho"
months(7) = "Julho"
months(8) = "Agosto"
months(9) = "Setembro"
months(10) = "Outubro"
months(11) = "Novembro"
months(12) = "Dezembro"

End Function

Function clients(month)
    
'--Insere coluna e realiza o PROCV buscando clientes
databaseWorkbook.Activate
Worksheets(month).Select
Worksheets(month).Columns("A:A").Insert
Worksheets(month).Range("A1").FormulaR1C1 = "CLIENTES"
Worksheets(month).Range("A2").FormulaR1C1 = "=VLOOKUP(RC[40],'De Para'!R3C2:R500C7,6,0)"
Worksheets(month).Range("A2").AutoFill Destination:=Range("A2:A110000")
Worksheets(month).Cells.Select

'--Modifica Layout
With Selection.Font
    .Name = "Calibri Light"
    .Size = 9
End With

Cells.EntireColumn.AutoFit

End Function

Function otifDatabase(month)

otifWorkbook.Activate
'--Limpando base de dados dentro da planilha OTIF
Worksheets("Base Dados").Range("D3:AU500000").ClearContents
    
'--Seleciona Base dados do Banco Dados
databaseWorkbook.Activate
Worksheets(month).Select
Range("AR1").Select
Range(Selection.End(xlToLeft), Selection.End(xlDown)).Copy

'--Passa para base do OTIF
otifWorkbook.Activate
Sheets("Base Dados").Select
Range("D3").Select
Sheets("Base Dados").Paste
Rows("3").Delete
Range("D3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
    
End Function

Function refreshDinamic()
    
'--Atuaiza Dinâmicas
otifWorkbook.Activate
Sheets("Dinamica").Select
Range("E15").Select
ActiveWorkbook.RefreshAll

End Function

Function dateMonth()

'--Passa o valor do mês do OTIF gerado
otifWorkbook.Activate
Worksheets("Base Dados").Range("T3").Copy
Worksheets("Base Calculo").Select
Range("N58").Select
Range("N58").PasteSpecial Paste:=xlPasteValues

End Function

Function setLocations(month, months, locations)

Dim count As Integer
Dim cColumnsFirst As Integer
Dim cColumnsSecond As Integer
Dim cColumnsThree As Integer
Dim cColumnsFour As Integer
Dim moveColumns As Integer

otifWorkbook.Activate
Sheets("Base Calculo").Select

For count = 1 To 12
    If month = months(count) Then
    
'--Inicia as colunas em Janeiro
    cColumnsFirst = 4 + count
    cColumnsSecond = 21 + count
    cColumnsThree = 39 + count
    cColumnsFour = 57 + count
        
        'Primeira coluna-----------------------------------------------------
        Set locations(1) = Worksheets("Base Calculo").Cells(8, cColumnsFirst) 'OTIF TR
        Set locations(2) = Worksheets("Base Calculo").Cells(9, cColumnsFirst)
        
        Set locations(3) = Worksheets("Base Calculo").Cells(14, cColumnsFirst) 'OTIF AGV
        Set locations(4) = Worksheets("Base Calculo").Cells(15, cColumnsFirst)
        
        Set locations(5) = Worksheets("Base Calculo").Cells(20, cColumnsFirst) 'OTIF GERAL
        Set locations(6) = Worksheets("Base Calculo").Cells(21, cColumnsFirst)
        
        'Segunda Coluna-------------------------------------------------------
        Set locations(7) = Worksheets("Base Calculo").Cells(8, cColumnsSecond)
        Set locations(8) = Worksheets("Base Calculo").Cells(9, cColumnsSecond)
        
        Set locations(9) = Worksheets("Base Calculo").Cells(14, cColumnsSecond)
        Set locations(10) = Worksheets("Base Calculo").Cells(15, cColumnsSecond)
        
        Set locations(11) = Worksheets("Base Calculo").Cells(20, cColumnsSecond)
        Set locations(12) = Worksheets("Base Calculo").Cells(21, cColumnsSecond)
        
        'Terçeira Coluna-------------------------------------------------------
        Set locations(13) = Worksheets("Base Calculo").Cells(8, cColumnsThree)
        Set locations(14) = Worksheets("Base Calculo").Cells(9, cColumnsThree)
        
        Set locations(15) = Worksheets("Base Calculo").Cells(14, cColumnsThree)
        Set locations(16) = Worksheets("Base Calculo").Cells(15, cColumnsThree)
        
        Set locations(17) = Worksheets("Base Calculo").Cells(20, cColumnsThree)
        Set locations(18) = Worksheets("Base Calculo").Cells(21, cColumnsThree)
        
        'Quarta Coluna----------------------------------------------
        Set locations(19) = Worksheets("Base Calculo").Cells(8, cColumnsFour)
        Set locations(20) = Worksheets("Base Calculo").Cells(9, cColumnsFour)
        
        Set locations(21) = Worksheets("Base Calculo").Cells(14, cColumnsFour)
        Set locations(22) = Worksheets("Base Calculo").Cells(15, cColumnsFour)
        
        Set locations(23) = Worksheets("Base Calculo").Cells(20, cColumnsFour)
        Set locations(24) = Worksheets("Base Calculo").Cells(21, cColumnsFour)
        Exit For
    End If
Next count

'--Atribui os valores nos locais indicados

'-- OTIF CONSOLIDADO
'--OTIF TR
locations(1) = Worksheets("Base Calculo").Range("E47").Value 'nf
locations(2) = Worksheets("Base Calculo").Range("E48").Value 'bo
'--OTIF AGV
locations(3) = Worksheets("Base Calculo").Range("E47").Value 'nf
locations(4) = Worksheets("Base Calculo").Range("E49").Value 'bo
'--OTIF GERAL
locations(5) = Worksheets("Base Calculo").Range("E47").Value 'nf
locations(6) = Worksheets("Base Calculo").Range("E50").Value 'bo

'--OTIF FRACIONADO
'--OTIF TR
locations(7) = Worksheets("Base Calculo").Range("V35").Value 'nf
locations(8) = Worksheets("Base Calculo").Range("V36").Value 'bo
'--OTIF AGV
locations(9) = Worksheets("Base Calculo").Range("V35").Value 'nf
locations(10) = Worksheets("Base Calculo").Range("V37").Value 'bo
'--OTIF GERAL
locations(11) = Worksheets("Base Calculo").Range("V35").Value 'nf
locations(12) = Worksheets("Base Calculo").Range("V38").Value 'bo

'--OTIF FECHADO
'--OTIF TR
locations(13) = Worksheets("Base Calculo").Range("AN35").Value 'nf
locations(14) = Worksheets("Base Calculo").Range("AN36").Value 'bo
'--OTIF AGV
locations(15) = Worksheets("Base Calculo").Range("AN35").Value 'nf
locations(16) = Worksheets("Base Calculo").Range("AN37").Value 'bo
'--OTIF GERAL
locations(17) = Worksheets("Base Calculo").Range("AN35").Value 'nf
locations(18) = Worksheets("Base Calculo").Range("AN38").Value 'bo

'--OTIF AÉREO
'--OTIF TR
locations(19) = Worksheets("Base Calculo").Range("BF35").Value 'nf
locations(20) = Worksheets("Base Calculo").Range("BF36").Value 'bo
'--OTIF AGV
locations(21) = Worksheets("Base Calculo").Range("BF35").Value 'nf
locations(22) = Worksheets("Base Calculo").Range("BF37").Value 'bo
'--OTIF GERAL
locations(23) = Worksheets("Base Calculo").Range("BF35").Value 'nf
locations(24) = Worksheets("Base Calculo").Range("BF38").Value 'bo

End Function

Function inputValueClients(month)
    
Dim cColumns As Long
Dim cRows As Long

'Inputando valores na da aba cliente no acompanhamento anual
otifWorkbook.Activate
Sheets("cliente").Select

'-- Conta coluna
For cColumns = 6 To 17
    '-- conta linha
    For cRows = 75 To 100
        If Cells(cRows, cColumns).Value = month Then
        cRows = cRows + 1
            'verifica qual é o mês para inserir no local certo
           If month = "Janeiro" Then
                'Na celula especificada fazer o procv com seerro
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],R5C5:R74C10,6,0),""-"")"
                    'Porcentagem
                    Selection.Style = "Percent"
                    'Valores nesse formato
                    Selection.NumberFormat = "0.00%"
                    'Pega todo o range
                    Cells(cRows, cColumns).AutoFill Destination:=Range("F79:F94")
                        Range("F79:F94").Copy
                        Range("F79:F94").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                    
            '-- Repete para cada mês
           ElseIf month = "Fevereiro" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("G79:G94")
                        Range("G79:G94").Copy
                        Range("G79:G94").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For

            ElseIf month = "Março" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColums).AutoFill Destination:=Range("H79:H94")
                        Range("H79:H94").Copy
                        Range("H79:H94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
                    
           ElseIf month = "Abril" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("I79:I94")
                        Range("I79:I94").Copy
                        Range("I79:I94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
                
           ElseIf month = "Maio" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("J79:J94")
                        Range("J79:J94").Copy
                        Range("J79:J94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
                    
           ElseIf month = "Junho" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("K79:K94")
                        Range("K79:K94").Copy
                        Range("K79:K94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
               
           ElseIf month = "Julho" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("L79:L94")
                        Range("L79:L94").Copy
                        Range("L79:L94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
               
           ElseIf month = "Agosto" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("M79:M94")
                        Range("M79:M94").Copy
                        Range("M79:M94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
                
           ElseIf month = "Setembro" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("N79:N94")
                        Range("N79:N94").Copy
                        Range("N79:N94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
               
           ElseIf month = "Outubro" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("O79:O94")
                        Range("O79:O94").Copy
                        Range("O79:O94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
               
           ElseIf month = "Novembro" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("P79:P94")
                        Range("P79:P94").Copy
                        Range("P79:P94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
               
           ElseIf month = "Dezembro" Then
               Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],R5C5:R74C10,6,0),""-"")"
                    Selection.Style = "Percent"
                    Selection.NumberFormat = "0.00%"
                    Cells(cRows, cColumns).AutoFill Destination:=Range("Q79:Q94")
                        Range("Q79:Q94").Copy
                        Range("Q79:Q94").PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                Exit For
            End If
        End If
    Next cRows
Next cColumns

End Function

Function inputValueStates(month)

    'Inputando valores no consolidado anual dos valores das regiões
    'Ativando Planilha OTIF
    otifWorkbook.Activate
    
    'Ativando a aba região
    Sheets("Região").Select
    
    'Conta coluna
    For cColumns = 5 To 17
        'conta linha
        For cRows = 35 To 63
            'Assim que achar a celula com o nome do mês
            If Cells(cRows, cColumns).Value = month Then
            cRows = cRows + 1
                'verifica qual é o mês para inserir no local certo
               If month = "Janeiro" Then
                    'Faz o PROCV com SEERRO
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],R5C4:R31C8,3,0),""-"")"
                    'Desce o range
                        Cells(cRows, cColumns).AutoFill Destination:=Range("E36:E62")
                        'Aqui soma o valor de celula até chegar no inicio do preenchimento de bo que fica abaixo
                        cRows = cRows + 33
                    'Procv com seerro
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],R5C4:R31C8,4,0),""-"")"
                        'Desce o range
                        Cells(cRows, cColumns).AutoFill Destination:=Range("E69:E95")
                            'Copia o range
                            Range("Q79:Q94").Copy
                            'Cola como valor
                            Range("Q79:Q94").PasteSpecial Paste:=xlPasteValues
                        'Tira a seleção do range
                        Application.CutCopyMode = False
                    'Para encerrar o FOR
                    Exit For
                        
                'Para cada mês
               ElseIf month = "Fevereiro" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("F36:F62")
                            Range("F36:F62").Copy
                            Range("F36:F62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("F69:F95")
                            Range("F69:F95").Copy
                            Range("F69:F95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For

                ElseIf month = "Março" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("G36:G62")
                            Range("G36:G62").Copy
                            Range("G36:G62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("G69:G95")
                            Range("G69:G95").Copy
                            Range("G69:G95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                        
               ElseIf month = "Abril" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("H36:H62")
                            Range("H36:H62").Copy
                            Range("H36:H62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("H69:H95")
                            Range("H69:H95").Copy
                            Range("H69:H95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                    
               ElseIf month = "Maio" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("I36:I62")
                            Range("I36:I62").Copy
                            Range("I36:I62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("I69:I95")
                            Range("I69:I95").Copy
                            Range("I69:I95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                        
               ElseIf month = "Junho" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("J36:J62")
                            Range("J36:J62").Copy
                            Range("J36:J62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("J69:J95")
                            Range("J69:J95").Copy
                            Range("J69:J95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                   
               ElseIf month = "Julho" Then
                     Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("K36:K62")
                            Range("K36:K62").Copy
                            Range("K36:K62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("K69:K95")
                            Range("K69:K95").Copy
                            Range("K69:K95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                   
               ElseIf month = "Agosto" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("L36:L62")
                            Range("L36:L62").Copy
                            Range("L36:L62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("L69:L95")
                            Range("L69:L95").Copy
                            Range("L69:L95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                    
               ElseIf month = "Setembro" Then
                     Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("M36:M62")
                            Range("M36:M62").Copy
                            Range("M36:M62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("M69:M95")
                            Range("M69:M95").Copy
                            Range("M69:M95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                   
               ElseIf month = "Outubro" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("N36:N62")
                            Range("N36:N62").Copy
                            Range("N36:N62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("N69:N95")
                            Range("N36:N62").Copy
                            Range("N36:N62").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                   
               ElseIf month = "Novembro" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("O36:O62")
                            Range("O36:O62").Copy
                            Range("O36:O62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("O69:O95")
                            Range("O69:O95").Copy
                            Range("O69:O95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                   
               ElseIf month = "Dezembro" Then
                   Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],R5C4:R31C8,3,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("P36:P62")
                            Range("P36:P62").Copy
                            Range("P36:P62").PasteSpecial Paste:=xlPasteValues
                            Application.CutCopyMode = False
                        cRows = cRows + 33
                    Cells(cRows, cColumns).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],R5C4:R31C8,4,0),""-"")"
                        Cells(cRows, cColumns).AutoFill Destination:=Range("P69:P95")
                            Range("P69:P95").Copy
                            Range("P69:P95").PasteSpecial Paste:=xlPasteValues
                        Application.CutCopyMode = False
                    Exit For
                End If
            End If
        Next cRows
    Next cColumns

 End Function

Function message(firstCheckRows, lastCheckRows)

    '--menssagem
    MsgBox "Processo concluido!" & Chr(13) & Chr(13) & _
    "Quantidade de linhas iniciais:" & " " & firstCheckRows & Chr(13) & _
    "Quantidade de linhas finais:" & " " & lastCheckRows & Chr(13) & Chr(13)

End Function
