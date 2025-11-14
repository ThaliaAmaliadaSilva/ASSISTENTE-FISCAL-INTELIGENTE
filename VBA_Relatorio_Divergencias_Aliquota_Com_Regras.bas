Attribute VB_Name = "modRelatorioAliquotasComRegras"
Option Explicit

' Gera relatório de divergências de alíquota usando a aba "Regras_Aliquotas"
Public Sub Relatorio_Divergencias_Aliquota_Com_Regras()
    Dim wb As Workbook
    Dim wsOrigem As Worksheet
    Dim wsRel As Worksheet
    Dim wsRegras As Worksheet
    Dim ultimaLinha As Long, ultimaRegra As Long
    Dim i As Long, j As Long, linhaRel As Long
    Dim cfop As String, cst As String
    Dim aliqInformada As Double, aliqEsperada As Double
    Dim regraCfop As String, regraAliq As Double
    Dim regraEncontrada As String
    Dim erros As Long

    Set wb = ThisWorkbook
    Set wsOrigem = wb.Sheets("NotasFiscais")

    ' Verifica aba de regras
    On Error Resume Next
    Set wsRegras = wb.Sheets("Regras_Aliquotas")
    On Error GoTo 0
    If wsRegras Is Nothing Then
        MsgBox "Aba 'Regras_Aliquotas' não encontrada. Crie a aba com colunas: CFOP_Padrao | AliquotaEsperada", vbExclamation
        Exit Sub
    End If

    ' Apaga relatório antigo
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Relatorio_Erros").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Cria novo relatório
    Set wsRel = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsRel.Name = "Relatorio_Erros"
    wsRel.Range("A1:G1").Value = Array("Linha", "Nº NF", "CFOP", "CST", "Alíquota Informada", "Alíquota Esperada", "Regra Aplicada")
    wsRel.Rows(1).Font.Bold = True

    ' Colunas origem
    Const COL_NNF As Long = 1   ' A
    Const COL_CFOP As Long = 2  ' B
    Const COL_CST  As Long = 3  ' C
    Const COL_ALIQ As Long = 4  ' D

    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, COL_CFOP).End(xlUp).Row
    ultimaRegra = wsRegras.Cells(wsRegras.Rows.Count, 1).End(xlUp).Row

    linhaRel = 2
    erros = 0

    ' Loop nas linhas de notas
    For i = 2 To ultimaLinha
        cfop = Trim(CStr(wsOrigem.Cells(i, COL_CFOP).Value))
        cst = Trim(CStr(wsOrigem.Cells(i, COL_CST).Value))
        aliqInformada = Val(wsOrigem.Cells(i, COL_ALIQ).Value)

        aliqEsperada = -1
        regraEncontrada = "Nenhuma regra"

        ' Primeiro correspondência exata
        For j = 2 To ultimaRegra
            regraCfop = Trim(CStr(wsRegras.Cells(j, 1).Value))
            If regraCfop = "" Then GoTo ProxExact
            If InStr(regraCfop, "*") = 0 Then
                If cfop = regraCfop Then
                    regraAliq = Val(wsRegras.Cells(j, 2).Value)
                    aliqEsperada = regraAliq
                    regraEncontrada = "Exata: " & regraCfop
                    Exit For
                End If
            End If
ProxExact:
        Next j

        ' Depois padrões com curinga
        If aliqEsperada = -1 Then
            For j = 2 To ultimaRegra
                regraCfop = Trim(CStr(wsRegras.Cells(j, 1).Value))
                If regraCfop = "" Then GoTo ProxLike
                If InStr(regraCfop, "*") > 0 Then
                    If cfop Like regraCfop Then
                        regraAliq = Val(wsRegras.Cells(j, 2).Value)
                        aliqEsperada = regraAliq
                        regraEncontrada = "Padrão: " & regraCfop
                        Exit For
                    End If
                End If
ProxLike:
            Next j
        End If

        ' Se não achou, assume informada
        If aliqEsperada = -1 Then
            aliqEsperada = aliqInformada
            regraEncontrada = "Sem regra definida"
        End If

        ' Comparar
        If aliqInformada <> aliqEsperada Then
            wsRel.Cells(linhaRel, 1).Value = i
            wsRel.Cells(linhaRel, 2).Value = wsOrigem.Cells(i, COL_NNF).Value
            wsRel.Cells(linhaRel, 3).Value = cfop
            wsRel.Cells(linhaRel, 4).Value = cst
            wsRel.Cells(linhaRel, 5).Value = aliqInformada
            wsRel.Cells(linhaRel, 6).Value = aliqEsperada
            wsRel.Cells(linhaRel, 7).Value = regraEncontrada
            linhaRel = linhaRel + 1
            erros = erros + 1
        End If
    Next i

    wsRel.Columns("A:G").AutoFit

    If erros > 0 Then
        MsgBox "Foram encontradas " & erros & " divergências de alíquota. Veja 'Relatorio_Erros'.", vbExclamation, "Relatório Gerado"
    Else
        MsgBox "Nenhuma divergência de alíquota encontrada.", vbInformation, "Relatório Gerado"
    End If
End Sub