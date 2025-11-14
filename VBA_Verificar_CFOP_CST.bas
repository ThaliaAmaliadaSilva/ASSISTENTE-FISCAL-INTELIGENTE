Attribute VB_Name = "modVerificarCfopCst"
Option Explicit

Public Sub Verificar_CFOP_CST()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim cfop As String, cst As String
    Dim erros As Long

    Set ws = ThisWorkbook.Sheets("NotasFiscais")

    Const COL_CFOP As Long = 2    ' B
    Const COL_CST  As Long = 3    ' C

    ultimaLinha = ws.Cells(ws.Rows.Count, COL_CFOP).End(xlUp).Row

    ' Limpa cores
    ws.Range(ws.Cells(2, 1), ws.Cells(ultimaLinha, 10)).Interior.ColorIndex = xlNone

    erros = 0

    For i = 2 To ultimaLinha
        cfop = Trim(CStr(ws.Cells(i, COL_CFOP).Value))
        cst = Trim(CStr(ws.Cells(i, COL_CST).Value))

        If (cfop Like "5*" And cst <> "000" And cst <> "060") Or _
           (cfop Like "6*" And cst <> "000" And cst <> "010") Or _
           (cfop Like "1*" And cst <> "070" And cst <> "000" And cst <> "020" And cst <> "060") Or _
           (cfop Like "2*" And cst <> "030" And cst <> "000" And cst <> "060") Then

            ws.Rows(i).Interior.Color = RGB(255, 153, 153)
            erros = erros + 1
        End If
    Next i

    If erros > 0 Then
        MsgBox "Foram encontradas " & erros & " linhas com divergências de CFOP x CST.", vbExclamation, "Verificação Concluída"
    Else
        MsgBox "Nenhuma divergência de CFOP x CST encontrada.", vbInformation, "Verificação Concluída"
    End If
End Sub