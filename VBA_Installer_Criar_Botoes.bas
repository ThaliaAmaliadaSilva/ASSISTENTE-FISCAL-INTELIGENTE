Attribute VB_Name = "modInstaller"
Option Explicit

' Cria botões na aba "NotasFiscais" e salva o arquivo como .xlsm com macros embutidas.
' Execute este módulo após importar os módulos principais (Verificar... e Relatorio...).
Public Sub Criar_Botoes_e_SalvarXLSM()
    Dim ws As Worksheet
    Dim btn1 As Shape, btn2 As Shape
    Dim topPos As Double, leftPos As Double

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("NotasFiscais")

    ' Remove botões existentes com os nomes definidos, se houver
    On Error Resume Next
    ws.Shapes("btnVerificar").Delete
    ws.Shapes("btnRelatorio").Delete
    On Error GoTo ErrHandler

    leftPos = 10
    topPos = 10

    ' Adiciona botão Verificar_CFOP_CST
    Set btn1 = ws.Shapes.AddFormControl(Type:=xlButtonControl, Left:=leftPos, Top:=topPos, Width:=160, Height:=30)
    btn1.Name = "btnVerificar"
    btn1.TextFrame.Characters.Text = "Verificar CFOP / CST"
    btn1.OnAction = "Verificar_CFOP_CST"

    ' Adiciona botão Relatorio_Divergencias_Aliquota_Com_Regras
    Set btn2 = ws.Shapes.AddFormControl(Type:=xlButtonControl, Left:=leftPos + 170, Top:=topPos, Width:=220, Height:=30)
    btn2.Name = "btnRelatorio"
    btn2.TextFrame.Characters.Text = "Gerar Relatório de Divergências de Alíquota"
    btn2.OnAction = "Relatorio_Divergencias_Aliquota_Com_Regras"

    ' Salva o workbook com macros embutidas (.xlsm) no mesmo diretório
    Dim savePath As String
    savePath = ThisWorkbook.Path & Application.PathSeparator & "Assistente_Fiscal_Protótipo_Regras_Macros.xlsm"
    ThisWorkbook.SaveAs Filename:=savePath, FileFormat:=52

    MsgBox "Botões criados e arquivo salvo como:" & vbCrLf & savePath, vbInformation, "Instalação Concluída"
    Exit Sub

ErrHandler:
    MsgBox "Erro ao criar botões: " & Err.Description, vbExclamation, "Erro"
End Sub