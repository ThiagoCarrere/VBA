Dim Linha As Long 'Variavel global
'Função que colore a linha da celula selecionada
Private Sub Workbook_Activate()
    On Error Resume Next
    Linha = ActiveCell.Row 'Variavel global recebe a o numero da linha selecionada
    If Linha >= 2 Then
        Range(Cells(Linha, 1), Cells(Linha, 6)).Interior.ColorIndex = 8 '= RGB(158, 111, 213) Destaca linha
        Range(Cells(Linha, 1), Cells(Linha, 6)).Font.Bold = True
    End If
End Sub

'Função para limpar a linha antes de colorir a proxima selecao
Private Sub Workbook_Deactivate()
    On Error Resume Next
    If Linha >= 2 Then
        Range(Cells(Linha, 1), Cells(Linha, 6)).Interior.ColorIndex = xlNone    'limpa a cor anterior
        Range(Cells(Linha, 1), Cells(Linha, 6)).Font.Bold = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
   Call Workbook_Deactivate
   Call Workbook_Activate
End Sub