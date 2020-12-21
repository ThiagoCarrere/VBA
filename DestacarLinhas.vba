
Dim Linha As Long 'Variavel global
Dim restricao as integer 'Variavel que vai indicar a partir de qual linha que a funcao terá inicio 

restricao = 2 'nesse caso comecará a partir da linha 2 ou seja, na linha 3

'Funcao que colore a linha da celula selecionada
Private Sub Workbook_Activate()
    On Error Resume Next
    Linha = ActiveCell.Row 'Variavel global recebe a o numero da linha selecionada
    If Linha > restricao Then
        Range(Cells(Linha, 1), Cells(Linha, 6)).Interior.ColorIndex = 8 '= RGB(158, 111, 213) cor de teste
        'Range(Cells(Linha, 1), Cells(Linha, 6)).Font.Bold = True
    End If
End Sub

'Função para limpar a linha antes de colorir a proxima selecao
Private Sub Workbook_Deactivate()
    On Error Resume Next
If Linha > restricao Then
        Range(Cells(Linha, 1), Cells(Linha, 6)).Interior.ColorIndex = xlNone    'limpa a cor anterior
        'Range(Cells(Linha, 1), Cells(Linha, 6)).Font.Bold = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
   Call Workbook_Deactivate
   Call Workbook_Activate
End Sub
