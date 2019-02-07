Sub lsLigarTelaCheia()

    'Oculta todas as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de fórmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = False
    
    'Alterar o nome do Excel
    Application.Caption = " <> by Nome do Criador"
    
    With ActiveWindow
        'Ocultar barra horizontal
        .DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        .DisplayVerticalScrollBar = False
        
        'Ocultar guias das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os títulos de linha e coluna
        .DisplayHeadings = False
        
        'Oculta valores zero na planilha
        .DisplayZeros = False
        
        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With

End Sub

Sub lsDesligarTelaCheia()

'Reexibe os menus
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    
    'Reexibir a barra de fórmulas
    Application.DisplayFormulaBar = True
    
    'Reexibir a barra de status, disposta ao final da planilha
    Application.DisplayStatusBar = True
    
    'Reexibir o cabeçalho da Pasta de trabalho
    ActiveWindow.DisplayHeadings = True
    
    'Retornar o nome do Excel
    Application.Caption = ""
    
    With ActiveWindow
        'Reexibir barra horizontal
        .DisplayHorizontalScrollBar = True
        
        'Reexibir barra vertical
        .DisplayVerticalScrollBar = True
        
        'Reexibir guias das planilhas
        .DisplayWorkbookTabs = True
        
        'Reexibir os títulos de linha e coluna
        .DisplayHeadings = True
        
        'Reexibir valores zero na planilha
        .DisplayZeros = True
        
        'Reexibir as linhas de grade da planilha
        .DisplayGridlines = True
    End With

End Sub

'para chamar a função coloque o seguinte codigo na abertura da planilha 

Private Sub Workbook_Open()     
     
     lsLigarTelaCheia
     
End Sub
