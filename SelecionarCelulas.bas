Attribute VB_Name = "SelecionarCelulas"
Sub SelecCelulas()
    Dim cel As Integer
    Dim linha As Integer
    Dim lin As Integer
    
    cel = 5
    linha = 6
    lin = 6
    
    Range("O6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Planilha1.Cells(linha, 1).Select
    
    Do While Planilha1.Cells(linha, 1) > 0
    
    Planilha1.Cells(linha, 1).Select
    
        If InStr(Planilha1.Cells(linha, 9), "Entressafra 2023/24") > 0 Then
            GoTo Proximo
        Else
            Selection.Copy
            Planilha1.Cells(cel, 15).Select
                If ActiveSheet.Select = 0 Then
                    GoTo Colar
                Else
                    ActiveCell.Offset(1, 0).Select
                End If
        End If

Colar:
        ActiveSheet.Paste
        cel = cel + 1
        
Proximo:
        linha = linha + 1
    Loop
    
    Planilha1.Cells(lin, 15).Select
    
    Do While Planilha1.Cells((lin + 1), 15) > 0
    
        Selection.Resize(Selection.Rows.Count + 1).Select
        
        lin = lin + 1
    
    Loop
    
End Sub

Sub SelecCelulasEntressafra()
    Dim cel As Integer
    Dim linha As Integer
    Dim lin As Integer
    
    cel = 5
    linha = 6
    lin = 6
    
    Range("O6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Planilha1.Cells(linha, 1).Select
    
    Do While Planilha1.Cells(linha, 1) > 0
    
    Planilha1.Cells(linha, 1).Select
    
        If InStr(Planilha1.Cells(linha, 9), "Entressafra 2023/24") > 0 Then
            Selection.Copy
            Planilha1.Cells(cel, 15).Select
                If ActiveSheet.Select = 0 Then
                    GoTo Colar
                Else
                    ActiveCell.Offset(1, 0).Select
                End If
        Else
            GoTo Proximo
        End If

Colar:
        ActiveSheet.Paste
        cel = cel + 1
        
Proximo:
        linha = linha + 1
    Loop
    
    Planilha1.Cells(lin, 15).Select
    
    Do While Planilha1.Cells((lin + 1), 15) > 0
    
        Selection.Resize(Selection.Rows.Count + 1).Select
        
        lin = lin + 1
    
    Loop
    
End Sub
