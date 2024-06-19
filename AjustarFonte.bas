Attribute VB_Name = "AjustarFonte"
Sub AjustarFonte(tabela As Variant)

    With ThisWorkbook.Sheets("Por Solicitante").Range(tabela)
        .Font.Name = "Times New Roman"
        .Font.Size = 14
    End With
    
End Sub

