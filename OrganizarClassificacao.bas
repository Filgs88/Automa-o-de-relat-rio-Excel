Attribute VB_Name = "OrganizarClassificacao"
Sub OrganizarClassificacao()

'Classificar Tabela 11
    
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Add(Range("Tabela11[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Add(Range("Tabela11[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Add(Range("Tabela11[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Add(Range("Tabela11[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Add(Range("Tabela11[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Classificar Tabela 12

    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Add(Range("Tabela12[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Add(Range("Tabela12[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Add(Range("Tabela12[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Add(Range("Tabela12[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Add(Range("Tabela12[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Classificar Tabela 13
    
        ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Add(Range("Tabela13[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Add(Range("Tabela13[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Add(Range("Tabela13[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Add(Range("Tabela13[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Add(Range("Tabela13[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
     
    'Classificar Tabela 14
    
        ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Add(Range("Tabela14[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Add(Range("Tabela14[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Add(Range("Tabela14[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Add(Range("Tabela14[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Add(Range("Tabela14[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        
    'Classificar Tabela 15
    
        ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Add(Range("Tabela15[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Add(Range("Tabela15[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Add(Range("Tabela15[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Add(Range("Tabela15[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Add(Range("Tabela15[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Classificar Tabela 16
    
        ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Add(Range("Tabela16[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Add(Range("Tabela16[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Add(Range("Tabela16[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Add(Range("Tabela16[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Add(Range("Tabela16[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Classificar Tabela 17
    
        ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Add(Range("Tabela17[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Add(Range("Tabela17[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Add(Range("Tabela17[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Add(Range("Tabela17[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Add(Range("Tabela17[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Classificar Tabela 18
    
        ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Add(Range("Tabela18[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Add(Range("Tabela18[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Add(Range("Tabela18[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Add(Range("Tabela18[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Add(Range("Tabela18[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Classificar Tabela 19
    
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Add(Range("Tabela19[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Add(Range("Tabela19[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Add(Range("Tabela19[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Add(Range("Tabela19[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Add(Range("Tabela19[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Classificar Tabela20
    
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Add(Range("Tabela20[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Add(Range("Tabela20[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Add(Range("Tabela20[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Add(Range("Tabela20[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Add(Range("Tabela20[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Classificar Tabela 21
    
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Add(Range("Tabela21[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Add(Range("Tabela21[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Add(Range("Tabela21[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Add(Range("Tabela21[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Add(Range("Tabela21[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Classificar Tabela 22
    
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Add(Range("Tabela22[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Add(Range("Tabela22[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Add(Range("Tabela22[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Add(Range("Tabela22[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Add(Range("Tabela22[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela 23
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Add(Range("Tabela23[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Add(Range("Tabela23[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Add(Range("Tabela23[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Add(Range("Tabela23[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Add(Range("Tabela23[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela 24
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Add(Range("Tabela24[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 113, 113)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Add(Range("Tabela24[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(147, 227, 255)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Add(Range("Tabela24[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(240, 169, 74)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Add(Range("Tabela24[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 255, 147)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Add(Range("Tabela24[[#All],[Situação]]"), xlSortOnCellColor, _
        xlAscending, , xlSortNormal).SortOnValue.Color = RGB(217, 217, 217)
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
