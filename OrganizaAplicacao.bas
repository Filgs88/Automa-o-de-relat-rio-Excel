Attribute VB_Name = "OrganizaAplicacao"
Sub ClassificaAplicacao()
Attribute ClassificaAplicacao.VB_ProcData.VB_Invoke_Func = " \n14"

' Tabela Hugo
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort. _
        SortFields.Add2 Key:=Range("Tabela11[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela11").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Luizinho
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort. _
        SortFields.Add2 Key:=Range("Tabela12[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela12").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Pedro
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort. _
        SortFields.Add2 Key:=Range("Tabela13[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela13").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Ramão
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort. _
        SortFields.Add2 Key:=Range("Tabela14[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela14").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Luiz Gonzaga
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort. _
        SortFields.Add2 Key:=Range("Tabela15[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela15").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Alex
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort. _
        SortFields.Add2 Key:=Range("Tabela16[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela16").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Dijalma
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort. _
        SortFields.Add2 Key:=Range("Tabela17[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela17").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Edson
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort. _
        SortFields.Add2 Key:=Range("Tabela18[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela18").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela PCM
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort. _
        SortFields.Add2 Key:=Range("Tabela19[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela19").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Laboratorio/Quimico
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort. _
        SortFields.Add2 Key:=Range("Tabela22[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela22").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Preditiva
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort. _
        SortFields.Add2 Key:=Range("Tabela20[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela20").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Eloir
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort. _
        SortFields.Add2 Key:=Range("Tabela21[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela21").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Daniel Barboza
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort. _
        SortFields.Add2 Key:=Range("Tabela23[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela23").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Tabela Marcelo Aparecido
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort. _
        SortFields.Add2 Key:=Range("Tabela24[[#All],[Aplicação]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Por Solicitante").ListObjects("Tabela24").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub
