Attribute VB_Name = "AtualizarRelatorio"
Sub FiltroAd()
Attribute FiltroAd.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' FiltroAd Macro
'
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Sheets("Dados").Range("DadosSC[#All]").AdvancedFilter Action:= _
        xlFilterCopy, CriteriaRange:=Range("A1:L2"), CopyToRange:=Range("A5:L5"), _
        Unique:=False
End Sub

Sub FiltroRelatorio()
Attribute FiltroRelatorio.VB_ProcData.VB_Invoke_Func = "h\n14"
    
    Application.ScreenUpdating = False

    ResetarRelatorio.ApagarRelatorio
    
    Sheets("Busca").Select
    Range("a2:k2").Select
    Selection.ClearContents

    'Hugo
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Hugo Okahara*aprovado"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("H6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("H:H")
    
    'Luizinho
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*José Luiz*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("O6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("O:O")
    
    'Pedro
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Pedro Filho*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("V6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("V:V")
    
    'Ramão
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Ramão Francisco*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AC6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AC:AC")
    
    'Pedreiro
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Luiz Gonzaga*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AJ6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AJ:AJ")
    
    'Willian Jefrei
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Willian Jefrei*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AQ6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AQ:AQ")
    
    'Fafatoba
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Dijalma Batista*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AX6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AX:AX")
    
    'Ihnhé
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Edson Siqueira*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BE6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BE:BE")
    
    'PCM
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*PCM*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BL6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BL:BL")
    
    'Preditiva
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*preditiva*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BZ6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BZ:BZ")
    
    'Eloir
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Eloir*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("CG6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("CG:CG")
    
    'Laboratório
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Natalia*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "<>0"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BS6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BS:BS")
    
    'Daniel
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Daniel Barboza*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("CN6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("CN:CN")
    
    'Marchello
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45355"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Marcelo Aparecido*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulas
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("CU6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("CU:CU")
    
    OrganizaAplicacao.ClassificaAplicacao
    OrganizarClassificacao.OrganizarClassificacao
    
    Sheets("Busca").Select
    Range("A2:L2").Select
    Selection.ClearContents
    Sheets("Por Solicitante").Select
    
    ArrumarLinhas.ArrumarLinhas
    
    AreaImpressao.AreaImpressao
    
    ExportarArquivo.ExportarSafra
    
    Application.ScreenUpdating = True
    
End Sub

Sub FiltroRelatorioEntressafra()
    
    Application.ScreenUpdating = False

    ResetarRelatorio.ApagarRelatorio
    
    Sheets("Busca").Select
    Range("a2:k2").Select
    Selection.ClearContents

    'Hugo
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Hugo Okahara*aprovado*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("H6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("H:H")
    
    'Luizinho
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*José Luiz*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("O6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("O:O")
    
    'Pedro
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Pedro Filho*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("V6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("V:V")
    
    'Ramão
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Ramão Francisco*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AC6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AC:AC")
    
    'Pedreiro
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Luiz Gonzaga*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AJ6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AJ:AJ")
    
    'Willian jefrei
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Willian Jefrei*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AQ6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AQ:AQ")
    
    'Fafatoba
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Dijalma Batista*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("AX6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("AX:AX")
    
    'Ihnhé
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Edson Siqueira*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BE6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BE:BE")
    
    'PCM
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*PCM*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BL6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BL:BL")
    
    'Preditiva
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*preditiva*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BZ6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BZ:BZ")
    
    'Eloir
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Eloir*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("CG6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("CG:CG")
    
    'Laboratório
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Natalia*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "<>0"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("BS6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("BS:BS")
    
    'Daniel
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Daniel Barboza*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("CN6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("CN:CN")
    
    'Marchello
    Sheets("Busca").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = ">=45231"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "<>SV"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "*Marcelo Aparecido*"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "p"
    FiltroAd
    SelecionarCelulas.SelecCelulasEntressafra
    Selection.Copy
    Sheets("Por Solicitante").Select
    Range("CU6").Select
    ActiveSheet.Paste
    AjustarFonte.AjustarFonte ("CU:CU")
    
    OrganizaAplicacao.ClassificaAplicacao
    OrganizarClassificacao.OrganizarClassificacao
    
    Sheets("Busca").Select
    Range("A2:L2").Select
    Selection.ClearContents
    Sheets("Por Solicitante").Select
    
    ArrumarLinhas.ArrumarLinhas
    
    AreaImpressao.AreaImpressaoEntressafra
    
    ExportarArquivo.ExportarEntressafra
    
    Application.ScreenUpdating = True
    
End Sub
