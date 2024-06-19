Attribute VB_Name = "AreaImpressao"
Sub AreaImpressao()
Attribute AreaImpressao.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim AHugoOkahara, AJoseLuiz, APedroFilho, ARamaoFrancisco, ALuizGonzaga, AWillianJefrei, ADijalmaBatista, AEdsonSiqueira, APCM, ALaboratorio, APreditiva, AEloirDavila, ADanielBarboza, AMarceloAparecido As Integer
    
'As variaveis contem o nome dos encarregados abreviados.
    AHugoOkahara = SelecionarArea(8)
    AJoseLuiz = SelecionarArea(15)
    APedroFilho = SelecionarArea(22)
    ARamaoFrancisco = SelecionarArea(29)
    ALuizGonzaga = SelecionarArea(36)
    AWillianJefrei = SelecionarArea(43)
    ADijalmaBatista = SelecionarArea(50)
    AEdsonSiqueira = SelecionarArea(57)
    APCM = SelecionarArea(64)
    ALaboratorio = SelecionarArea(71)
    APreditiva = SelecionarArea(78)
    AEloirDavila = SelecionarArea(85)
    ADanielBarboza = SelecionarArea(92)
    AMarceloAparecido = SelecionarArea(99)
    
    ActiveSheet.PageSetup.PrintArea = "$A$4:$F$25" & _
    ",$DB$4:$DF$19" & _
    ",$H$4:$M$" & AHugoOkahara & _
    ",$O$4:$T$" & AJoseLuiz & _
    ",$V$4:$AA$" & APedroFilho & _
    ",$AC$4:$AH$" & ARamaoFrancisco & _
    ",$AJ$4:$AO$" & ALuizGonzaga & _
    ",$AQ$4:$AV$" & AWillianJefrei & _
    ",$CG$4:$CL$" & AEloirDavila & _
    ",$CN$4:$CS$" & ADanielBarboza & _
    ",$CU$4:$CZ$" & AMarceloAparecido & _
    ",$AX$4:$BC$" & ADijalmaBatista & _
    ",$BE$4:$BJ$" & AEdsonSiqueira & _
    ",$BZ$4:$CE$" & APreditiva & _
    ",$BL$4:$BQ$" & APCM & _
    ",$BS$4:$BX$" & ALaboratorio
End Sub

Sub AreaImpressaoEntressafra()
    Dim AHugoOkahara, AJoseLuiz, APedroFilho, ARamaoFrancisco, ALuizGonzaga, AWillianJefrei, ADijalmaBatista, AEdsonSiqueira, APCM, ALaboratorio, APreditiva, AEloirDavila, ADanielBarboza, AMarceloAparecido As Integer
    
'As variaveis contem o nome dos encarregados abreviados.
    AHugoOkahara = SelecionarArea(8)
    AJoseLuiz = SelecionarArea(15)
    APedroFilho = SelecionarArea(22)
    ARamaoFrancisco = SelecionarArea(29)
    ALuizGonzaga = SelecionarArea(36)
    AWillianJefrei = SelecionarArea(43)
    ADijalmaBatista = SelecionarArea(50)
    AEdsonSiqueira = SelecionarArea(57)
    APCM = SelecionarArea(64)
    ALaboratorio = SelecionarArea(71)
    APreditiva = SelecionarArea(78)
    AEloirDavila = SelecionarArea(85)
    ADanielBarboza = SelecionarArea(92)
    AMarceloAparecido = SelecionarArea(99)
    
    ActiveSheet.PageSetup.PrintArea = "$A$27:$F$47" & _
    ",$DB$4:$DF$19" & _
    ",$H$4:$M$" & AHugoOkahara & _
    ",$O$4:$T$" & AJoseLuiz & _
    ",$V$4:$AA$" & APedroFilho & _
    ",$AC$4:$AH$" & ARamaoFrancisco & _
    ",$AJ$4:$AO$" & ALuizGonzaga & _
    ",$AQ$4:$AV$" & AWillianJefrei & _
    ",$CG$4:$CL$" & AEloirDavila & _
    ",$CN$4:$CS$" & ADanielBarboza & _
    ",$CU$4:$CZ$" & AMarceloAparecido & _
    ",$AX$4:$BC$" & ADijalmaBatista & _
    ",$BE$4:$BJ$" & AEdsonSiqueira & _
    ",$BZ$4:$CE$" & APreditiva & _
    ",$BL$4:$BQ$" & APCM & _
    ",$BS$4:$BX$" & ALaboratorio
End Sub

Function SelecionarArea(p1 As Integer)

    Dim linha As Integer
    
    linha = 4
    
    Planilha3.Cells(linha, p1).Select

    Do While Planilha3.Cells(linha + 1, p1) <> ""
    
        If Planilha3.Cells(linha + 1, p1 + 5).Value = "ENTREGUE" Then
            GoTo Fim
        Else
            Selection.Resize(Selection.Rows.Count + 1).Select
        
            linha = linha + 1
        End If
    Loop
    
Fim:
    
    SelecionarArea = linha

End Function
