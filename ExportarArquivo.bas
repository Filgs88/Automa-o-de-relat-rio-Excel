Attribute VB_Name = "ExportarArquivo"
Sub ExportarSafra()
Attribute ExportarSafra.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "\\192.168.1.88\Users\PC\OneDrive - MSFT\PCM\01. PCMI\22. Servi�os e Materiais\Relat�rios\Relatorio de Solicita��es\Relat�rio de SC pendentes safra.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
End Sub

Sub ExportarEntressafra()

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "\\192.168.1.88\Users\PC\OneDrive - MSFT\PCM\01. PCMI\22. Servi�os e Materiais\Relat�rios\Relatorio de Solicita��es\Relat�rio de SC pendentes entressafra.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
End Sub

