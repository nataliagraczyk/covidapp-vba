Attribute VB_Name = "ShowUF"
Sub UF_Welcome()
    
    WelcomeUF.Show

End Sub

Sub UF_Exit()

    ExitUF.Show

End Sub

Sub UF_Raporty()

    RaportUF.Show

End Sub

Sub ochrona()
Sheets("KRAJ").Unprotect
Sheets("COUNTRY").Unprotect

End Sub

Sub abc()
ActiveWindow.DisplayWorkbookTabs = True

End Sub
