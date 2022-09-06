Attribute VB_Name = "Ustawienia"
Sub OtworzZamknijUstawienia()
'makro otwieraj¹ce/zamykaj¹ce ustawienia

Application.ScreenUpdating = False
    ActiveSheet.Unprotect

    If ActiveSheet.Columns("AE:AI").Hidden = True Then
        ActiveSheet.Columns("AE:AI").Hidden = False
    Else
        ActiveSheet.Columns("AE:AI").Hidden = True
    End If
    
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False

    ActiveSheet.Range("A1:AI48").Select
    ActiveWindow.Zoom = True

    Range("A1").Select
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
        
Application.ScreenUpdating = True
    
End Sub

Sub JezykPolski()
' makro zmieniaj¹ce jêzyk aplikacji na polski
    Application.ScreenUpdating = False
    
    Sheets("RAPORT").Visible = True
    Sheets("RAPORT").Select


    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False

    ActiveSheet.Range("A1:AI48").Select
    ActiveWindow.Zoom = True

    Range("A1").Select
    
    Sheets("RAPORT").Select
    Sheets("REPORT").Visible = False
    
    Application.ScreenUpdating = True
   
End Sub

Sub JezykAngielski()
'makro zmieniaj¹ce jêzyk aplikacji na angielski
    Application.ScreenUpdating = False
    
    Sheets("REPORT").Visible = True
    Sheets("REPORT").Select

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False

    ActiveSheet.Range("A1:AI48").Select
    ActiveWindow.Zoom = True

    Range("A1").Select
    
    Sheets("REPORT").Select
    Sheets("RAPORT").Visible = False
    
    
    Application.ScreenUpdating = True

End Sub

Sub JezykPolski_KRAJ()
'makro zmieniaj¹ce jêzyk na polski w arkuszu KRAJ
    Application.ScreenUpdating = False

    Sheets("KRAJ").Visible = True
    Sheets("KRAJ").Select

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False

    ActiveSheet.Range("A1:AI48").Select
    ActiveWindow.Zoom = True

    Range("A1").Select
    
    Sheets("KRAJ").Select
    Sheets("COUNTRY").Visible = False
    
    Application.ScreenUpdating = True
   
End Sub

Sub JezykAngielski_KRAJ()
'makro zmieniaj¹ce jêzyk na angielski w arkuszu KRAJ
    Application.ScreenUpdating = False
    
    Sheets("COUNTRY").Visible = True
    Sheets("COUNTRY").Select

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False

    ActiveSheet.Range("A1:AI48").Select
    ActiveWindow.Zoom = True

    Range("A1").Select
    
    Sheets("COUNTRY").Select
    Sheets("KRAJ").Visible = False
    
    Application.ScreenUpdating = True

End Sub
'FUNKCJA POMOCNICZA
Function IsIn(element, arr) As Boolean
    IsIn = False
    For Each X In arr
        If element = X Then
            IsIn = True
            Exit Function
        End If
    Next X
End Function

Sub MoveBtn_raport()
'makro przypisane do przycisku, zmieniaj¹ce motyw aplikacji
Dim adres As String
Dim xSh As Worksheet
Dim ThisSheet As Worksheet
Dim chosenSheets()

Application.ScreenUpdating = False

adres = ActiveSheet.Shapes("Mve").BottomRightCell.Address

Set ThisSheet = ActiveSheet

    chosenSheets = Array("RAPORT", "REPORT", "KRAJ", "COUNTRY")

For Each xSh In ActiveWorkbook.Worksheets

    If IsIn(xSh.Name, chosenSheets) Then

    With xSh.Shapes("Mve")
    
    If .BottomRightCell.Address = "$AH$25" Then
    .IncrementLeft -26
    xSh.Range("AF24").Font.ThemeColor = xlThemeColorLight1
    xSh.Range("AH24").Font.ThemeColor = xlThemeColorDark1
    
    'jasny
    
        With xSh.Range("A1:AD48").Interior
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
        End With


    Else
    .IncrementLeft 26
    xSh.Range("AH24").Font.ThemeColor = xlThemeColorLight1
    xSh.Range("AF24").Font.ThemeColor = xlThemeColorDark1
    
    'ciemny
    xSh.Range("A1:AD48").Interior.ThemeColor = xlThemeColorLight1
    

    End If
    
    End With
    
    End If
    
Next xSh


ThisSheet.Range("AI1").Select
    
Application.ScreenUpdating = True

End Sub
'pomocnicze
Sub GetActiveSheetIndex()
MsgBox ActiveSheet.Index
End Sub

Sub ustawienie_startowe_przyciskow()
Dim xSh As Worksheet
Dim chosenSheets()
'makro zmieniaj¹ce ustawienia startowe przycisków
chosenSheets = Array("RAPORT", "REPORT", "KRAJ", "COUNTRY")

For Each xSh In ActiveWorkbook.Worksheets

    If IsIn(xSh.Name, chosenSheets) Then
    
 With xSh.Shapes("Mve")
    
    If .BottomRightCell.Address = "$AH$25" Then

    Else
    .IncrementLeft 26
    xSh.Range("AH24").Font.ThemeColor = xlThemeColorLight1
    xSh.Range("AF24").Font.ThemeColor = xlThemeColorDark1
    
    'ciemny
    xSh.Range("A1:AD48").Interior.ThemeColor = xlThemeColorLight1
    

    End If
    
    End With

    End If

Next xSh

End Sub

Sub Instrukcja()
'makro otwieraj¹ce instrukcje PDF
Dim ThisSheet As Worksheet
Dim Path As String

Set ThisSheet = ActiveSheet
   
    If ThisSheet.Name = "KRAJ" Then
        Path = ThisWorkbook.Path & "\Szablony\instrukcja.pdf"
    ElseIf ThisSheet.Name = "RAPORT" Then
        Path = ThisWorkbook.Path & "\Szablony\instrukcja.pdf"
    ElseIf ThisSheet.Name = "COUNTRY" Then
        Path = ThisWorkbook.Path & "\Szablony\instruction.pdf"
    ElseIf ThisSheet.Name = "REPORT" Then
        Path = ThisWorkbook.Path & "\Szablony\instruction.pdf"
    End If

ThisWorkbook.FollowHyperlink Address:=Path, NewWindow:=True

End Sub

Sub HomePage()
'makro wracaj¹ce do strony g³ównej (przypisane do strza³ki w arkuszu KRAJ/COUNTRY)
Set ThisSheet = ActiveSheet
   
    If ThisSheet.Name = "KRAJ" Then
        Sheets("REPORT").Visible = False
        Sheets("COUNTRY").Visible = False
        Sheets("RAPORT").Visible = True
        Sheets("RAPORT").Select
        'Sheets("
    ElseIf ThisSheet.Name = "COUNTRY" Then
        Sheets("REPORT").Visible = True
        Sheets("RAPORT").Visible = False
        Sheets("KRAJ").Visible = False
        Sheets("REPORT").Select
    Else: Exit Sub
    End If

End Sub
