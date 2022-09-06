Attribute VB_Name = "Wskazniki_ogolne1"
Public Ilosc_Przypadkow As Long, Ilosc_Zgonow As Long, Ilosc_Wyzdrowien As Long, Ilosc_szczepien As String, Szczepienia_pelne As Long, Szczepienia_1 As Long
Public W_Nowych As Double, W_Zgonow As Double, W_Wyzdrowien As Double
Public W_Zgonow_D As Double, W_Wyzdrowien_D As Double
Public Przypadki_Nowe As Long, Zgony_Nowe As Long, Wyzdrowienia_Nowe As Long


'Podstawowe wskaŸniki dla danych ogólnych

Sub Licz_ogolne()

'Iloœæ przypadków, wyzdrowieñ i œmierci
Ilosc_Przypadkow = ThisWorkbook.Worksheets("Przypadki").Cells(Rows.count, 2).End(xlUp).Value
Ilosc_Wyzdrowien = ThisWorkbook.Worksheets("Przypadki").Cells(Rows.count, 3).End(xlUp).Value
Ilosc_Zgonow = ThisWorkbook.Worksheets("Przypadki").Cells(Rows.count, 4).End(xlUp).Value
Ilosc_szczepien = ThisWorkbook.Worksheets("Vaccinated").Cells(Rows.count, 2).End(xlUp).Value
Szczepienia_pelne = ThisWorkbook.Worksheets("Vaccinated").Cells(Rows.count, 3).End(xlUp).Value
Szczepienia_1 = ThisWorkbook.Worksheets("Vaccinated").Cells(Rows.count, 4).End(xlUp).Value

Przypadki_Nowe = Ilosc_Przypadkow - ThisWorkbook.Worksheets("H_confirmed").Cells(Rows.count, 3).End(xlUp).Value
Wyzdrowienia_Nowe = Ilosc_Wyzdrowien - ThisWorkbook.Worksheets("H_recovered").Cells(Rows.count, 3).End(xlUp).Value
Zgony_Nowe = Ilosc_Zgonow - ThisWorkbook.Worksheets("H_deaths").Cells(Rows.count, 3).End(xlUp).Value

'WskaŸniki ca³oœciowe
W_Zgonow = Ilosc_Zgonow / Ilosc_Przypadkow
W_Wyzdrowien = Ilosc_Wyzdrowien / Ilosc_Przypadkow


'Wstawienie do raportu
        Sheets("RAPORT").Unprotect
        Sheets("REPORT").Unprotect
    
        ThisWorkbook.Worksheets("RAPORT").Range("P12").Value = "'+ " & Przypadki_Nowe
        ThisWorkbook.Worksheets("REPORT").Range("P12").Value = "'+ " & Przypadki_Nowe
        With Sheets("RAPORT").Range("P12")
            .Font.Color = RGB(255, 0, 0)
        End With
        With Sheets("REPORT").Range("P12")
            .Font.Color = RGB(255, 0, 0)
        End With
        
        ThisWorkbook.Worksheets("RAPORT").Range("P24").Value = "'+ " & Wyzdrowienia_Nowe
        ThisWorkbook.Worksheets("REPORT").Range("P24").Value = "'+ " & Wyzdrowienia_Nowe
         With Sheets("RAPORT").Range("P24")
            .Font.Color = RGB(0, 255, 0)
        End With
        
         With Sheets("REPORT").Range("P24")
            .Font.Color = RGB(0, 255, 0)
        End With
 
    
        ThisWorkbook.Worksheets("RAPORT").Range("U24").Value = "'+ " & Zgony_Nowe
        ThisWorkbook.Worksheets("REPORT").Range("U24").Value = "'+ " & Zgony_Nowe
         With Sheets("RAPORT").Range("U24")
            .Font.Color = RGB(255, 0, 0)
        End With
         
         With Sheets("REPORT").Range("U24")
            .Font.Color = RGB(255, 0, 0)
        End With
        
        Sheets("RAPORT").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
        
        Sheets("REPORT").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
  
End Sub

