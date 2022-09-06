Attribute VB_Name = "Kalendarz_ogolne"
Public DataKalendarz As Date 'zmienna publiczba - wybrana data z kalendarza
'Kod pochodzi ze strony labmasters.pl

Sub Pokaz_kalendarz()
Dim Kolumna As Long, Liczba_dni As Long
Dim Tablica_daty() As String
Dim Data As String

Liczba_dni = ThisWorkbook.Worksheets("H_confirmed").Range("B1").CurrentRegion.Columns.count - 1
'Stworzenie pomocniczej tablicy do przeszukiwania dat
ReDim Tablica_daty(1 To Liczba_dni)
    'Wczytanie dat do tablicy
    For i = 1 To UBound(Tablica_daty)
        Tablica_daty(i) = Right(ThisWorkbook.Worksheets("H_deaths").Cells(1, i + 1), 10)
    Next i
'Sprawdzenie ostatniej daty aktualizacji danych
rok = Left(Tablica_daty(1), 4)
miesiac = Mid(Tablica_daty(1), 6, 2)
dzien = Right(Tablica_daty(1), 2)


'Wyœwietla formularz UserForm Kalendarz
Ponow:
    Kalendarz.Show 'wyœwietlenie

'Sprawdzenie czy wybrano poprawn¹ datê
If DataKalendarz <> 0 Then
    If DataKalendarz < DateSerial(2020, 1, 22) Or _
            DataKalendarz > DateSerial(rok, miesiac, dzien) Then
        MsgBox "Wybrano niepoprawn¹ datê!" & vbNewLine & vbNewLine & _
            "Wybierz datê z przedzia³u od 22.01.2020 do " & Tablica_daty(1), vbCritical + vbInformation, "Brak danych"
            DataKalendarz = Empty
        GoTo Ponow
    End If
Else
    Unload Kalendarz
    Exit Sub
End If
    
'Znalezienie odpowiedniego elementu w tablicy i przypisanie indeksu do kolumny
Data = Year(DataKalendarz) & "-" & Mid(DataKalendarz, 4, 2) & "-" & Left(DataKalendarz, 2)
    For i = 1 To UBound(Tablica_daty)
        If DataKalendarz = Tablica_daty(i) Then
            Kolumna = i + 1
            Exit For
        End If
    Next i
    
    Sheets("COUNTRY").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
        
'Wpisanie odpowiednuch danych do raportu
    Sheets("RAPORT").Unprotect
    Sheets("REPORT").Unprotect


    ThisWorkbook.Worksheets("REPORT").Range("H33") = Data
    ThisWorkbook.Worksheets("REPORT").Range("B39") = ThisWorkbook.Worksheets("H_confirmed").Cells(Rows.count, Kolumna).End(xlUp).Value
    ThisWorkbook.Worksheets("REPORT").Range("K40") = ThisWorkbook.Worksheets("H_recovered").Cells(Rows.count, Kolumna).End(xlUp).Value
    ThisWorkbook.Worksheets("REPORT").Range("F40") = ThisWorkbook.Worksheets("H_deaths").Cells(Rows.count, Kolumna).End(xlUp).Value
    
    ThisWorkbook.Worksheets("RAPORT").Range("H33") = Data
    ThisWorkbook.Worksheets("RAPORT").Range("B39") = ThisWorkbook.Worksheets("H_confirmed").Cells(Rows.count, Kolumna).End(xlUp).Value
    ThisWorkbook.Worksheets("RAPORT").Range("K40") = ThisWorkbook.Worksheets("H_recovered").Cells(Rows.count, Kolumna).End(xlUp).Value
    ThisWorkbook.Worksheets("RAPORT").Range("F40") = ThisWorkbook.Worksheets("H_deaths").Cells(Rows.count, Kolumna).End(xlUp).Value
    DataKalendarz = Empty
    
     Sheets("REPORT").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
    
     Sheets("RAPORT").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
End Sub





