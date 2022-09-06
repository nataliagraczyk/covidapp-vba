Attribute VB_Name = "Metryczka"
Public Country As String

Sub Metryczka1()
'makro uzupe³niaj¹ce metryczkê
'Arkusz COUNTRY
'Dim Country As String
Dim All As Variant
Dim Tabela As Range
Dim Tabela2 As Range
Dim Capital As String
Dim Active As Variant
Dim Recovered As Variant
Dim Deaths As Variant
Dim Vaccinated As Variant
Dim Continent As String
Dim Population As Variant
Dim Area As Variant
Dim LifeE As Variant
Dim Latitude As Variant
Dim Longitude As Variant
Dim Kraj As String
Dim Dictionary As Range
Dim ThisSheet As Worksheet
Dim Kontynent As String
Dim TK As Range
Dim TK2 As Range

Sheets("KRAJ").Unprotect
Sheets("COUNTRY").Unprotect


Set Dictionary = Sheets("Dictionary").Range("R1:S194")
Set ThisSheet = ActiveSheet
Set TK = Sheets("Dictionary").Range("AB1:AC6")
Set TK2 = Sheets("Dictionary").Range("Q1:R193")
Set Tabela = Sheets("Przypadki").Range("A1:M194")
Set Tabela2 = Sheets("Vaccinated").Range("A1:E194")
 
    Kraj = Sheets("KRAJ").Range("B6").Value
    Kontynent = Application.WorksheetFunction.VLookup(Kraj, Tabela, 8, 0)
    Continent = Application.WorksheetFunction.VLookup(Kontynent, TK, 2, 0)
    Country = Application.WorksheetFunction.VLookup(Kraj, TK2, 2, 0)


Set Tabela = Sheets("Przypadki").Range("A1:M194")
Set Tabela2 = Sheets("Vaccinated").Range("A1:E194")


All = Application.WorksheetFunction.VLookup(Kraj, Tabela, 2, 0)
Recovered = Application.WorksheetFunction.VLookup(Kraj, Tabela, 3, 0)
Active = All - Recovered
Deaths = Application.WorksheetFunction.VLookup(Kraj, Tabela, 4, 0)
Capital = Application.WorksheetFunction.VLookup(Kraj, Tabela, 10, 0)
On Error Resume Next
Vaccinated = Application.WorksheetFunction.VLookup(Kraj, Tabela2, 3, 0)
On Error GoTo 0
Population = Application.WorksheetFunction.VLookup(Kraj, Tabela, 5, 0)
Area = Application.WorksheetFunction.VLookup(Kraj, Tabela, 6, 0)
LifeE = Application.WorksheetFunction.VLookup(Kraj, Tabela, 7, 0)
Latitude = Application.WorksheetFunction.VLookup(Kraj, Tabela, 11, 0)
Longitude = Application.WorksheetFunction.VLookup(Kraj, Tabela, 12, 0)


'Sheets("KRAJ").Range("A1").Select
Sheets("KRAJ").Range("B6").Value = Country
Sheets("KRAJ").Range("E9").Value = Continent
Sheets("KRAJ").Range("C13").Value = All
Sheets("KRAJ").Range("C19").Value = Recovered
Sheets("KRAJ").Range("G13").Value = Active
Sheets("KRAJ").Range("G19").Value = Deaths
Sheets("KRAJ").Range("L19").Value = Vaccinated
Sheets("KRAJ").Range("I28").Value = Capital
Sheets("KRAJ").Range("I30").Value = Population
Sheets("KRAJ").Range("I32").Value = Area
Sheets("KRAJ").Range("I34").Value = LifeE
Sheets("KRAJ").Range("I36").Value = Latitude
Sheets("KRAJ").Range("I38").Value = Longitude


Sheets("COUNTRY").Range("E9").Value = Kontynent
Sheets("COUNTRY").Range("C13").Value = All
Sheets("COUNTRY").Range("C19").Value = Recovered
Sheets("COUNTRY").Range("G13").Value = Active
Sheets("COUNTRY").Range("G19").Value = Deaths
Sheets("COUNTRY").Range("L19").Value = Vaccinated
Sheets("COUNTRY").Range("I28").Value = Capital
Sheets("COUNTRY").Range("I30").Value = Population
Sheets("COUNTRY").Range("I32").Value = Area
Sheets("COUNTRY").Range("I34").Value = LifeE
Sheets("COUNTRY").Range("I36").Value = Latitude
Sheets("COUNTRY").Range("I38").Value = Longitude

        Sheets("KRAJ").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
    
        Sheets("COUNTRY").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True

End Sub


