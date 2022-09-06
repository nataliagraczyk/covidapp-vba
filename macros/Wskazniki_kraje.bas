Attribute VB_Name = "Wskazniki_kraje"
Public Kraj As String
Public Ilosc_przypadkow_K As Long, Ilosc_Zgonow_K As Long, Ilosc_Wyzdrowien_K As Long, Zaszczepieni_K As Long, Zaszczepieni_1_K As Long, Zaszczepieni_Calosc As Long
Public Przypadki_nowe_k As Long, Zgony_nowe_k As Long, Wyzdrowienia_nowe_k As Long
Public W_Zgonow_K As Double, W_Wyzdrowien_K As Double, W_Zaszczepieni_K As Double, W_Zaszczepieni_1_K As Double, W_Zaszczepieni_K_Calosc As Double
Public Miejsce_ogolne As Long, Miejsce_przypadki_nowe As Long, Miejsce_zgony As Long, Miejsce_zgony_nowe As Long, Miejsce_wyzdrowienia As Long, Miejsce_wyzdrowienia_nowe As Long, Miejsce_szczepienia As Long
Public Kraj_lista(1 To 8) As Long


'Wskazniki ogólne
Sub Licz_kraj_ogolne()

Ilosc_przypadkow_K = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Przypadki").Range("A1").CurrentRegion, 2)
Ilosc_Wyzdrowien_K = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Przypadki").Range("A1").CurrentRegion, 3)
Ilosc_Zgonow_K = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Przypadki").Range("A1").CurrentRegion, 4)
Zaszczepieni_K = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Vaccinated").Range("A1").CurrentRegion, 3)
Zaszczepieni_1_K = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Vaccinated").Range("A1").CurrentRegion, 4)
Zaszczepieni_Calosc = Zaszczepieni_K + Zaszczepieni_1


W_Zgonow_K = Ilosc_Zgonow_K / Ilosc_przypadkow_K 'Zgony/L.przypadków
W_Wyzdrowien_K = Ilosc_Wyzdrowien_K / Ilosc_przypadkow_K 'Wyzdrowienia/L.przypadków

'Zaszczepieni/L.ludnoœci
W_Zaszczepieni_Calosc_K = Zaszczepieni_Calosc / Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Przypadki").Range("A1").CurrentRegion, 5)
W_Zaszczepieni_K = Zaszczepieni_K / Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Przypadki").Range("A1").CurrentRegion, 5)
W_Zaszczepieni_1_K = Zaszczepieni_1_K / Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Przypadki").Range("A1").CurrentRegion, 5)

'Nowe
Przypadki_nowe_k = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("E2").CurrentRegion, 2)
Zgony_nowe_k = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("M2").CurrentRegion, 2)
Wyzdrowienia_nowe_k = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("U2").CurrentRegion, 2)

'Ranking
Miejsce_ogolne = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("A2").CurrentRegion, 3)
Miejsce_przypadki_nowe = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("E2").CurrentRegion, 3)
Miejsce_zgony = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("I2").CurrentRegion, 3)
Miejsce_zgony_nowe = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("M2").CurrentRegion, 3)
Miejsce_wyzdrowienia = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("Q2").CurrentRegion, 3)
Miejsce_wyzdrowienia_nowe = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("U2").CurrentRegion, 3)
Miejsce_szczepienia = Application.WorksheetFunction.VLookup(Kraj, ThisWorkbook.Worksheets("Pomocniczy_rankingi").Range("Y2").CurrentRegion, 3)

'Wpisanie danych do listy
Kraj_lista(1) = Ilosc_przypadkow_K
Kraj_lista(2) = Przypadki_nowe_k
Kraj_lista(3) = Ilosc_Zgonow_K
Kraj_lista(4) = Zgony_nowe_k
Kraj_lista(5) = Ilosc_Wyzdrowien_K
Kraj_lista(6) = Wyzdrowienia_nowe_k
Kraj_lista(7) = Zaszczepieni_Calosc
Kraj_lista(8) = Zaszczepieni_K

End Sub
