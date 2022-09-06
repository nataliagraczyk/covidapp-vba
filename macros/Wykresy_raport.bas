Attribute VB_Name = "Wykresy_raport"
    Public WybranyKraj As String, Kontynent As String, WybranyWskaznik As String
    Public PozycjaKrajuSlownik As Long, OstatniWiersz As Long, OstatniaKolumna As Long
    Public ListaKrajowZakres As Range, DatyBazaZakres As Range

Sub WykresyKraje_raport()

    Application.ScreenUpdating = False
    WybranyKraj = StrConv(Sheets("COUNTRY").Range("B6").Value, 3)
    
    ''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''Confirmed''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''
    
    Sheets("Dictionary").Activate
    Kontynent = Application.WorksheetFunction.VLookup(WybranyKraj, Sheets("Dictionary").Range("A2:B500"), 2, 0)
    PozycjaKrajuSlownik = Application.WorksheetFunction.Match(Kontynent, Sheets("Dictionary").Range("A1:N1"), 0)
    PozycjaKraju2 = Application.WorksheetFunction.Match(WybranyKraj, Sheets("H_confirmed").Range("A1:A1000"), 0)
    
    OstatniWiersz = Sheets("Dictionary").Cells(1, PozycjaKrajuSlownik).End(xlDown).Row
    OstatniaKolumna = Sheets("H_confirmed").Range("A1").End(xlToRight).Column
    
    Sheets("H_confirmed").Activate
    Set DatyBazaZakres = Range(Cells(1, 2), Cells(1, OstatniaKolumna))
    
    Dim DatyBaza As Variant
    DatyBaza = DatyBazaZakres.Value
    
    Sheets("Wykresy").Activate

    
    For i = 1 To OstatniaKolumna - 1
        
        Dim DatyWykresOSX As Date
        DatyWykresOSX = Format(CDate(Mid(DatyBaza(1, i), 17, 10)), "YYYY-MM-DD")
        DatyBaza(1, i) = DatyWykresOSX
                        
    Next i
    
    Dim NrKol As Long: NrKol = 2
        
    Sheets("Dictionary").Activate
    Set ListaKrajowZakres = Sheets("Dictionary").Range(Cells(2, PozycjaKrajuSlownik), Cells(OstatniWiersz, PozycjaKrajuSlownik))
    
    Dim ListaKrajow As Variant, WartosciKrajow As Variant
    
    ListaKrajow = ListaKrajowZakres.Value
    ReDim WartosciKrajow(1 To UBound(ListaKrajow))
    
    For i = LBound(ListaKrajow) To UBound(ListaKrajow)
    
        Dim abcd As Long
        abcd = Application.WorksheetFunction.VLookup(ListaKrajow(i, 1), Sheets("H_confirmed").Range("A1:C1000"), k, 0)
        WartosciKrajow(i) = abcd
        
    Next
    
    'Wykres Liniowy
    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_liniowy_confirmed").Activate
    ActiveChart.FullSeriesCollection(1).Values = "=H_confirmed!" & Range(Cells(PozycjaKraju2, 2), Cells(PozycjaKraju2, OstatniaKolumna)).Address
    ActiveChart.FullSeriesCollection(1).XValues = DatyBaza
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
    ActiveChart.Axes(xlValue).TickLabelPosition = xlHigh
    Application.CommandBars("Format Object").Visible = False

    
    'Wykres kolumnowy
    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_slupkowy_confirmed").Activate
    ActiveChart.ChartTitle.Text = "Countries in " & Kontynent
    ActiveChart.FullSeriesCollection(1).Values = WartosciKrajow
    ActiveChart.FullSeriesCollection(1).XValues = ListaKrajow
    
    ''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''Deaths'''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''
    
    'Wykres Liniowy
    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_liniowy_deaths").Activate
    ActiveChart.FullSeriesCollection(1).Values = "=H_deaths!" & Range(Cells(PozycjaKraju2, 2), Cells(PozycjaKraju2, OstatniaKolumna)).Address
    ActiveChart.FullSeriesCollection(1).XValues = DatyBaza
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
    ActiveChart.Axes(xlValue).TickLabelPosition = xlHigh
    Application.CommandBars("Format Object").Visible = False

    
    'Wykres kolumnowy
    
    Sheets("H_deaths").Activate
    For i = LBound(ListaKrajow) To UBound(ListaKrajow)
    
        abcd = Application.WorksheetFunction.VLookup(ListaKrajow(i, 1), Sheets("H_deaths").Range("A1:C1000"), NrKol, 0)
        WartosciKrajow(i) = abcd
        
    Next
    
    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_slupkowy_deaths").Activate
    ActiveChart.ChartTitle.Text = "Countries in " & Kontynent
    ActiveChart.FullSeriesCollection(1).Values = WartosciKrajow
    ActiveChart.FullSeriesCollection(1).XValues = ListaKrajow
    
    ''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''Recovered''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''

    'Wykres Liniowy
    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_liniowy_recovered").Activate
    ActiveChart.FullSeriesCollection(1).Values = "=H_recovered!" & Range(Cells(PozycjaKraju2, 2), Cells(PozycjaKraju2, OstatniaKolumna)).Address
    ActiveChart.FullSeriesCollection(1).XValues = DatyBaza
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
    ActiveChart.Axes(xlValue).TickLabelPosition = xlHigh
    Application.CommandBars("Format Object").Visible = False

    
    'Wykres kolumnowy
    
    Sheets("H_recovered").Activate
    For i = LBound(ListaKrajow) To UBound(ListaKrajow)
    
        abcd = Application.WorksheetFunction.VLookup(ListaKrajow(i, 1), Sheets("H_recovered").Range("A1:C1000"), NrKol, 0)
        WartosciKrajow(i) = abcd
        
    Next
    
    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_slupkowy_recovered").Activate
    ActiveChart.ChartTitle.Text = "Countries in " & Kontynent
    ActiveChart.FullSeriesCollection(1).Values = WartosciKrajow
    ActiveChart.FullSeriesCollection(1).XValues = ListaKrajow
    
    '''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''Vaccinated''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''
    
    'Wykres kolumnowy
        Sheets("Vaccinated").Activate
    For i = LBound(ListaKrajow) To UBound(ListaKrajow)
        
        On Error Resume Next
        abcd = Application.WorksheetFunction.VLookup(ListaKrajow(i, 1), Sheets("Vaccinated").Range("A1:C1000"), 3, 0)
        On Error GoTo 0
        WartosciKrajow(i) = abcd

    Next

    Sheets("wykresy").Activate
    ActiveSheet.ChartObjects("EkranKraj_slupkowy_vaccinated").Activate
    ActiveChart.ChartTitle.Text = "Countries in " & Kontynent
    ActiveChart.FullSeriesCollection(1).Values = WartosciKrajow
    ActiveChart.FullSeriesCollection(1).XValues = ListaKrajow
    

End Sub

