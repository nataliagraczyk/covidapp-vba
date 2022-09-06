Attribute VB_Name = "Wykresy"
Dim WybranyKraj As String, Kontynent As String, WybranyWskaznik As String, Kraj As String
Dim PozycjaKrajuSlownik As Long, OstatniWiersz As Long, OstatniaKolumna As Long
Dim ListaKrajowZakres As Range, DatyBazaZakres As Range
Dim LiczbaOdkrytych As Integer

Sub UF_PokazRaporty()
'makro otwieraj¹ce UF do wyœwietlenia zapisanych raportów
    ShowReportsUF.Show

End Sub

Sub WykresyKrajeANG()
'makro uzupe³niaj¹ce wykres liniowy oraz kolumnowy - wersja angielska
    On Error Resume Next
    ActiveSheet.Shapes.Range(Array("OkienkoError")).Delete
    On Error GoTo 0

    Application.ScreenUpdating = False
    
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''''''PRZYGOTOWANIE ZMIENNYCH'''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    
    'Pobieramy nazwê kraju z komórki 'B6' z okna krajowego
    WybranyKraj = Sheets("COUNTRY").Range("B6").Value

    'Dostosowujemy nazwê wybranego wskaznika do nazwy zak³adki w arkuszu
    If Sheets("COUNTRY").CB_Indicator.Value = "" Then Exit Sub
    WybranyWskaznik = "H_" & LCase(Sheets("COUNTRY").CB_Indicator.Value)
    
    If WybranyWskaznik = "H_vaccinated" Then
        WybranyWskaznik = "Vaccinated"
    End If
    
    'Wyszukujemy nazwy kontynentu oraz umiejscowienie wiersza z danym krajem w s³owniku
    Kontynent = Application.WorksheetFunction.VLookup(WybranyKraj, Sheets("Dictionary").Range("A2:B500"), 2, 0)
    PozycjaKrajuSlownik = Application.WorksheetFunction.Match(Kontynent, Sheets("Dictionary").Range("A1:N1"), 0)
    PozycjaKraju2 = Application.WorksheetFunction.Match(WybranyKraj, Sheets(WybranyWskaznik).Range("A1:A1000"), 0)
    
    OstatniWiersz = Sheets("Dictionary").Cells(1, PozycjaKrajuSlownik).End(xlDown).Row
    OstatniaKolumna = Sheets(WybranyWskaznik).Range("A1").End(xlToRight).Column
    
    Sheets(WybranyWskaznik).Activate
    Set DatyBazaZakres = Range(Cells(1, 2), Cells(1, OstatniaKolumna))
    
    Dim DatyBaza As Variant
    DatyBaza = DatyBazaZakres.Value
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''WYKRES LINIOWY Z DANYMI HISTORYCZNYMI'''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Je¿eli wybranym wskaznikiem jest 'vaccinated' to robimy obejœcie w postaci zakrycia wykresu czrn¹ ramk¹ prostk¹tn¹
    'z informacj¹ o tym, ¿e nie ma danym historczynych do wykresu liniowego
    Sheets("Wykresy").Activate
    If WybranyWskaznik = "Vaccinated" Then
    
        'Usuwamy ramkê jeœli by³a (np przy poprzednim ustawieniu)
        Sheets("COUNTRY").Activate
        On Error Resume Next
        ActiveSheet.Shapes.Range(Array("OkienkoError")).Delete
        
        'Tworzymy ramkê o dok³adnie zmierzonych wymiarach do zakrycia wykresu
        ActiveSheet.Shapes.AddShape(msoShapeRectangle, 66.25, 266.25, 583.75, 311.25). _
            Select
        Selection.ShapeRange.Name = "OkienkoError"
        Selection.ShapeRange("OkienkoError").IncrementLeft 688.75
        Selection.ShapeRange("OkienkoError").IncrementTop 137.5
        Selection.ShapeRange("OkienkoError").ShapeStyle = msoShapeStylePreset8
        Selection.ShapeRange("OkienkoError").TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange("OkienkoError").TextFrame2.TextRange.Characters.Text = "No daily data found. Category 'Vaccinated' has only total number of cases"
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 7). _
            ParagraphFormat
            .FirstLineIndent = 0
            .Alignment = msoAlignCenter
        End With
        
        GoTo Dalej
    
    Else
        
        'Jeœli u¿ytkownik nie wybra³ kategorii 'vaccinated' to po prostu usuwamy ramkê jeœli jest
        On Error Resume Next
        ActiveSheet.Shapes.Range(Array("OkienkoError")).Delete
    
    End If
    
    'Tworzymy oœ X wykresu liniowego, która zawiera daty
    For i = 1 To OstatniaKolumna - 1
        
        Dim DatyWykresOSX As Date
        DatyWykresOSX = Format(CDate(Mid(DatyBaza(1, i), 17, 10)), "YYYY-MM-DD")
        DatyBaza(1, i) = DatyWykresOSX
                        
    Next
    
    'Podajemy dane do wykresu liniowego o nazwie "EkranKraj_liniowy", który znajduje siê ju¿ na ekranie kraju
    Sheets("COUNTRY").Activate
    ActiveSheet.ChartObjects("EkranKraj_liniowy").Activate
    ActiveChart.FullSeriesCollection(1).Values = "=" & WybranyWskaznik & "!" & Range(Cells(PozycjaKraju2, 2), Cells(PozycjaKraju2, OstatniaKolumna)).Address 'dane dotycz¹ce wybranego wskaŸnika
    ActiveChart.FullSeriesCollection(1).XValues = DatyBaza 'daty na osi X
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True 'daty wystêpuj¹ w bazie w odwrotnej kolejnoœci, wiêc REversePlotOrder musi byæ ustawiony na 'True'
    ActiveChart.Axes(xlValue).TickLabelPosition = xlHigh
    Application.CommandBars("Format Object").Visible = False

Dalej:
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''WYKRES KOLUMNOWY: KRAJ NA TLE KONTYNENTU'''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Ustawiamy numer kolumny danych w bazie - dla wszystkich wskaŸników poza 'vaccinated' pobieramy dane z 2 kolumny,
    'w przypadku 'vaccinated' z trzeciej
    Dim NrKol As Long: NrKol = 2
    Dim ListaKrajow As Variant, WartosciKrajow As Variant
    Dim ZakresListy As Range, ZakresSortowania As Range, NazwyKrajow As Range
    
    If WybranyWskaznik = "Vaccinated" Then NrKol = 3
    
    Sheets("Dictionary").Activate
    Set ListaKrajowZakres = Sheets("Dictionary").Range(Cells(2, PozycjaKrajuSlownik), Cells(OstatniWiersz, PozycjaKrajuSlownik))
    
    ListaKrajow = ListaKrajowZakres.Value
    ReDim WartosciKrajow(1 To UBound(ListaKrajow))
    
    'Tworzymy listê wartoœci dla danego kraju do wykresu kolumnowego
    For i = LBound(ListaKrajow) To UBound(ListaKrajow)
    
        Dim KrajWartosc As Long
        KrajWartosc = Application.WorksheetFunction.VLookup(ListaKrajow(i, 1), Sheets(WybranyWskaznik).Range("A1:C1000"), NrKol, 0)
        WartosciKrajow(i) = KrajWartosc
        
    Next
    
    '''SORTOWANIE WYKRESU W ARKUSZU'''
    
    'Czyœcimy dane z poprzedniego wykresu
    Columns("AQ:AR").ClearContents
    
    'Ustawiamy nag³ówki kolumn jako Kraj oraz Wartoœæ
    Cells(1, 43).Value = "Kraj"
    Cells(1, 44).Value = "Wartoœæ"
    
    'Wklejamy listê krajów oraz w pêtli listê wartoœci dla danego kraju
    Range(Cells(2, 43), Cells(UBound(ListaKrajow), 43)).Value = ListaKrajow
    
    For i = 2 To UBound(ListaKrajow)
        Cells(i, 44).Value = WartosciKrajow(i - 1)
    Next i
        
    'Ustawiamy zakres zmiennych oraz kolumnê, po której bêdziemy sortowaæ - czyli kolumnê wartoœci
    Set ZakresSortowania = Range(Cells(1, 44), Cells(UBound(ListaKrajow), 44))
    Set ZakresListy = Range(Cells(1, 43), Cells(UBound(ListaKrajow), 44))
    
    ZakresListy.Select
    ActiveWorkbook.Worksheets("Dictionary").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dictionary").Sort.SortFields.Add2 Key:=ZakresSortowania, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Dictionary").Sort
        .SetRange ZakresListy
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set NazwyKrajow = Range(Cells(1, 43), Cells(UBound(ListaKrajow), 43)).Value = ListaKrajow
    Set ZakresSortowania = Range(Cells(1, 43), Cells(UBound(ListaKrajow), 43)).Value = WartosciKrajow
    
    Sheets("COUNTRY").Activate
    ActiveSheet.ChartObjects("EkranKraj_slupkowy").Activate
    ActiveChart.ChartTitle.Text = "Countries in " & Kontynent
    ActiveChart.SetSourceData Source:=ZakresListy
    
    Application.ScreenUpdating = True

End Sub

Sub WykresyKrajePL()
'makro uzupe³niaj¹ce wykres liniowy oraz kolumnowy - wersja polska

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''MAKRO ANALOGICZNE DO 'WykresyKrajeANG' TYLKO DLA KARTY W POLSKIEJ WERSJI''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error Resume Next
    ActiveSheet.Shapes.Range(Array("OkienkoError")).Delete
    
    Application.ScreenUpdating = False
    
    Sheets("Dictionary").Activate
    WybranyKraj = Application.WorksheetFunction.VLookup(Sheets("KRAJ").Range("B6").Value, Sheets("Dictionary").Range("R1:S1000"), 2, 0)
    
    Sheets("KRAJ").Activate
    If Sheets("KRAJ").CB_Wskaznik.Value = "" Then Exit Sub
    WybranyWskaznik = "H_" & LCase(Application.WorksheetFunction.VLookup(Sheets("KRAJ").CB_Wskaznik.Value, Sheets("Dictionary").Range("W1:X5"), 2, 0))
    
    If WybranyWskaznik = "H_vaccinated" Then
        WybranyWskaznik = "Vaccinated"
    End If
    
    Sheets("Dictionary").Activate
    Kontynent = Application.WorksheetFunction.VLookup(WybranyKraj, Sheets("Dictionary").Range("A2:B500"), 2, 0)
    PozycjaKrajuSlownik = Application.WorksheetFunction.Match(Kontynent, Sheets("Dictionary").Range("A1:N1"), 0)
    PozycjaKraju2 = Application.WorksheetFunction.Match(WybranyKraj, Sheets(WybranyWskaznik).Range("A1:A1000"), 0)
    
    OstatniWiersz = Sheets("Dictionary").Cells(1, PozycjaKrajuSlownik).End(xlDown).Row
    OstatniaKolumna = Sheets(WybranyWskaznik).Range("A1").End(xlToRight).Column
    
    Sheets(WybranyWskaznik).Activate
    Set DatyBazaZakres = Range(Cells(1, 2), Cells(1, OstatniaKolumna))
    
    Dim DatyBaza As Variant
    DatyBaza = DatyBazaZakres.Value
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''WYKRES LINIOWY Z DANYMI HISTORYCZNYMI'''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    Sheets("Wykresy").Activate
    If WybranyWskaznik = "Vaccinated" Then
    
        Sheets("KRAJ").Activate
        On Error Resume Next
        ActiveSheet.Shapes.Range(Array("OkienkoError")).Delete
        
        ActiveSheet.Shapes.AddShape(msoShapeRectangle, 66.25, 266.25, 583.75, 311.25). _
            Select
        Selection.ShapeRange.Name = "OkienkoError"
        Selection.ShapeRange("OkienkoError").IncrementLeft 688.75
        Selection.ShapeRange("OkienkoError").IncrementTop 137.5
        Selection.ShapeRange("OkienkoError").ShapeStyle = msoShapeStylePreset8
        Selection.ShapeRange("OkienkoError").TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange("OkienkoError").TextFrame2.TextRange.Characters.Text = "No daily data found. Category 'Vaccinated' has only total number of cases"
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 7). _
            ParagraphFormat
            .FirstLineIndent = 0
            .Alignment = msoAlignCenter
        End With
        
        GoTo Dalej
    
    Else
        
        On Error Resume Next
        ActiveSheet.Shapes.Range(Array("OkienkoError")).Delete
    
    End If
    
    For i = 1 To OstatniaKolumna - 1
        
        Dim DatyWykresOSX As Date
        DatyWykresOSX = Format(CDate(Mid(DatyBaza(1, i), 17, 10)), "YYYY-MM-DD")
        DatyBaza(1, i) = DatyWykresOSX
                        
    Next
    
    Sheets("KRAJ").Activate
    ActiveSheet.ChartObjects("EkranKraj_liniowy").Activate
    ActiveChart.FullSeriesCollection(1).Values = "=" & WybranyWskaznik & "!" & Range(Cells(PozycjaKraju2, 2), Cells(PozycjaKraju2, OstatniaKolumna)).Address
    ActiveChart.FullSeriesCollection(1).XValues = DatyBaza
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
    ActiveChart.Axes(xlValue).TickLabelPosition = xlHigh
    Application.CommandBars("Format Object").Visible = False

Dalej:

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''WYKRES KOLUMNOWY: KRAJ NA TLE KONTYNENTU'''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim NrKol As Long: NrKol = 2
    Dim ListaKrajow As Variant, WartosciKrajow As Variant
    
    If WybranyWskaznik = "Vaccinated" Then NrKol = 3
    
    Sheets("Dictionary").Activate
    Set ListaKrajowZakres = Sheets("Dictionary").Range(Cells(2, PozycjaKrajuSlownik), Cells(OstatniWiersz, PozycjaKrajuSlownik))
    
    ListaKrajow = ListaKrajowZakres.Value
    ReDim WartosciKrajow(1 To UBound(ListaKrajow))
    
    For i = LBound(ListaKrajow) To UBound(ListaKrajow)
    
        Dim abcd As Long
        abcd = Application.WorksheetFunction.VLookup(ListaKrajow(i, 1), Sheets(WybranyWskaznik).Range("A1:C1000"), NrKol, 0)
        WartosciKrajow(i) = abcd
        
    Next
    
    Dim ZakresListy As Range, ZakresSortowania As Range, NazwyKrajow As Range
    
    Columns("43:44").ClearContents
    
    Cells(1, 43).Value = "Kraj"
    Cells(1, 44).Value = "Wartoœæ"

    Range(Cells(2, 43), Cells(UBound(ListaKrajow), 43)).Value = ListaKrajow
    
    For i = 2 To UBound(ListaKrajow)
        Cells(i, 44).Value = WartosciKrajow(i - 1)
    Next i
    
    '''SORTOWANIE WYKRESU W ARKUSZU'''
        
    Set ZakresSortowania = Range(Cells(1, 44), Cells(UBound(ListaKrajow), 44))
    Set ZakresListy = Range(Cells(1, 43), Cells(UBound(ListaKrajow), 44))
    
    ZakresListy.Select
    ActiveWorkbook.Worksheets("Dictionary").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dictionary").Sort.SortFields.Add2 Key:=ZakresSortowania, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets("Dictionary").Sort
        .SetRange ZakresListy
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set NazwyKrajow = Range(Cells(1, 43), Cells(UBound(ListaKrajow), 43)).Value = ListaKrajow
    Set ZakresSortowania = Range(Cells(1, 43), Cells(UBound(ListaKrajow), 43)).Value = WartosciKrajow
    
    Dim NazwaWykresu As String
    
    If Kontynent = "Africa" Then
        NazwaWykresu = "Kraje w Afryce"
    ElseIf Kontynent = "Asia" Then
        NazwaWykresu = "Kraje w Azji"
    ElseIf Kontynent = "Europe" Then
        NazwaWykresu = "Kraje w Europie"
    ElseIf Kontynent = "North America" Then
        NazwaWykresu = "Kraje w Ameryce Pó³nocnej"
    ElseIf Kontynent = "Europe" Then
        NazwaWykresu = "Kraje w Australi i Oceanii"
    Else
        NazwaWykresu = "Kraje w Ameryce Po³udniowej"
    End If
    
    Sheets("KRAJ").Activate
    ActiveSheet.ChartObjects("EkranKraj_slupkowy").Activate
    ActiveChart.ChartTitle.Text = NazwaWykresu
    ActiveChart.SetSourceData Source:=ZakresListy
    
    Application.ScreenUpdating = True

End Sub


