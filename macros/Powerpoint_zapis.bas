Attribute VB_Name = "Powerpoint_zapis"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Stworzenie prezentacji w PowerPoincie
'Obiekt: Microsoft PowerPoint
'https://msdn.microsoft.com/en-us/library/fp161225.aspx
'https://msdn.microsoft.com/en-us/library/office/ff744643(v=office.16).aspx
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Raport_PowerPoint_Pl()

Application.ScreenUpdating = False
'Deklaracja zmiennych
    Dim PowerPointA As PowerPoint.Application 'Aplikacja PowerPoint
    Dim PrezentacjaPP As PowerPoint.Presentation 'Prezentacja PowerPoint
    Dim SlajdPP As PowerPoint.Slide 'Slajd prezentacji PowerPoint
    Dim WykresPP As PowerPoint.Shape 'Wykres w formacie obrazu
    Dim Wykres As Excel.ChartObject 'Wykres w Excelu
    
    Dim Tekst As String 'tekst z opisem wskaŸnika
    Dim NazwaPP As String 'Nazwa pliku z szablonem prezentacji PowerPoint
    Dim Data As String
    Dim Kraj_pl As String
    
'Wczytanie zmiennych
    NazwaPP = ThisWorkbook.Path & "\Szablony\" & "Powerpoint_szablon_pl.pptx"
    Kraj_pl = ThisWorkbook.Worksheets("Kraj").Range("B6").Value
    Kraj = Application.WorksheetFunction.VLookup(Kraj_pl, ThisWorkbook.Worksheets("Dictionary").Range("R1").CurrentRegion, 3)
    
    Data = Left(ThisWorkbook.Worksheets("Przypadki").Range("M2"), 10)
    
    Call Licz_ogolne
    Call Licz_kraj_ogolne
    Call WykresyKraje_raport
    
    'Wczytanie i otwarcie szablonu prezentacji PowerPoint
   'Wczytanie nowego obiektu bêd¹cego aplikacj¹ PowerPoint
    Set PowerPointA = New PowerPoint.Application
    PowerPointA.Visible = False 'widoczna aplikacja PowerPoint
    'Otwarcie szablonu prezentacji PowerPoint
    Set PrezentacjaPP = PowerPointA.Presentations.Open(NazwaPP)
    
    'Stworzenie prezentacji
    With PrezentacjaPP
    'Uzupe³nienie slajdów wstêpnych
        'Slajd nr 1
        .Slides(1).Shapes(2).TextFrame.TextRange = Kraj_pl & vbNewLine & "Raport COVID-19" & vbNewLine & Data
    'Stworzenie slajdów ze wskaŸnikami
    'Slajd 2
            'Dodanie nowego slajdu (typ ppLayoutText: tytu³ oraz pole tekstowe)
            Set SlajdPP = .Slides.Add(.Slides.count + 1, ppLayoutText)
            With SlajdPP
            'Wstawienie tytu³u slajdu
                With .Shapes(1)
                  .TextFrame.TextRange = "Dane dla œwiata: "
                  .Top = 60
                  .Left = 40
                  .Height = 40
                  .TextEffect.FontSize = 44
                End With
                
                With .Shapes(2)
                  .TextFrame.TextRange = "Liczba wszystkich przypadków: " & Ilosc_Przypadkow & vbNewLine & "Liczba zgonów: " & Ilosc_Zgonow _
                                        & vbNewLine & "Liczba wyzdrowieñ: " & Ilosc_Wyzdrowien & vbNewLine & "Liczba szczepieñ: " & _
                                        Ilosc_szczepien & " (w tym " & Szczepienia_pelne & " zaszczepionych w pe³ni)"
                  .Top = 120
                  .Left = 40
                  .Height = 40
                  .TextEffect.FontSize = 20
                End With
                
                 'Stworzenie slajdów ze wskaŸnikami
        For i = 1 To 7
            'Dodanie nowego slajdu (typ ppLayoutText: tytu³ oraz pole tekstowe)
            Set SlajdPP = PrezentacjaPP.Slides.Add(PrezentacjaPP.Slides.count + 1, ppLayoutText)
            With SlajdPP
            'Wstawienie tytu³u slajdu
                With .Shapes(1)
                  .TextFrame.TextRange = Kraj_pl & vbNewLine & Sheets("wykresy").Cells(9 + i, 1)
                  .Top = 40
                  .Left = 40
                  .Height = 60
                  .TextEffect.FontSize = 32
                End With
            'Wstawienie wykresu
                'Wklejenie (w formacie Metafile Picture)
                If i = 4 Then Sheets("wykresy").ChartObjects(i + 2).Chart.ChartArea.Copy
                If i = 5 Then Sheets("wykresy").ChartObjects(i - 1).Chart.ChartArea.Copy
                If i <> 4 And i <> 5 Then
                    Sheets("wykresy").ChartObjects(i).Chart.ChartArea.Copy
                End If
                On Error Resume Next
Again:
                .Shapes.PasteSpecial DataType:=ppPasteMetafilePicture
                If Err <> 0 Then
                    Err = 0
                    GoTo Again
                End If
                On Error GoTo 0
                Set WykresPP = .Shapes(SlajdPP.Shapes.count)
                'Pozycjonowanie wykresu na slajdzie
                With WykresPP
                    .Left = 40
                    .Top = 100
                    .LockAspectRatio = msoFalse 'odblokowanie wsp. proporcji
                    .Width = 640
                    .Height = 300
                End With
            
                'Formatowanie pola
                With .Shapes(2)
                  .TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone 'usuniêcie punktora
                  'pozycjonowanie pola
                  .Top = 410
                  .Left = 40
                  .Width = 640
                  .Height = 120
                  'wstawienie tekstu
                  .TextFrame.TextRange = Kraj_pl & vbNewLine & Sheets("wykresy").Cells(18 + i, 1) & Kraj_lista(i)
                  'zmiana wielkoœci czcionki
                  .TextEffect.FontSize = 20
                End With
            End With
        Next i
      End With
    
    'Stworzenie slajdu koñcowego
        Set SlajdPP = PowerPointA.Presentations(1).Slides.Add _
            (PowerPointA.Presentations(1).Slides.count + 1, ppLayoutText)
        SlajdPP.Shapes(1).Delete
        With SlajdPP.Shapes(1)
            .TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone 'usuniêcie punktora
            .TextFrame.TextRange = "Koniec"
            .TextFrame.HorizontalAnchor = msoAnchorCenter
            .TextFrame.VerticalAnchor = msoAnchorMiddle
            .TextEffect.FontBold = msoCTrue
            .TextEffect.FontSize = 40
        End With
    End With
'Zapisanie prezentacji PowerPoint
  PrezentacjaPP.SaveAs ThisWorkbook.Path & "\Raport_Covid19_" & Kraj_pl & Format(Now(), "yyyymmddhhss") & ".pptx"
    
    PrezentacjaPP.Close
    PowerPointA.Quit
    
    Worksheets("Kraj").Activate
    Application.ScreenUpdating = True
End Sub

Sub Raport_PowerPoint_En()

Application.ScreenUpdating = False

'Deklaracja zmiennych
    Dim PowerPointA As PowerPoint.Application 'Aplikacja PowerPoint
    Dim PrezentacjaPP As PowerPoint.Presentation 'Prezentacja PowerPoint
    Dim SlajdPP As PowerPoint.Slide 'Slajd prezentacji PowerPoint
    Dim WykresPP As PowerPoint.Shape 'Wykres w formacie obrazu
    Dim Wykres As Excel.ChartObject 'Wykres w Excelu
    
    Dim Tekst As String 'tekst z opisem wskaŸnika
    Dim NazwaPP As String 'Nazwa pliku z szablonem prezentacji PowerPoint
    Dim Data As String
    Dim Kraj_pl As String
    
'Wczytanie zmiennych
    NazwaPP = ThisWorkbook.Path & "\Szablony\" & "Powerpoint_szablon_pl.pptx"
    Kraj = ThisWorkbook.Sheets("Country").Range("B6").Value
    
    Data = Left(ThisWorkbook.Worksheets("Przypadki").Range("M2"), 10)
    
    Call Licz_ogolne
    Call Licz_kraj_ogolne
    Call WykresyKraje_raport
    
    'Wczytanie i otwarcie szablonu prezentacji PowerPoint
   'Wczytanie nowego obiektu bêd¹cego aplikacj¹ PowerPoint
    Set PowerPointA = New PowerPoint.Application
    'PowerPointA.Visible = False 'widoczna aplikacja PowerPoint
    'Otwarcie szablonu prezentacji PowerPoint
    Set PrezentacjaPP = PowerPointA.Presentations.Open(NazwaPP)
    
    'Stworzenie prezentacji
    With PrezentacjaPP
    'Uzupe³nienie slajdów wstêpnych
        'Slajd nr 1
        .Slides(1).Shapes(2).TextFrame.TextRange = Kraj & vbNewLine & "COVID-19 Report" & vbNewLine & Data
    'Stworzenie slajdów ze wskaŸnikami
    'Slajd 2
            'Dodanie nowego slajdu (typ ppLayoutText: tytu³ oraz pole tekstowe)
            Set SlajdPP = .Slides.Add(.Slides.count + 1, ppLayoutText)
            With SlajdPP
            'Wstawienie tytu³u slajdu
                With .Shapes(1)
                  .TextFrame.TextRange = "World Data: "
                  .Top = 60
                  .Left = 40
                  .Height = 40
                  .TextEffect.FontSize = 44
                End With
                
                With .Shapes(2)
                  .TextFrame.TextRange = "Total cases: " & Ilosc_Przypadkow & vbNewLine & "Deaths: " & Ilosc_Zgonow _
                                        & vbNewLine & "Recovered: " & Ilosc_Wyzdrowien & vbNewLine & "Vaccinated: " & _
                                        Ilosc_szczepien & " (including " & Szczepienia_pelne & " fully vaccinated)"
                  .Top = 120
                  .Left = 40
                  .Height = 40
                  .TextEffect.FontSize = 20
                End With
                
                 'Stworzenie slajdów ze wskaŸnikami
        For i = 1 To 7
            'Dodanie nowego slajdu (typ ppLayoutText: tytu³ oraz pole tekstowe)
            Set SlajdPP = PrezentacjaPP.Slides.Add(PrezentacjaPP.Slides.count + 1, ppLayoutText)
            With SlajdPP
            'Wstawienie tytu³u slajdu
                With .Shapes(1)
                  .TextFrame.TextRange = Kraj & vbNewLine & Sheets("wykresy").Cells(9 + i, 3)
                  .Top = 40
                  .Left = 40
                  .Height = 60
                  .TextEffect.FontSize = 32
                End With
            'Wstawienie wykresu
                'Wklejenie (w formacie Metafile Picture)
                If i = 4 Then Sheets("wykresy").ChartObjects(i + 2).Chart.ChartArea.Copy
                If i = 5 Then Sheets("wykresy").ChartObjects(i - 1).Chart.ChartArea.Copy
                If i <> 4 And i <> 5 Then
                    Sheets("wykresy").ChartObjects(i).Chart.ChartArea.Copy
                End If
                On Error Resume Next
Again:
                .Shapes.PasteSpecial DataType:=ppPasteMetafilePicture
                If Err <> 0 Then
                    Err = 0
                    GoTo Again
                End If
                On Error GoTo 0
                Set WykresPP = .Shapes(SlajdPP.Shapes.count)
                'Pozycjonowanie wykresu na slajdzie
                With WykresPP
                    .Left = 40
                    .Top = 100
                    .LockAspectRatio = msoFalse 'odblokowanie wsp. proporcji
                    .Width = 640
                    .Height = 300
                End With
            
                'Formatowanie pola
                With .Shapes(2)
                  .TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone 'usuniêcie punktora
                  'pozycjonowanie pola
                  .Top = 410
                  .Left = 40
                  .Width = 640
                  .Height = 120
                  'wstawienie tekstu
                  .TextFrame.TextRange = Kraj & vbNewLine & Sheets("wykresy").Cells(18 + i, 3) & Kraj_lista(i)
                  'zmiana wielkoœci czcionki
                  .TextEffect.FontSize = 20
                End With
            End With
        Next i
      End With
    
    'Stworzenie slajdu koñcowego
        Set SlajdPP = PowerPointA.Presentations(1).Slides.Add _
            (PowerPointA.Presentations(1).Slides.count + 1, ppLayoutText)
        SlajdPP.Shapes(1).Delete
        With SlajdPP.Shapes(1)
            .TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone 'usuniêcie punktora
            .TextFrame.TextRange = "The End"
            .TextFrame.HorizontalAnchor = msoAnchorCenter
            .TextFrame.VerticalAnchor = msoAnchorMiddle
            .TextEffect.FontBold = msoCTrue
            .TextEffect.FontSize = 40
        End With
    End With
'Zapisanie prezentacji PowerPoint
  PrezentacjaPP.SaveAs ThisWorkbook.Path & "\Report_Covid19_" & Kraj & Format(Now(), "yyyymmddhhss") & ".pptx"
    
    PrezentacjaPP.Close
    PowerPointA.Quit
    
    Worksheets("Kraj").Activate
Application.ScreenUpdating = True
End Sub

