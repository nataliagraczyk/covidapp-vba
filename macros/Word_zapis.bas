Attribute VB_Name = "Word_zapis"
Public Kraj_pl As String
Public Plik As String

Sub Eksport_word_Pl()
Application.ScreenUpdating = False
Dim Data As String
Data = Left(ThisWorkbook.Worksheets("Przypadki").Range("M2"), 10)

Dim Wykres As Chart

Kraj_pl = ThisWorkbook.Worksheets("Kraj").Range("B6")
Kraj = Application.WorksheetFunction.VLookup(Kraj_pl, ThisWorkbook.Worksheets("Dictionary").Range("R1").CurrentRegion, 3)

Call Licz_ogolne
Call Licz_kraj_ogolne
Call WykresyKraje_raport


'Otwiera plik Worda z szablonem, wstawia w wyznaczonym miejscu skopiowany z Excela tekst,
'zapisuje plik na dysku i zamyka go.
'Metoda: zaznaczenie zdefiniowanej nazwy i wklejenie w jej miejsce tekstu
'Deklaracja zmiennych

    Dim WordA As Word.Application 'aplikacja Word
    Dim SciezkaSzablon As String 'œcie¿ka do pliku Worda z szablonem
    Dim SciezkaZapis As String 'œcie¿ka do zapisywanego pliku Worda
    
    'Test
    
'Wczytanie zmiennych
    SciezkaSzablon = ActiveWorkbook.Path & "\Szablony\Word_templatka_pl.docx"
    SciezkaZapis = ActiveWorkbook.Path & "\Raport_Covid19_" & Kraj_pl & Format(Now(), "yyyymmddhhss") & ".docx"

'Stworzenie nowego dokumentu w Wordzie
    'Wczytanie nowego obiektu bêd¹cego aplikacj¹ Worda
    Set WordA = New Word.Application
    
    With WordA
        .Visible = False 'widoczna aplikacja Word
    
    'Otwarcie szablonu (pliku Word)
        .Documents.Open (SciezkaSzablon)
    
    'Uzupe³nienie raportu
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_tytul"
        .Selection = Kraj_pl  'wklejenie
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_1_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_ogólna"
        .Selection = Ilosc_Przypadkow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_ogolne"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_nowych_ogolne"
        .Selection = Przypadki_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_2_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_nowe_ogolne"
        .Selection = Miejsce_przypadki_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_ogolne"
        .Selection = Ilosc_Zgonow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_3_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_ogolne"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_ogolne"
        .Selection = Zgony_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_4_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_nowe"
        .Selection = Miejsce_zgony_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_ogolne"
        .Selection = Ilosc_Wyzdrowien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_5_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_ogolne"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_ogolne"
        .Selection = Wyzdrowienia_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_6_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_nowe"
        .Selection = Miejsce_wyzdrowienia_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_ogolne"
        .Selection = Ilosc_szczepien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_ogolne"
        .Selection = Szczepienia_pelne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_ogolne"
        .Selection = Szczepienia_1
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_7_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_ogolne"
        .Selection = Miejsce_szczepienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_przypadki_ogolne"
        .Selection = Format(W_Nowych, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_ogolne"
        .Selection = Format(W_Zgonow, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_ogolne"
        .Selection = Format(W_Wyzdrowien, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_przypadki_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_kraj"
        .Selection = Ilosc_przypadkow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_przypadki_kraj"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_przypadki_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_nowe_kraj"
        .Selection = Przypadki_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_zgony_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_kraj"
        .Selection = Ilosc_Zgonow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_kraj"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_kraj"
        .Selection = Zgony_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_kraj"
        .Selection = Format(W_Zgonow_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_wyzdrowienia_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_kraj"
        .Selection = Ilosc_Wyzdrowien_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_kraj"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_kraj"
        .Selection = Wyzdrowienia_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_kraj"
        .Selection = Format(W_Wyzdrowien_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_szczepienia_1"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_kraj"
        .Selection = Zaszczepieni_Calosc
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_kraj"
        .Selection = Zaszczepieni_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_kraj"
        .Selection = Zaszczepieni_1_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_kraj"
        .Selection = Miejsce_szczepienia
         
        'Wykresy
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_vaccinated").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_szczepienia"
       .Selection.Delete
       .Selection.Paste
       
    'Zapisanie i zamkniêcie pliku
        .ActiveDocument.SaveAs Filename:=SciezkaZapis 'zapisanie pliku
        .ActiveDocument.Close 'zamkniêcie pliku
        .Quit 'zamkniêcie aplikacji Word
    End With

    ThisWorkbook.Worksheets("Kraj").Select
    Application.ScreenUpdating = True
End Sub

Sub Eksport_word_en()
Application.ScreenUpdating = False
Dim Data As String
Data = Left(ThisWorkbook.Worksheets("Przypadki").Range("M2"), 10)

Dim Kraj_pl As String
Dim Wykres As Chart

Kraj_pl = ThisWorkbook.Worksheets("Country").Range("B6").Value
Kraj = Kraj_pl

Call Licz_ogolne
Call Licz_kraj_ogolne
Call WykresyKraje_raport


'Otwiera plik Worda z szablonem, wstawia w wyznaczonym miejscu skopiowany z Excela tekst,
'zapisuje plik na dysku i zamyka go.
'Metoda: zaznaczenie zdefiniowanej nazwy i wklejenie w jej miejsce tekstu
'Deklaracja zmiennych

    Dim WordA As Word.Application 'aplikacja Word
    Dim SciezkaSzablon As String 'œcie¿ka do pliku Worda z szablonem
    Dim SciezkaZapis As String 'œcie¿ka do zapisywanego pliku Worda
    
    'Test
    
'Wczytanie zmiennych
    SciezkaSzablon = ActiveWorkbook.Path & "\Szablony\Word_templatka_ang.docx"
    SciezkaZapis = ActiveWorkbook.Path & "\Report_Covid19_" & Kraj & Format(Now(), "yyyymmddhhss") & ".docx"

'Stworzenie nowego dokumentu w Wordzie
    'Wczytanie nowego obiektu bêd¹cego aplikacj¹ Worda
    Set WordA = New Word.Application
    
    With WordA
        .Visible = False 'widoczna aplikacja Word
    
    'Otwarcie szablonu (pliku Word)
        .Documents.Open (SciezkaSzablon)
    
    'Uzupe³nienie raportu
             .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_tytul"
        .Selection = Kraj_pl  'wklejenie
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_1_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_ogólna"
        .Selection = Ilosc_Przypadkow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_ogolne"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_nowych_ogolne"
        .Selection = Przypadki_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_2_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_nowe_ogolne"
        .Selection = Miejsce_przypadki_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_ogolne"
        .Selection = Ilosc_Zgonow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_3_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_ogolne"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_ogolne"
        .Selection = Zgony_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_4_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_nowe"
        .Selection = Miejsce_zgony_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_ogolne"
        .Selection = Ilosc_Wyzdrowien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_5_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_ogolne"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_ogolne"
        .Selection = Wyzdrowienia_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_6_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_nowe"
        .Selection = Miejsce_wyzdrowienia_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_ogolne"
        .Selection = Ilosc_szczepien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_ogolne"
        .Selection = Szczepienia_pelne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_ogolne"
        .Selection = Szczepienia_1
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_7_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_ogolne"
        .Selection = Miejsce_szczepienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_przypadki_ogolne"
        .Selection = Format(W_Nowych, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_ogolne"
        .Selection = Format(W_Zgonow, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_ogolne"
        .Selection = Format(W_Wyzdrowien, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_przypadki_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_kraj"
        .Selection = Ilosc_przypadkow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_przypadki_kraj"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_przypadki_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_nowe_kraj"
        .Selection = Przypadki_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_zgony_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_kraj"
        .Selection = Ilosc_Zgonow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_kraj"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_kraj"
        .Selection = Zgony_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_kraj"
        .Selection = Format(W_Zgonow_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_wyzdrowienia_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_kraj"
        .Selection = Ilosc_Wyzdrowien_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_kraj"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_kraj"
        .Selection = Wyzdrowienia_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_kraj"
        .Selection = Format(W_Wyzdrowien_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_szczepienia_1"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_kraj"
        .Selection = Zaszczepieni_Calosc
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_kraj"
        .Selection = Zaszczepieni_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_kraj"
        .Selection = Zaszczepieni_1_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_kraj"
        .Selection = Miejsce_szczepienia
        
        'Wykresy
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_vaccinated").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_szczepienia"
       .Selection.Delete
       .Selection.Paste
        
    'Zapisanie i zamkniêcie pliku
        .ActiveDocument.SaveAs Filename:=SciezkaZapis 'zapisanie pliku
        .ActiveDocument.Close 'zamkniêcie pliku
        .Quit 'zamkniêcie aplikacji Word
    End With
    ThisWorkbook.Worksheets("Country").Select
    Application.ScreenUpdating = True
End Sub


Sub Eksport_pdf_pl()

Application.ScreenUpdating = False
Dim Data As String
Data = Left(ThisWorkbook.Worksheets("Przypadki").Range("M2"), 10)

Dim Wykres As Chart

Kraj_pl = ThisWorkbook.Worksheets("Kraj").Range("B6")
Kraj = Application.WorksheetFunction.VLookup(Kraj_pl, ThisWorkbook.Worksheets("Dictionary").Range("R1").CurrentRegion, 3)

Call Licz_ogolne
Call Licz_kraj_ogolne
Call WykresyKraje_raport


'Otwiera plik Worda z szablonem, wstawia w wyznaczonym miejscu skopiowany z Excela tekst,
'zapisuje plik na dysku i zamyka go.
'Metoda: zaznaczenie zdefiniowanej nazwy i wklejenie w jej miejsce tekstu
'Deklaracja zmiennych

    Dim WordA As Word.Application 'aplikacja Word
    Dim SciezkaSzablon As String 'œcie¿ka do pliku Worda z szablonem
    Dim SciezkaZapis As String 'œcie¿ka do zapisywanego pliku Worda
    
    'Test
    
'Wczytanie zmiennych
    SciezkaSzablon = ActiveWorkbook.Path & "\Szablony\Word_templatka_pl.docx"
    SciezkaZapis = ActiveWorkbook.Path & "\Raport_Covid19_" & Kraj_pl & Format(Now(), "yyyymmddhhss") & ".pdf"
    Plik = SciezkaZapis
'Stworzenie nowego dokumentu w Wordzie
    'Wczytanie nowego obiektu bêd¹cego aplikacj¹ Worda
    Set WordA = New Word.Application
    
    With WordA
        .Visible = False 'widoczna aplikacja Word
    
    'Otwarcie szablonu (pliku Word)
        .Documents.Open (SciezkaSzablon)
    
    'Uzupe³nienie raportu
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_tytul"
        .Selection = Kraj_pl  'wklejenie
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_1_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_ogólna"
        .Selection = Ilosc_Przypadkow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_ogolne"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_nowych_ogolne"
        .Selection = Przypadki_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_2_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_nowe_ogolne"
        .Selection = Miejsce_przypadki_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_ogolne"
        .Selection = Ilosc_Zgonow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_3_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_ogolne"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_ogolne"
        .Selection = Zgony_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_4_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_nowe"
        .Selection = Miejsce_zgony_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_ogolne"
        .Selection = Ilosc_Wyzdrowien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_5_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_ogolne"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_ogolne"
        .Selection = Wyzdrowienia_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_6_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_nowe"
        .Selection = Miejsce_wyzdrowienia_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_ogolne"
        .Selection = Ilosc_szczepien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_ogolne"
        .Selection = Szczepienia_pelne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_ogolne"
        .Selection = Szczepienia_1
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_7_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_ogolne"
        .Selection = Miejsce_szczepienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_przypadki_ogolne"
        .Selection = Format(W_Nowych, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_ogolne"
        .Selection = Format(W_Zgonow, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_ogolne"
        .Selection = Format(W_Wyzdrowien, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_przypadki_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_kraj"
        .Selection = Ilosc_przypadkow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_przypadki_kraj"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_przypadki_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_nowe_kraj"
        .Selection = Przypadki_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_zgony_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_kraj"
        .Selection = Ilosc_Zgonow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_kraj"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_kraj"
        .Selection = Zgony_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_kraj"
        .Selection = Format(W_Zgonow_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_wyzdrowienia_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_kraj"
        .Selection = Ilosc_Wyzdrowien_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_kraj"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_kraj"
        .Selection = Wyzdrowienia_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_kraj"
        .Selection = Format(W_Wyzdrowien_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_szczepienia_1"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_kraj"
        .Selection = Zaszczepieni_Calosc
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_kraj"
        .Selection = Zaszczepieni_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_kraj"
        .Selection = Zaszczepieni_1_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_kraj"
        .Selection = Miejsce_szczepienia
         
        'Wykresy
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_vaccinated").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_szczepienia"
       .Selection.Delete
       .Selection.Paste
       
    

'Zapisanie do pdf
        .ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            SciezkaZapis, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        
    .ActiveDocument.Close savechanges:=False 'zamkniêcie pliku
    .Quit 'zamkniêcie aplikacji Word
    End With
    
    ThisWorkbook.Worksheets("Kraj").Select
    Application.ScreenUpdating = True
End Sub

Sub Eksport_pdf_en()
Application.ScreenUpdating = False
Dim Data As String
Data = Left(ThisWorkbook.Worksheets("Przypadki").Range("M2"), 10)

Dim Kraj_pl As String
Dim Wykres As Chart

Kraj_pl = ThisWorkbook.Worksheets("Country").Range("B6").Value
Kraj = Application.WorksheetFunction.VLookup(Kraj_pl, ThisWorkbook.Worksheets("Dictionary").Range("R1").CurrentRegion, 3)

Call Licz_ogolne
Call Licz_kraj_ogolne
Call WykresyKraje_raport


'Otwiera plik Worda z szablonem, wstawia w wyznaczonym miejscu skopiowany z Excela tekst,
'zapisuje plik na dysku i zamyka go.
'Metoda: zaznaczenie zdefiniowanej nazwy i wklejenie w jej miejsce tekstu
'Deklaracja zmiennych

    Dim WordA As Word.Application 'aplikacja Word
    Dim SciezkaSzablon As String 'œcie¿ka do pliku Worda z szablonem
    Dim SciezkaZapis As String 'œcie¿ka do zapisywanego pliku Worda
    
    'Test
    
'Wczytanie zmiennych
    SciezkaSzablon = ActiveWorkbook.Path & "\Szablony\Word_templatka_ang.docx"
    SciezkaZapis = ActiveWorkbook.Path & "\Report_Covid19_" & Kraj & Format(Now(), "yyyymmddhhss") & ".pdf"
    Plik = SciezkaZapis
'Stworzenie nowego dokumentu w Wordzie
    'Wczytanie nowego obiektu bêd¹cego aplikacj¹ Worda
    Set WordA = New Word.Application
    
    With WordA
        .Visible = False 'widoczna aplikacja Word
    
    'Otwarcie szablonu (pliku Word)
        .Documents.Open (SciezkaSzablon)
    
    'Uzupe³nienie raportu
             .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_tytul"
        .Selection = Kraj_pl  'wklejenie
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_1_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_ogólna"
        .Selection = Ilosc_Przypadkow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_ogolne"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Liczba_nowych_ogolne"
        .Selection = Przypadki_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_2_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_nowe_ogolne"
        .Selection = Miejsce_przypadki_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_ogolne"
        .Selection = Ilosc_Zgonow
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_3_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_ogolne"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_ogolne"
        .Selection = Zgony_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_4_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_nowe"
        .Selection = Miejsce_zgony_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_ogolne"
        .Selection = Ilosc_Wyzdrowien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_5_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_ogolne"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_ogolne"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_ogolne"
        .Selection = Wyzdrowienia_Nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_6_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_nowe"
        .Selection = Miejsce_wyzdrowienia_nowe
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_ogolne"
        .Selection = Ilosc_szczepien
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_ogolne"
        .Selection = Szczepienia_pelne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_ogolne"
        .Selection = Szczepienia_1
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_7_ogolne"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_ogolne"
        .Selection = Miejsce_szczepienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_przypadki_ogolne"
        .Selection = Format(W_Nowych, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_ogolne"
        .Selection = Format(W_Zgonow, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_ogolne"
        .Selection = Format(W_Wyzdrowien, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_przypadki_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_kraj"
        .Selection = Ilosc_przypadkow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_przypadki_kraj"
        .Selection = Miejsce_ogolne
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_przypadki_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Przypadki_nowe_kraj"
        .Selection = Przypadki_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_zgony_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_kraj"
        .Selection = Ilosc_Zgonow_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_zgony_kraj"
        .Selection = Miejsce_zgony
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_zgony_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Zgony_nowe_kraj"
        .Selection = Zgony_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_zgony_kraj"
        .Selection = Format(W_Zgonow_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_wyzdrowienia_1"
        .Selection = Kraj_pl

        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_kraj"
        .Selection = Ilosc_Wyzdrowien_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_wyzdrowienia_kraj"
        .Selection = Miejsce_wyzdrowienia
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Data_wyzdrowienia_kraj"
        .Selection = Data
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Wyzdrowienia_nowe_kraj"
        .Selection = Wyzdrowienia_nowe_k
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="W_wyzdrowienia_kraj"
        .Selection = Format(W_Wyzdrowien_K, "Percent")
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Kraj_szczepienia_1"
        .Selection = Kraj_pl
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_calosc_kraj"
        .Selection = Zaszczepieni_Calosc
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_w_pelni_kraj"
        .Selection = Zaszczepieni_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Szczepienia_czesciowe_kraj"
        .Selection = Zaszczepieni_1_K
        
        .Selection.GoTo what:=wdGoToBookmark, Name:="Miejsce_szczepienia_kraj"
        .Selection = Miejsce_szczepienia
        
        'Wykresy
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_confirmed").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_przypadki_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_deaths").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_smierci_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_liniowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_liniowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_recovered").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_wyzdrowienia_kolumnowy"
       .Selection.Delete
       .Selection.Paste
       
       Set Wykres = Sheets("wykresy").ChartObjects("EkranKraj_slupkowy_vaccinated").Chart
       Wykres.CopyPicture
       .Selection.GoTo what:=wdGoToBookmark, Name:="Wykres_szczepienia"
       .Selection.Delete
       .Selection.Paste
       
       'Zapisanie do pdf
        .ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            SciezkaZapis, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        
    .ActiveDocument.Close savechanges:=False 'zamkniêcie pliku
    .Quit 'zamkniêcie aplikacji Word
    End With
    
    ThisWorkbook.Worksheets("Country").Select
    Application.ScreenUpdating = True
End Sub
