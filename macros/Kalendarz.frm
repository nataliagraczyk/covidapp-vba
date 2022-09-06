VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Kalendarz 
   Caption         =   "Kalendarz"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7050
   OleObjectBlob   =   "Kalendarz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Kalendarz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''' 0. AKTYWACJA USERFORMA '''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserForm_Activate()
'Makro uruchamiane po aktywacji UserForma

'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Stwórz listê miesiêcy
  For i = 1 To 12
    Me.LB_Miesiac.AddItem MonthName(i)
  Next i

  'Stwórz listê lat
  For i = 2020 To Year(Now)
    Me.LB_Rok.AddItem i * 1
  Next i

  Call Aktualizuj_listy

  'Zmieñ kolory przycisków
  For i = 1 To 42
    Me.Controls("TB" & i - 1).Value = False
    Me.Controls("TB" & i - 1).ForeColor = RGB(0, 0, 0)
  Next i

  'Szare weekendy
  On Error Resume Next
    For i = 5 To 40 Step 7
      Me.Controls("TB" & i).BackColor = RGB(220, 220, 220)
      Me.Controls("TB" & i + 1).BackColor = RGB(220, 220, 220)
    Next i
  On Error GoTo koniec

  'Wczytanie aktualnej daty
    LB_Miesiac.ListIndex = Month(Date) - 1 'miesi¹c
    LB_Rok.ListIndex = Year(Date) - 2020 'rok

  'Aktualizuj przyciski kalendarza
  Call Aktualizuj_przyciski

  Exit Sub

koniec:
  Unload Me

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''' 1. AKTUALIZACJA PRZYCISKÓW '''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Aktualizuj_przyciski()
'Makro aktualizuj¹ce przyciski kalendarza reprezentuj¹ce kolejne dni

'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Poka¿ przyciski od 36 do 42 - ostatni wiersz
  For i = 36 To 42
    Me.Controls("TB" & i - 1).Visible = True
  Next i

  'Wczytaj miesi¹c
  For i = 1 To 12
    If i = LB_Miesiac.ListIndex + 1 Then miesiac = i
  Next i

  'Wczytaj rok
  rok = Me.LB_Rok.List(Me.LB_Rok.ListIndex)

  'Zmienne pomocnicze
  Pierwszy_dzien = Weekday(DateSerial(rok, miesiac, 1), vbMonday)
  Ostatni_dzien = Day(DateSerial(rok, miesiac + 1, 1) - 1)
  Ostatni_dzien_pop = Day(DateSerial(rok, miesiac, 1) - 1)

  'Poprzedni miesi¹c
  For i = Pierwszy_dzien - 1 To 1 Step -1
    Me.Controls("TB" & i - 1).Caption = Ostatni_dzien_pop - Pierwszy_dzien + 1 + i
    Me.Controls("TB" & i - 1).ForeColor = RGB(120, 120, 120)
    'Me.Controls("TB" & i - 1).Enabled = False
  Next i

  'Odpowiedni miesi¹c
  For i = Pierwszy_dzien To Ostatni_dzien + Pierwszy_dzien - 1
    Me.Controls("TB" & i - 1).Caption = i - Pierwszy_dzien + 1
    Me.Controls("TB" & i - 1).ForeColor = RGB(0, 0, 0)
    Me.Controls("TB" & i - 1).Enabled = True
  Next i

  'Jeœli nie ma potrzeby pokazywania ostatniego rzêdu to ukryj
  If i <= 36 Then
    For i = 36 To 42
      Me.Controls("TB" & i - 1).Visible = False
    Next i
  End If

  'Nastêpny miesi¹c
  For i = Ostatni_dzien + Pierwszy_dzien To 42
    Me.Controls("TB" & i - 1).Caption = i - Ostatni_dzien - Pierwszy_dzien + 1
    Me.Controls("TB" & i - 1).ForeColor = RGB(120, 120, 120)
    'Me.Controls("TB" & i - 1).Enabled = False
  Next i

  Exit Sub

koniec:
  Unload Me

End Sub

Private Sub Aktualizuj_listy()
'Makro aktualizuj¹ce listy miesiêcy i lat

'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Aktualizacja zaznaczenia list
  If Me.LB_Rok.ListIndex = -1 Then Me.LB_Rok.ListIndex = 0
  If Me.LB_Miesiac.ListIndex = -1 Then Me.LB_Miesiac.ListIndex = 0

  Exit Sub

koniec:
  Unload Me

End Sub

Private Sub LB_Miesiac_Click()
  'Po klikniêciu na listê z miesi¹cami aktualizuj wszystko

'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  Call Aktualizuj_listy
  Call Aktualizuj_przyciski

  Exit Sub

koniec:
  Unload Me

End Sub

Private Sub LB_Rok_Click()
  'Po klikniêciu na listê z latami aktualizuj wszystko

  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  Call Aktualizuj_listy
  Call Aktualizuj_przyciski

  Exit Sub

koniec:
  Unload Me

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''' 2. WYBÓR DATY - KLIKANIE NA PRZYCISKI'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub TB0_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB0.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB1_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB1.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB2_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB2.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB3_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB3.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB4_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB4.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB5_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB5.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB6_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB6.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB7_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB7.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB8_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB8.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB9_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB9.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB10_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB10.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB11_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB11.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB12_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB12.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB13_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB13.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB14_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB14.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB15_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB15.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB16_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB16.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB17_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB17.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB18_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB18.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB19_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB19.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB20_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB20.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB21_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB21.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB22_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB22.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB23_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB23.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB24_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB24.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB25_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB25.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB26_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB26.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB27_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB27.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB28_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB28.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB29_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB29.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB30_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB30.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB31_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB31.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB32_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB32.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB33_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB33.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB34_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB34.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB35_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB35.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB36_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB36.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB37_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB37.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB38_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB38.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB39_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB39.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB40_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB40.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub

Private Sub TB41_Click()
  'Jeœli b³¹d to idŸ do koñca
  On Error GoTo koniec

  'Wczytanie daty do zmiennej DataKalendarz
    dzien = TB41.Caption
    miesiac = LB_Miesiac.ListIndex + 1
    rok = LB_Rok.List(LB_Rok.ListIndex)
    DataKalendarz = DateSerial(rok, miesiac, dzien)

koniec:
    'Koniec procedury
    Unload Me
End Sub


