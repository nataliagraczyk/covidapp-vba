VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShowReportsUF 
   Caption         =   "COVID-19 App"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7520
   OleObjectBlob   =   "ShowReportsUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ShowReportsUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Poni¿szy kod, za wyj¹tkiem makr Image2_Click() oraz UserForm_Activate() pochodzi ze strony:
'http://www.excelfox.com/forum/showthread.php/539-Remove-UserForm-s-TitleBar-And-Frame
'autor:  Rick Rothstein
'Przygotowany przez niego kod umo¿liwia pozbycie siê z userforma szpetnego obramowania oraz górnego paska

'**** Start of API Calls To Remove The UserForm's Title Bar ****
Private Declare PtrSafe Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
  

Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
  

Private Declare PtrSafe Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
  

Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
               (ByVal hwnd As Long) As Long
'**** End of API Calls To Remove The UserForm's Title Bar ****

'**** Start of API Calls To Allow User To Slide UserForm Around The Screen ****
Private Declare PtrSafe Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
               (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
 

Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
 

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'**** End of API Calls To Allow User To Slide UserForm Around The Screen ****


Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = xlPrimaryButton And Shift = 1 Then
    Call ReleaseCapture
    Call SendMessage(hWndForm, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
  End If
End Sub

Private Sub CB_Open_Click()

Dim Katalog As String
Dim PowerPointApp As PowerPoint.Application
Dim WordApp As Word.Application

Katalog = ActiveWorkbook.Path & "\"

If ShowReportsUF.LB_Raporty.ListIndex < 0 Then Exit Sub

If Right(ShowReportsUF.LB_Raporty.Value, 4) = ".doc" Or Right(ShowReportsUF.LB_Raporty.Value, 4) = "docx" Then
   
    Set WordApp = CreateObject("Word.Application")
    WordApp.Documents.Open Katalog & ShowReportsUF.LB_Raporty.Value
    WordApp.Visible = True

    Exit Sub
    
ElseIf Right(ShowReportsUF.LB_Raporty.Value, 4) = "pptx" Then
    
    ActiveWorkbook.FollowHyperlink Ktalog & ShowReportsUF.LB_Raporty.Value
    
    Exit Sub

ElseIf Right(ShowReportsUF.LB_Raporty.Value, 4) = ".pdf" Then

     ActiveWorkbook.FollowHyperlink Ktalog & ShowReportsUF.LB_Raporty.Value
     
Else: Exit Sub

End If

End Sub

Private Sub CommandButtonWyjdz_Click()
    Unload Me
End Sub

Private Sub LB_Raporty_Change()

If ShowReportsUF.LB_Raporty.ListIndex < 0 Then
  ShowReportsUF.CB_Open.Enabled = False
Else
  ShowReportsUF.CB_Open.Enabled = True
End If




End Sub

Private Sub UserForm_Initialize()

Application.ScreenUpdating = False

    Dim hWndForm As Long
   Dim Style As Long, Menu As Long
   hWndForm = FindWindow("ThunderDFrame", Me.Caption)
   Style = GetWindowLong(hWndForm, &HFFF0)
   Style = Style And Not &HC00000
   SetWindowLong hWndForm, &HFFF0, Style
   DrawMenuBar hWndForm

Dim ThisSheet As Worksheet

Set ThisSheet = ActiveSheet
   
    If ThisSheet.Name = "KRAJ" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "Dostêpne raporty:"
        ShowReportsUF.CB_Open.Caption = "Otwórz"
        ShowReportsUF.CommandButtonWyjdz.Caption = "WyjdŸ"
    ElseIf ThisSheet.Name = "RAPORT" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "Dostêpne raporty:"
        ShowReportsUF.CB_Open.Caption = "Otwórz"
        ShowReportsUF.CommandButtonWyjdz.Caption = "WyjdŸ"
    ElseIf ThisSheet.Name = "COUNTRY" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "Reports available:"
        ShowReportsUF.CB_Open.Caption = "Open"
        ShowReportsUF.CommandButtonWyjdz.Caption = "Exit"
    ElseIf ThisSheet.Name = "REPORT" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "Reports available:"
        ShowReportsUF.CB_Open.Caption = "Open"
        ShowReportsUF.CommandButtonWyjdz.Caption = "Exit"
    End If


Dim i As Integer, count As Integer
Dim ash As Worksheet

Set ash = ActiveSheet

ShowReportsUF.LB_Raporty.ColumnCount = 2
ShowReportsUF.LB_Raporty.ColumnWidths = "200,75"

'Kod z zajêæ 9. z pierwszej czêœci Kursu
'Wypisanie wszystkich plików Excela ze wskazanego katalogu
Dim Katalog As String 'lokalizacja katalogu
Dim Plik As String 'nazwa kolejnego pliku w katalogu
Dim Wiersz As Long, Kolumna As Long 'miejsce wypisywania nazw plików
Dim ListaRozszerzen() As Variant

If ShowReportsUF.LB_Raporty.ListIndex < 0 Then
  ShowReportsUF.CB_Open.Enabled = False
Else
  ShowReportsUF.CB_Open.Enabled = True
End If


ListaRozszerzen = Array("*.doc*", "*.pptx", "*.pdf")

Wiersz = 1
Kolumna = 37

Sheets("Dictionary").Activate

Cells(Wiersz, Kolumna) = "Nazwa pliku"
Cells(Wiersz, Kolumna + 1) = "Data modyfikacji"
   
    
    'Wskazanie lokalizacji danych
    Katalog = ActiveWorkbook.Path & "\"
    
'    For i = 0 To 2
'
'    'Wypisanie kolejnych plików Exela z katalogu
'    Plik = Dir(Katalog & ListaRozszerzen(i)) 'pobranie pierwszego pliku Excela z katalogu
    
    count = 0

    For i = 0 To 2
    Plik = Dir(Katalog & ListaRozszerzen(i))

    Do While Plik <> "" 'powtarzanie a¿ do napotkania pustego pliku

        count = count + 1
        Wiersz = Wiersz + 1
        Cells(Wiersz, Kolumna) = Plik
        Cells(Wiersz, Kolumna + 1) = FileDateTime(Katalog & Plik)
        Plik = Dir() 'przejœcie do kolejnego pliku Excela
    
    Loop
Next

If count = 0 Then

    If ThisSheet.Name = "KRAJ" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "Brak raportów do wyœwietlenia."
    ElseIf ThisSheet.Name = "RAPORT" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "Brak raportów do wyœwietlenia."
    ElseIf ThisSheet.Name = "COUNTRY" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "There are no reports to display"
    ElseIf ThisSheet.Name = "REPORT" Then
        ShowReportsUF.LabelDostepneRaporty.Caption = "There are no reports to display"
    End If
    
    GoTo Zakoncz

End If

Dim ZakresListy As Range, ZakresSortowania As Range

Set ZakresListy = Sheets("Dictionary").Range(Cells(1, Kolumna), Cells(Wiersz, Kolumna + 1))
Set ZakresSortowania = Sheets("Dictionary").Range(Cells(1, Kolumna + 1), Cells(Wiersz, Kolumna + 1))

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


Wiersz = 1

For i = 1 To count
   ShowReportsUF.LB_Raporty.AddItem
   ShowReportsUF.LB_Raporty.List(i - 1, 0) = Sheets("Dictionary").Cells(Wiersz + i, Kolumna)
   ShowReportsUF.LB_Raporty.List(i - 1, 1) = Sheets("Dictionary").Cells(Wiersz + i, Kolumna + 1)
Next i


Zakoncz:

Sheets(ash.Index).Activate

Application.ScreenUpdating = True

End Sub
