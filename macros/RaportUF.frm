VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RaportUF 
   Caption         =   "COVID-19 App"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8160
   OleObjectBlob   =   "RaportUF.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "RaportUF"
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


Private Sub CB_Next_Click()

Dim ash1 As Worksheet

Set ash1 = ActiveSheet
    Sheets("COUNTRY").Unprotect
    Sheets("KRAJ").Unprotect
    
    Sheets("COUNTRY").Visible = True
    Sheets("COUNTRY").Range("B6").Value = RaportUF.CBX_Kraj.Value
    Call WykresyKrajeANG
        
    Sheets("KRAJ").Visible = True
    Sheets("KRAJ").Range("B6").Value = RaportUF.CBX_Kraj.Value
    Call WykresyKrajePL

    
    Call Metryczka1
    Unload RaportUF
    
    If ash1.Name = "REPORT" Then
        Sheets("COUNTRY").Activate
        Sheets("KRAJ").Visible = False
    Else
        Sheets("KRAJ").Activate
        Sheets("COUNTRY").Visible = False
    End If
    
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False

    ActiveSheet.Range("A1:AI48").Select
    ActiveWindow.Zoom = True

    Range("A1").Select
    
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

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = xlPrimaryButton And Shift = 1 Then
    Call ReleaseCapture
    Call SendMessage(hWndForm, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
  End If
End Sub


Private Sub CB_Exit_Click()

    Unload RaportUF
    
    
End Sub

Private Sub CBX_Kontynent_Change()

    Application.ScreenUpdating = False
    
    Dim ash As Worksheet
    Set ash = ActiveSheet
    
    Dim PozycjaKrajuSlownik As Long, OstatniWiersz As Long
    Dim ListaKrajow As Range

    If RaportUF.CBX_Kontynent = "" Then
        RaportUF.CBX_Kraj.List = Sheets("Dictionary").Range("A2:A193").Value
    Else
        Sheets("Dictionary").Activate
        PozycjaKrajuSlownik = Application.WorksheetFunction.Match(RaportUF.CBX_Kontynent.Value, Sheets("Dictionary").Range("A1:N1"), 0)
        OstatniWiersz = Sheets("Dictionary").Cells(1, PozycjaKrajuSlownik).End(xlDown).Row

        Set ListaKrajow = Sheets("Dictionary").Range(Cells(2, PozycjaKrajuSlownik), Cells(OstatniWiersz, PozycjaKrajuSlownik))
        
        RaportUF.CBX_Kraj.Value = ""
        RaportUF.CBX_Kraj.List = ListaKrajow.Value

    End If
    
    ash.Activate
    Application.ScreenUpdating = True

End Sub

Private Sub CBX_Kraj_Change()

    If RaportUF.CBX_Kraj = "" Then
        RaportUF.CB_Next.Enabled = False
    Else
        RaportUF.CB_Next.Enabled = True
    End If

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Initialize()


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
        RaportUF.LabelKontynent.Caption = "Wybierz kontynent"
        RaportUF.LabelKraj.Caption = "Wybierz kraj*"
        RaportUF.CB_Next.Caption = "Dalej"
        RaportUF.CB_Exit.Caption = "WyjdŸ"
        RaportUF.LabelRequired.Caption = "Wymagane*"
    ElseIf ThisSheet.Name = "RAPORT" Then
        RaportUF.LabelKontynent.Caption = "Wybierz kontynent"
        RaportUF.LabelKraj.Caption = "Wybierz kraj*"
        RaportUF.CB_Next.Caption = "Dalej"
        RaportUF.CB_Exit.Caption = "WyjdŸ"
        RaportUF.LabelRequired.Caption = "Wymagane*"
    ElseIf ThisSheet.Name = "COUNTRY" Then
        RaportUF.LabelKontynent.Caption = "Choose continent"
        RaportUF.LabelKraj.Caption = "Choose country*"
        RaportUF.CB_Next.Caption = "Next"
        RaportUF.CB_Exit.Caption = "Exit"
        RaportUF.LabelRequired.Caption = "Required*"
    ElseIf ThisSheet.Name = "REPORT" Then
        RaportUF.LabelKontynent.Caption = "Choose continent"
        RaportUF.LabelKraj.Caption = "Choose country*"
        RaportUF.CB_Next.Caption = "Next"
        RaportUF.CB_Exit.Caption = "Exit"
        RaportUF.LabelRequired.Caption = "Required*"
    End If

    RaportUF.CBX_Kontynent.List = Array("Asia", "Africa", "Europe", "North America", "Oceania", "South America")
    RaportUF.CBX_Kraj.List = Sheets("Dictionary").Range("A2:A193").Value
    RaportUF.CB_Next.Enabled = False
    
End Sub
