VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExitUF 
   Caption         =   "COVID-19 App"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7080
   OleObjectBlob   =   "ExitUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExitUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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



Private Sub UserForm_Initialize()
    Dim hWndForm As Long
   Dim Style As Long, Menu As Long
   hWndForm = FindWindow("ThunderDFrame", Me.Caption)
   Style = GetWindowLong(hWndForm, &HFFF0)
   Style = Style And Not &HC00000
   SetWindowLong hWndForm, &HFFF0, Style
   DrawMenuBar hWndForm
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = xlPrimaryButton And Shift = 1 Then
    Call ReleaseCapture
    Call SendMessage(hWndForm, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
  End If
End Sub

Private Sub CommandButton5_Click()

    Unload Me
    
End Sub

Private Sub CommandButton4_Click()

    Unload Me
    
End Sub


Private Sub CommandButton_super_Click()

    opinia = "Super! Bardzo fajna aplikacja!"
    
End Sub

Private Sub CommandButton_ok_Click()


    opinia = "Aplikacja jest w porz¹dku, ale mog³o byæ lepiej!"
    
End Sub

Private Sub CommandButton_zle_Click()

    opinia = "Nie podoba mi siê Wasza aplikacja!"
    
End Sub

Private Sub TextBox_opinia_Change()

    OpiniaUsera = TextBox_opinia.Value

End Sub

Private Sub UserForm_Activate()

Dim ThisSheet As Worksheet

Set ThisSheet = ActiveSheet
   
    If ThisSheet.Name = "KRAJ" Then
        ExitUF.LabelOpinion.Caption = "Podziel siê z nami opini¹ na temat aplikacji"
        ExitUF.Label1.Caption = "Dodatkowy komentarz:"
    ElseIf ThisSheet.Name = "RAPORT" Then
        ExitUF.LabelOpinion.Caption = "Podziel siê z nami opini¹ na temat aplikacji"
        ExitUF.Label1.Caption = "Dodatkowy komentarz:"
    ElseIf ThisSheet.Name = "COUNTRY" Then
        ExitUF.LabelOpinion.Caption = "Share with us your opinion about the app"
        ExitUF.Label1.Caption = "Additional comment:"
    ElseIf ThisSheet.Name = "REPORT" Then
        ExitUF.Label1.Caption = "Additional comment:"
    End If


End Sub
