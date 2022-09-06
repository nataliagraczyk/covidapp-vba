VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WelcomeUF 
   Caption         =   "COVID-19 App"
   ClientHeight    =   2920
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8070
   OleObjectBlob   =   "WelcomeUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WelcomeUF"
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

Private Sub LabelExit_Click()

    Application.Quit
    
End Sub

Private Sub LabelNext_Click()

    If WelcomeUF.OB_Polski Then
    
        Sheets("REPORT").Visible = False
        Sheets("RAPORT").Visible = True
        Sheets("COUNTRY").Visible = False
        Sheets("KRAJ").Visible = False
        Sheets("RAPORT").Select
    
    Else
    
        Sheets("REPORT").Visible = True
        Sheets("RAPORT").Visible = False
        Sheets("COUNTRY").Visible = False
        Sheets("KRAJ").Visible = False
        Sheets("REPORT").Select
        
    End If

    Unload WelcomeUF

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Button = xlPrimaryButton And Shift = 1 Then
    Call ReleaseCapture
    Call SendMessage(hWndForm, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
  End If
End Sub


Private Sub CB_Next_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub OB_English_Click()

    Dim LastUpdate As Date
    LastUpdate = Mid(Sheets("H_deaths").Range("B1").Value, 17, 10)

    WelcomeUF.LabelWelcome.Caption = "Welcome, " & Application.UserName & "!"
    WelcomeUF.LabelComment.Caption = "Check out the COVID-19 App and be up-to-date with the pandemic situation."
    WelcomeUF.LabelLastUpdate = "Last data update: " & LastUpdate
    WelcomeUF.LabelSelectLang = "Select app language:"
    WelcomeUF.LabelNext.Caption = "Next"
    WelcomeUF.LabelExit.Caption = "Exit"

    
End Sub

Private Sub OB_Polski_Click()

    Dim LastUpdate As Date
    LastUpdate = Mid(Sheets("H_deaths").Range("B1").Value, 17, 10)

    WelcomeUF.LabelWelcome.Caption = "Witaj, " & Application.UserName & "!"
    WelcomeUF.LabelComment.Caption = "SprawdŸ aplikacjê COVID-19 App i b¹dŸ na bie¿¹co z sytuacj¹ pandemiczn¹."
    WelcomeUF.LabelLastUpdate = "Ostatnia aktualizacja danych: " & LastUpdate
    WelcomeUF.LabelSelectLang = "Wybierz jêzyk aplikacji:"
    WelcomeUF.LabelNext.Caption = "Dalej"
    WelcomeUF.LabelExit.Caption = "Wyjœcie"
    

End Sub


Private Sub UserForm_Initialize()

Dim hWndForm As Long
   Dim Style As Long, Menu As Long
   hWndForm = FindWindow("ThunderDFrame", Me.Caption)
   Style = GetWindowLong(hWndForm, &HFFF0)
   Style = Style And Not &HC00000
   SetWindowLong hWndForm, &HFFF0, Style
   DrawMenuBar hWndForm

    Dim LastUpdate As Date
    LastUpdate = Mid(Sheets("H_deaths").Range("B1").Value, 17, 10)

    WelcomeUF.LabelTodaysDate.Caption = WeekdayName(Weekday(Date, vbUseSystemDayOfWeek)) & ", " & Date

    WelcomeUF.LabelWelcome.Caption = "Welcome, " & Application.UserName & "!"
    WelcomeUF.LabelLastUpdate = "Last data update: " & LastUpdate


End Sub
