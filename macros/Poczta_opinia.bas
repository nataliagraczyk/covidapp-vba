Attribute VB_Name = "Poczta_opinia"
Option Explicit
Public opinia As String
Public OpiniaUsera As String

Sub Mail_opinia()

Application.ScreenUpdating = False
'Deklaracja zmiennych
Dim ThisSheet As Worksheet
Dim OutlookAO As Outlook.Application 'aplikacja Outlook
Dim NowyMailO As Outlook.MailItem 'nowa wiadomo�� e-mail

Set ThisSheet = ActiveSheet


ExitUF.Show

'Uruchomienie Outlooka i stworzenie nowej wiadomo�ci
    Set OutlookAO = New Outlook.Application
    Set NowyMailO = OutlookAO.CreateItem(olMailItem)

'Uzupe�nienie wiadomo�ci
    With NowyMailO
        .Display 'wy�wietlenie okna
        .To = "covid19app.opinie@gmail.com" 'dodanie adresat�w do pola wy�lij do
        .CC = "" 'dodanie adresat�w do pola kopia do
        .BCC = "" 'dodanie adresat�w do pola ukryta kopia do
        .Subject = "Opinia o aplikacji COVID-19" ' temat wiadomo�ci
        .Body = opinia & vbNewLine & vbNewLine & _
                OpiniaUsera
        .Send
    End With
    
Application.ScreenUpdating = True

End Sub



