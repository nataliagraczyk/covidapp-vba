Attribute VB_Name = "Poczta_opinia"
Option Explicit
Public opinia As String
Public OpiniaUsera As String

Sub Mail_opinia()

Application.ScreenUpdating = False
'Deklaracja zmiennych
Dim ThisSheet As Worksheet
Dim OutlookAO As Outlook.Application 'aplikacja Outlook
Dim NowyMailO As Outlook.MailItem 'nowa wiadomoœæ e-mail

Set ThisSheet = ActiveSheet


ExitUF.Show

'Uruchomienie Outlooka i stworzenie nowej wiadomoœci
    Set OutlookAO = New Outlook.Application
    Set NowyMailO = OutlookAO.CreateItem(olMailItem)

'Uzupe³nienie wiadomoœci
    With NowyMailO
        .Display 'wyœwietlenie okna
        .To = "covid19app.opinie@gmail.com" 'dodanie adresatów do pola wyœlij do
        .CC = "" 'dodanie adresatów do pola kopia do
        .BCC = "" 'dodanie adresatów do pola ukryta kopia do
        .Subject = "Opinia o aplikacji COVID-19" ' temat wiadomoœci
        .Body = opinia & vbNewLine & vbNewLine & _
                OpiniaUsera
        .Send
    End With
    
Application.ScreenUpdating = True

End Sub



