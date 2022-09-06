Attribute VB_Name = "Poczta"
Option Explicit
Public Mail As String
Sub Mail1()
'Deklaracja zmiennych
Dim ThisSheet As Worksheet
Dim OutlookA As Outlook.Application 'aplikacja Outlook
Dim NowyMail As Outlook.MailItem 'nowa wiadomoœæ e-mail

Set ThisSheet = ActiveSheet


If ThisSheet.Name = "KRAJ" Then
    Kraj = Range("B6")
    Call Eksport_pdf_pl
Else
    Kraj = Range("B6")
    Call Eksport_pdf_en
End If

OutlookUF.Show

'Uruchomienie Outlooka i stworzenie nowej wiadomoœci
    Set OutlookA = New Outlook.Application
    Set NowyMail = OutlookA.CreateItem(olMailItem)

'Uzupe³nienie wiadomoœci
    With NowyMail
        .Display 'wyœwietlenie okna
        .To = Mail 'dodanie adresatów do pola wyœlij do
        .CC = "" 'dodanie adresatów do pola kopia do
        .BCC = "" 'dodanie adresatów do pola ukryta kopia do
        .Subject = "RAPORT COVID-19" ' temat wiadomoœci
        .Body = "Czeœæ," & vbNewLine & vbNewLine & _
                "W za³¹czniku przesy³amy plik z raportem COVID-19" & vbNewLine & vbNewLine & _
                "Pozdrawiamy i dziêkujemy za skorzystanie z naszej aplikacji!" & vbNewLine & _
                "Karolina Ogierman" & vbNewLine & _
                "Natalia Graczyk" & vbNewLine & _
                "Wojciech Bondaruk"
                'treœæ wiadomoœci
        .Attachments.Add Plik 'dodanie za³¹cznika - plik z okreœlonej lokalizacji

'Wys³anie wiadomoœci
        .Send
    End With
End Sub

