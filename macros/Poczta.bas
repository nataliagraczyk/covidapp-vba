Attribute VB_Name = "Poczta"
Option Explicit
Public Mail As String
Sub Mail1()
'Deklaracja zmiennych
Dim ThisSheet As Worksheet
Dim OutlookA As Outlook.Application 'aplikacja Outlook
Dim NowyMail As Outlook.MailItem 'nowa wiadomość e-mail

Set ThisSheet = ActiveSheet


If ThisSheet.Name = "KRAJ" Then
    Kraj = Range("B6")
    Call Eksport_pdf_pl
Else
    Kraj = Range("B6")
    Call Eksport_pdf_en
End If

OutlookUF.Show

'Uruchomienie Outlooka i stworzenie nowej wiadomości
    Set OutlookA = New Outlook.Application
    Set NowyMail = OutlookA.CreateItem(olMailItem)

'Uzupełnienie wiadomości
    With NowyMail
        .Display 'wyświetlenie okna
        .To = Mail 'dodanie adresatów do pola wyślij do
        .CC = "" 'dodanie adresatów do pola kopia do
        .BCC = "" 'dodanie adresatów do pola ukryta kopia do
        .Subject = "RAPORT COVID-19" ' temat wiadomości
        .Body = "Cześć," & vbNewLine & vbNewLine & _
                "W załączniku przesyłamy plik z raportem COVID-19" & vbNewLine & vbNewLine & _
                "Pozdrawiamy i dziękujemy za skorzystanie z naszej aplikacji!" & vbNewLine & _
                "Karolina Ogierman" & vbNewLine & _
                "Natalia Graczyk" & vbNewLine & _
                "Wojciech Bondaruk"
                'treść wiadomości
        .Attachments.Add Plik 'dodanie załącznika - plik z określonej lokalizacji

'Wysłanie wiadomości
        .Send
    End With
End Sub

