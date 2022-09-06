Attribute VB_Name = "Poczta"
Option Explicit
Public Mail As String
Sub Mail1()
'Deklaracja zmiennych
Dim ThisSheet As Worksheet
Dim OutlookA As Outlook.Application 'aplikacja Outlook
Dim NowyMail As Outlook.MailItem 'nowa wiadomo�� e-mail

Set ThisSheet = ActiveSheet


If ThisSheet.Name = "KRAJ" Then
    Kraj = Range("B6")
    Call Eksport_pdf_pl
Else
    Kraj = Range("B6")
    Call Eksport_pdf_en
End If

OutlookUF.Show

'Uruchomienie Outlooka i stworzenie nowej wiadomo�ci
    Set OutlookA = New Outlook.Application
    Set NowyMail = OutlookA.CreateItem(olMailItem)

'Uzupe�nienie wiadomo�ci
    With NowyMail
        .Display 'wy�wietlenie okna
        .To = Mail 'dodanie adresat�w do pola wy�lij do
        .CC = "" 'dodanie adresat�w do pola kopia do
        .BCC = "" 'dodanie adresat�w do pola ukryta kopia do
        .Subject = "RAPORT COVID-19" ' temat wiadomo�ci
        .Body = "Cze��," & vbNewLine & vbNewLine & _
                "W za��czniku przesy�amy plik z raportem COVID-19" & vbNewLine & vbNewLine & _
                "Pozdrawiamy i dzi�kujemy za skorzystanie z naszej aplikacji!" & vbNewLine & _
                "Karolina Ogierman" & vbNewLine & _
                "Natalia Graczyk" & vbNewLine & _
                "Wojciech Bondaruk"
                'tre�� wiadomo�ci
        .Attachments.Add Plik 'dodanie za��cznika - plik z okre�lonej lokalizacji

'Wys�anie wiadomo�ci
        .Send
    End With
End Sub

