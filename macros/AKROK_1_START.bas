Attribute VB_Name = "AKROK_1_START"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''WITAMY W APLIKACJI COVID-19'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Autorzy:
'Karolina Ogierman
'Natalia Graczyk
'Wojciech Bondaruk

'Na potrzeby aplikacji zosta³ stworzony e-mail, na którego mo¿na przesy³aæ uwagi dotycz¹ce aplikacji:
'covid19app.opinie@gmail.com

'Apliacja Covid-19 jest narzêdziem s³u¿¹cym do monitorowania zaka¿eñ na ca³ym œwiecie.
'Dane pochodz¹ ze strony:
'https://github.com/M-Media-Group/Covid-19-API

'U¿ytkownik ma do wyboru monitorowanie zarówno wskaŸnika zaka¿eñ, jak i wyzdrowieñ, zgonów oraz liczby szczepionek
'Pierwszy (g³ówny) arkusz stanowi raport ogólny dla ca³ego œwiata
'Drugi arkusz stanowi arkusz, zawieraj¹cy dane dla wybranego przez u¿ytkownika kraju

'Dane odœwie¿aj¹ siê z ka¿dym otwarciem arkusza i makro wykonuj¹ce aktualizacjê znajduje siê w Objects "Ten skoroszyt"


'Arkusz G³ówny (RAPORT/REPORT):

'1. Modu³ "Wskazniki_ogolne1" zawiera makro wyliczaj¹ce podstawowe statystyk dla danych ogólnych (Arkusz RAPORT/REPORT)
'2. Modu³ "Ustawienia" zawiera makra s³u¿¹ce do obs³ugi panelu ustawieñ
'3. Modu³ "Kalendarz_ogolne" zawiera makro, s³u¿¹ce do obs³ugi kalendarza

'Arkusz KRAJ/COUNTRY:

'1. Modu³ "Metryczka" zawiera makro, które za pomoc¹ funkcji VLOOKUP uzupe³nia metryczkê dla wybranego przez u¿ytkownika kraju
'2. Modu³ "Wskazniki_kraje" zawiera makro, wyliczaj¹ce wskaŸniki zaka¿eñ dla wybranego przez u¿ytkownika kraju
'3. Modu³ "Wykresy" oraz modu³ "Wykresy_raport" zawieraj¹ makra, które podstawiaj¹ odpowiednie serie do wykresów
'4. Modu³ "Word_zapis" oraz "Powerpoint_zapis" zawieraj¹ makra, które generuj¹ raporty w odpowiednim formacie
'5. Modu³ "Poczta" zawiera makro, które wysy³a na wskazany adres email raport PDF dla wybranego kraju
'6. Modu³ "Poczta_opinia" zawiera makro, które jest odpowiedzialne za wysy³kê na maila covidowego opinii dodanej przez u¿ytkownika
'7. Modu³ "ShowUF" jest modu³em pomocniczym, które s³u¿y do pokazania wybranych UserFormów




