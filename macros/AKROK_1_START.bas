Attribute VB_Name = "AKROK_1_START"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''WITAMY W APLIKACJI COVID-19'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Autorzy:
'Karolina Ogierman
'Natalia Graczyk
'Wojciech Bondaruk

'Na potrzeby aplikacji zosta� stworzony e-mail, na kt�rego mo�na przesy�a� uwagi dotycz�ce aplikacji:
'covid19app.opinie@gmail.com

'Apliacja Covid-19 jest narz�dziem s�u��cym do monitorowania zaka�e� na ca�ym �wiecie.
'Dane pochodz� ze strony:
'https://github.com/M-Media-Group/Covid-19-API

'U�ytkownik ma do wyboru monitorowanie zar�wno wska�nika zaka�e�, jak i wyzdrowie�, zgon�w oraz liczby szczepionek
'Pierwszy (g��wny) arkusz stanowi raport og�lny dla ca�ego �wiata
'Drugi arkusz stanowi arkusz, zawieraj�cy dane dla wybranego przez u�ytkownika kraju

'Dane od�wie�aj� si� z ka�dym otwarciem arkusza i makro wykonuj�ce aktualizacj� znajduje si� w Objects "Ten skoroszyt"


'Arkusz G��wny (RAPORT/REPORT):

'1. Modu� "Wskazniki_ogolne1" zawiera makro wyliczaj�ce podstawowe statystyk dla danych og�lnych (Arkusz RAPORT/REPORT)
'2. Modu� "Ustawienia" zawiera makra s�u��ce do obs�ugi panelu ustawie�
'3. Modu� "Kalendarz_ogolne" zawiera makro, s�u��ce do obs�ugi kalendarza

'Arkusz KRAJ/COUNTRY:

'1. Modu� "Metryczka" zawiera makro, kt�re za pomoc� funkcji VLOOKUP uzupe�nia metryczk� dla wybranego przez u�ytkownika kraju
'2. Modu� "Wskazniki_kraje" zawiera makro, wyliczaj�ce wska�niki zaka�e� dla wybranego przez u�ytkownika kraju
'3. Modu� "Wykresy" oraz modu� "Wykresy_raport" zawieraj� makra, kt�re podstawiaj� odpowiednie serie do wykres�w
'4. Modu� "Word_zapis" oraz "Powerpoint_zapis" zawieraj� makra, kt�re generuj� raporty w odpowiednim formacie
'5. Modu� "Poczta" zawiera makro, kt�re wysy�a na wskazany adres email raport PDF dla wybranego kraju
'6. Modu� "Poczta_opinia" zawiera makro, kt�re jest odpowiedzialne za wysy�k� na maila covidowego opinii dodanej przez u�ytkownika
'7. Modu� "ShowUF" jest modu�em pomocniczym, kt�re s�u�y do pokazania wybranych UserForm�w




