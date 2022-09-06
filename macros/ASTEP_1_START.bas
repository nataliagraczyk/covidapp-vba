Attribute VB_Name = "AKROK_1_START"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''WELCOME IN COVID-19 APPLICATION'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Authors:
'Karolina Ogierman
'Natalia Graczyk
'Wojciech Bondaruk

'An email has been set up for comments on the application:
'covid19app.opinie@gmail.com

'The Covid-19 app is a tool for monitoring infections worldwide.
'The data comes from the website:
'https://github.com/M-Media-Group/Covid-19-API

'The user can choose to monitor both infection rates and recoveries, deaths and vaccine counts.
'The first (main) sheet is a global report for the whole world
'The second sheet is a sheet containing data for the country of the user's choice

'The data refreshes each time the worksheet is opened and the macro that performs the update is in Objects 'This workbook'


'Main sheet (REPORT/REPORT):

'1. The "Indicators_general1" module contains a macro that calculates basic statistics for the general data (REPORT/REPORT sheet).
'2. The "Settings" module contains macros for operating the settings panel.
'3. The "Calendar_general" module contains a macro for operating the calendar.

'COUNTRY/COUNTRY sheet:

'1. The "Metrics" module contains a macro which, using the VLOOKUP function, completes a metric for the country selected by the user.
'2. The "Indicators_countries" module contains a macro which calculates the infection rates for the country selected by the user
'3. The "Graphs" module and the "Graphs_report" module contain macros which substitute the relevant series into the graphs
'4. The 'Word_Record' and 'Powerpoint_Record' modules contain macros that generate reports in the appropriate format.
'5. The "Mail" module contains a macro which sends a PDF report for a selected country to a specified e-mail address
'6. The "Mail_opinion" module contains a macro which is responsible for sending a covidien opinion added by the user to an email address
'7. The "ShowUF" module is an auxiliary module which is used to show selected UserForms
