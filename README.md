# The COVID-19 App to monitor the number of coronavirus infections worldwide
This application was created for the class "Excel applications in the enterprise using the VBA part II".

The authors of the application are: Natalia Graczyk, Karolina Ogierman, Wojciech Bondaruk

The app was selected as one of the two best apps in the class and was featured on the Labmasters website as an accolade: [click](https://labmasters.pl/covid_19app/)

## Project Objective

The COVID-19 App is used to monitor the number of coronavirus infections worldwide. It is fully interactive and allows the user to freely customize the generated report by selecting: 
- the exact day from which the report is to be generated 
- country 
- indicators (number of cases / recoveries / deaths / vaccinations) 

The big advantage of the application is also the possibility of exporting reports to the selected format: PDF, MS Word, PowerPoint, or sending the report by e-mail

## Installation

To launch the application, download the Covid_19App.xlsm file and remember to enable macros in Excel.

When the application opens, the user interface appears, where we can select the day of the report, change the application settings or proceed to generate the report. To generate the report, simply select the continent and country and then specify in which form the report is to be saved.

## Data

The app uses dynamic data pulled into Excel by making API calls with Power Query. The dataset refreshes each time user opens the app

API: [click](https://github.com/M-Media-Group/Covid-19-API)

## Methods

Main sheet (REPORT):

- The "Indicators_general1" module contains a macro that calculates basic statistics for the general data (REPORT/REPORT sheet).
- The "Settings" module contains macros for operating the settings panel.
- The "Calendar_general" module contains a macro for operating the calendar.

COUNTRY sheet:

- The "Metrics" module contains a macro which, using the VLOOKUP function, completes a metric for the country selected by the user.
- The "Indicators_countries" module contains a macro which calculates the infection rates for the country selected by the user
- The "Graphs" module and the "Graphs_report" module contain macros which substitute the relevant series into the graphs
- The 'Word_Record' and 'Powerpoint_Record' modules contain macros that generate reports in the appropriate format.
- The "Mail" module contains a macro which sends a PDF report for a selected country to a specified e-mail address
- The "Mail_opinion" module contains a macro which is responsible for sending a covidien opinion added by the user to an email address
- The "ShowUF" module is an auxiliary module which is used to show selected UserForms

