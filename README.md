# The COVID-19 App to monitor the number of coronavirus infections worldwide
This application was created for the class "Excel applications in the enterprise using the VBA part II".

The authors of the application are: Natalia Graczyk, Karolina Ogierman, Wojciech Bondaruk

The app was selected as one of the two best apps in the class and was featured on the Labmasters website as an accolade: [click](https://labmasters.pl/covid_19app/)

![image](https://github.com/nataliagraczyk/covidapp-vba/tree/main/templates/main_report.jpg?raw=true)

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
