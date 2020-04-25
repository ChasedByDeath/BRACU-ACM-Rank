# -*- coding: utf-8 -*-
"""
Created on Sun Apr 26 00:30:27 2020

@author: Sadman Sakib
"""

Firstly, you have to have python3 installed in your machine. If it's not installed in your
machine then go to "https://www.python.org/downloads/" and install Python 3.7 or up in your
machine. Once you have python3 installed and running in your machine, you have to install 
gspread to run the project. To install gspread, run this "pip install gspread oauth2client"
command from your IDE's terminal. Once you have these installed, you're good to go.

Since this project doesn't have any UI, please interact through the main.py file.
Detailed information about how to interact is mentioned there.

I've added a service account here, but there's a limit of 100 requests per 100 second in 
google sheets api. So if many people try to run the project with this service account at a
certain time, you might get a quota error. To avoid that, you can create a project in 
"console.cloud.google.com" and associate it with this project. To do that, create a project in
the mentioned link, enable drive and sheets api there. Then create a service account from "APIs and services"
and download the credentials. After that, overwrite the BRACUCFRankCreds.json with the data
in your credentials json file. After that, you have to create your own contestantInfoSheet and 
share it with the email of your service account. To see how the data need look in your contestantInfoSheet
file,
"https://docs.google.com/spreadsheets/d/1SMynUiRj7MzCii97eht7NkWMJ4hFvxybgyzDH5_JzfQ/edit?usp=sharing"
Start writing name and handles exactly from row 4 and in collumn 1,2 accordingly.
After that copy your contestantInfoSheet file and paste it in the sheetPage.py in
"contestantInfoSheet" variable. Now you're set up to use the system.

You may follow this tutorial to set up the environment and service account too-
https://www.youtube.com/watch?v=cnPlKLEGR7E&t=271s