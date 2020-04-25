# -*- coding: utf-8 -*-
"""
Created on Sun Apr 17 18:07:12 2020

@author: Sadman Sakib
"""

import CFRestAPI as cf
import SheetPage as sp
'''
   PLEASE READ THE README.txt TO SETUP THE ENVIRONMENT.
   And to know some key information.
'''

'''
   While executing any of the function keep the other functions commented out.
'''

'''
   Please enter the sheet URL. And always have the link in your code. Otherwise it won't run.
   The sheet of the given link has to be an empty one. If it's not then there might be some
   undefined behavior. Also the link has to be shared with the service account with the 
   editor access.
   You can use this URL for experiment, it has been shared with the given service account-
   "https://docs.google.com/spreadsheets/d/1tbQ2lPjSNVZhPCmoNTqaTti3KKSS4sRVPyw9XO6cYj4/edit?usp=sharing"
   Open this spreadsheet before running the project to see how the changes were made. If there's
   already some data(as someone else can update it) please delete all the data and the
   additional sheets. Otherwise it might not work.
'''
URL = "Contest URL goes here"

'''
   This may take some time if the sheet is a new one since all the values like
   names,handles,profiles links and other adjustments to the spreadsheet has to
   be copied from contestantInfoSheet. Please have patience in that case.
'''
#Checking if the given url is a valid one
if sp.setCurrentRankingSheet(URL):
   '''
   Please enter the contest ID
   You can even make list of the contest IDs and then loop through it. Like following-
   contestList = [1330,1333,1334,1339,1335,1337,1343,1341]
   for contestID in contestList:
      if cf.setContestID(contestID):
         sp.initializeRankingUpdate()
      else:
         print("Entered contest ID seems to be invalid")
   '''
   contestID = "Contest ID goes here(Must be an Integer)"
   #Checking if the given contest ID is a valid one
   if cf.setContestID(contestID):
      sp.initializeRankingUpdate()
   else:
      print("Entered contest ID seems to be invalid")
      
   #To add a contestant
   '''
   sp.addAContestant("Name","handle")
   '''
   #To remove a contestant
   '''
   sp.removeAContestant("handle")
   '''
   #To remove a contest
   '''
   sp.removeAContest(contestID)
   '''
   #To know the point of a contestant in a particular contest
   '''
   if cf.setContestID("contestID"):
      cf.setContestantHandle("Handle")
      print(cf.getTotalPoints())
   '''
   sp.sortRankSheet()
else:
   print("The given URL might be invalid or not associated with your account.")