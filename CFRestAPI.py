# -*- coding: utf-8 -*-
"""
Created on Sat Mar 28 13:26:54 2020

@author: Sadman Sakib
"""

import requests
import json

def setContestID(id):
   #Setting the contest id
   global contestID
   contestID = id
   return getContestInfo()

def setContestantHandle(handle):
   #Setting the contestant handle
   global contestantHandle
   contestantHandle = handle

def getContestInfo():
   global problems
   global pointsPerProblem
   global totalPoints
   global contestName
   temp = 1
   #Receiving the contest details
   contestInfo = requests.get('https://codeforces.com/api/contest.standings?contestId='
                              +str(contestID)+'&from=1&count=1')
   data = json.loads(contestInfo.text)
   #Checking if the contestID was a valid one
   if data["status"] != "OK":
      print(data)
      return False
      #print(data["comment"])
   else:
      contestName = data["result"]["contest"]["name"]
      print(contestName)
      #Slicing it since the whole name can't be allocated in the sheet
      if contestName[:contestName.find(" ")] == "Educational":
         contestName = contestName[contestName.find("Round "):contestName.find(" (")]
         contestName = "Edu"+contestName
      else:
         if "#" in contestName:
            contestName = contestName[contestName.find("#"):]
            contestName = contestName[:contestName.find(")")]+")"
      problems = data["result"]["problems"]
      
      #Giving the problems it's weight to count total score
      for aProblem in problems:
         problemIndex = aProblem["index"]
         pointsPerProblem[problemIndex]=temp
         temp += 1
   return True
         
def problemPreProcess():
   global problems
   global alreadyCounted
   #Marking the questions as unsolved
   for aProblem in problems:
         problemIndex = aProblem["index"]
         alreadyCounted[problemIndex]=False
   
def getContestantInfo():
   global contestantInfo
   global totalPoints
   totalPoints = -1
   participationFlag = False
   #print("contestantHandle = ",contestantHandle)
   #Receiving the contestant's participation details for the contest
   contestantInfo = requests.get('https://codeforces.com/api/contest.status?contestId='
                                 +str(contestID)+'&handle='+contestantHandle)
   data = json.loads(contestantInfo.text)
   #Checking if the contestant handle and contest id was a valid one
   if data["status"] != "OK":
      totalPoints = -2     #So if the handle was a invalid one then the totalPoints will be -2
      print(data["comment"])
   else:
      submissions = data["result"]
      #print(submissions)
      #Iterating through all the submissions to check whether it was solved during contest time
      #And counting score if it was
      for singleSubmission in submissions:
         problemIndex = singleSubmission["problem"]["index"]
         programmerType = singleSubmission["author"]["participantType"]
         verdict = singleSubmission["verdict"]
         if(programmerType=="CONTESTANT" and not participationFlag):
             totalPoints = 5
             participationFlag = True
         if(verdict=="OK" and programmerType=="CONTESTANT" and not alreadyCounted[problemIndex]):
             alreadyCounted[problemIndex] = True
             totalPoints += pointsPerProblem[problemIndex]*5
def getTotalPoints():
   problemPreProcess()
   getContestantInfo()
   print("points = ",totalPoints)
   return totalPoints

def getContestName():
   global contestName
   return contestName

def getContestID():
   global contestID
   return contestID

def handleValidityCheck(handle):
   userInfo = requests.get("https://codeforces.com/api/user.info?handles="+handle)
   data = json.loads(userInfo.text)
   if data["status"] != "OK":
      print(data["comment"])
      return False
   else:
      return True

contestID = -1
contestName = ""
contestantInfo = ""
contestantHandle = ""
problems = ""
#This dictionary is to maintain the weight of the problems
pointsPerProblem = {}
#This dictionary is to prevent two or more accepted solution from getting counted twice or more
alreadyCounted = {}
totalPoints = 0