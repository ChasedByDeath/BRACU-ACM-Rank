# -*- coding: utf-8 -*-
"""
Created on Fri Apr 17 12:22:11 2020

@author: Sadman Sakib
"""

import sys
import time as time
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import CFRestAPI as cf
   

def A1Notation(n):
   notation = chr(65+n%26)
   rem = n//26
   print("rem",rem)
   if n >= 26:
      return A1Notation(rem-1)+notation
   else:
      return notation

def addSheetUpdate(sheetId,title,rowCount,colCount):
   addSheetReq = {
      "addSheet": {
         "properties": {
            "sheetId": sheetId,
            "title": title,
            "gridProperties": {
               "rowCount": rowCount,
               "columnCount": colCount
            }
         }
      }   
   }
   return addSheetReq
def sheetPropertiesUpdate(sheetId,title,rowCount,colCount):
   sheetUpdateReq = {
     "updateSheetProperties": {
         "properties": {
            "sheetId":sheetId,
            "title": title,
            "gridProperties": {
               "rowCount": rowCount,
               "columnCount": colCount
            }
            
         },
         "fields": "*"
      }    
   }
   return sheetUpdateReq

def deleteDimensionCommand(sheetId,dim,startIndex,endIndex):
   deleteDimReq = {
      "deleteDimension": {
         "range": {
            "sheetId": sheetId,
            "dimension": dim,
            "startIndex": startIndex,
            "endIndex": endIndex
         }
      }
   }
   return deleteDimReq

def sortCommand(sheetId,row,col,sortOrderOne,dimensionIndexOne,sortOrderTwo,dimensionIndexTwo):
   sortReq = {
         "sortRange": {
            "range": {
               "sheetId": sheetId,
               "startRowIndex": row,
               "startColumnIndex": col,
            },
            "sortSpecs": [
               {
               "sortOrder": sortOrderOne,
               "dimensionIndex": dimensionIndexOne,
               },
               {
               "sortOrder": sortOrderTwo,
               "dimensionIndex": dimensionIndexTwo
               }
            ]
         }
      }
   return sortReq

def spreadSheetTitleUpdateCommand(title):
   titleUpdateReq = {      
      "updateSpreadsheetProperties": {
         "properties": {
            "title": title
         },
         "fields": "*"
      }
   }
   return titleUpdateReq

def widthHeightAdjustCommand(sheetId,dimension,startIndex,endIndex,pixelSize):
   adjustReq = {
      "updateDimensionProperties": {
            "range": {
                  "sheetId": sheetId,
                  "dimension": dimension,    #Values can be ROWS or COLUMNS or not used at all
                  "startIndex": startIndex,
                  "endIndex": endIndex
            },
            "properties": {
                  "pixelSize": pixelSize
            },
            "fields": "*"
      }
   }
   return adjustReq

def hideCommand(sheetId,start,end,dimension):
   colHideReq = {
         "updateDimensionProperties": {
               "range": {
                     "sheetId": sheetId,
                     "dimension": dimension,   #Value can only be ROWS or COLUMNS
                     "startIndex": start,
                     "endIndex": end
                     },
               "properties": {
                     "hiddenByUser": True,
                     },
               "fields": 'hiddenByUser',
            }
      }
   return colHideReq

def boldCommand(sheetId,rowStart,rowEnd,colStart,colEnd):
   rowBoldReq = { 
           "repeatCell":
               {
                "range": 
                  {
                      "sheetId": sheetId,
                      "startRowIndex": rowStart,
                      "endRowIndex": rowEnd,
                      "startColumnIndex": colStart,
                      "endColumnIndex": colEnd  
                  },
                "cell":
                  {
                      "userEnteredFormat":
                        {
                            "textFormat":
                              {
                                  "bold": True       
                              }
                        }
                  },
                "fields": "userEnteredFormat.textFormat.bold"
            }
      }
   return rowBoldReq

def repeatValueUpdateCommand(sheetId,value,startRow,endRow,startCol,endCol,color):
   valueType = ""
   red,green,blue = 1,1,1
   if isinstance(value,int):
      valueType = "numberValue"
   elif isinstance(value,str):
      valueType = "stringValue"
   elif isinstance(value,bool):
      valueType = "boolValue"
      
   if color.lower() == "yellow":
      blue = 0
   elif color.lower() == "red":
      green = 0
      blue = 0
   elif color.lower() == "cyan":
      red = 0
   elif color.lower() == "green":
      red = 0
      green = 0.5
      blue = 0
   elif color.lower() == "lime":
      red = 0
      blue = 0
   repeatValUpdateReq = {
         "repeatCell": {
            "range": {
               "sheetId": sheetId,
               "startRowIndex": startRow,
               "endRowIndex": endRow,
               "startColumnIndex": startCol,
               "endColumnIndex": endCol
            },
            "cell": {
               "userEnteredValue": {
                  valueType: value
               },
               "userEnteredFormat": {
                  "backgroundColor": {
                     "red": red,
                     "green": green,
                     "blue": blue
                  },
                  "horizontalAlignment": "CENTER",
                  "verticalAlignment": "MIDDLE"
               }
            },
            "fields": "*"
         }
      }
   return repeatValUpdateReq

def valueUpdateCommand(sheetId,value,row,column,color):
   valueType = ""
   red,green,blue = 1,1,1
   if isinstance(value,str):
      valueType = "stringValue"
   elif isinstance(value,int):
      valueType = "numberValue"
   elif isinstance(value,bool):
      valueType = "boolValue"
      
   if color.lower() == "yellow":
      blue = 0
   elif color.lower() == "red":
      green = 0
      blue = 0
   elif color.lower() == "cyan":
      red = 0
   elif color.lower() == "green":
      red = 0
      green = 0.5
      blue = 0
   elif color.lower() == "lime":
      red = 0
      blue = 0

   updateValReq = {
               "updateCells": {
                    "rows": [
                        {
                           "values": [
                              {
                                 "userEnteredValue": {
                                       valueType: value
                                 },
                                 "userEnteredFormat": {
                                       "backgroundColor": {
                                             "red": red,
                                             "green": green,
                                             "blue": blue
                                       },
                                       "horizontalAlignment": "CENTER",
                                       "verticalAlignment": "MIDDLE"            #There's an option of padding if needed
                                 },
                              }
                           ]
                        }
                     ],
                     "fields": "*",
                     "range": {
                           "sheetId": sheetId,
                           "startRowIndex": row,
                           "endRowIndex": row+1,
                           "startColumnIndex": column,
                           "endColumnIndex": column+1
                     }
                }
            }
   return updateValReq

def formattedValueUpdateCommand(sheetId,value,formattedVal,row,column,color):
   valueType = ""
   numberFormatType = ""
   red,green,blue = 1,1,1
   if isinstance(value,str):
      valueType = "stringValue"
      numberFormatType = "TEXT"
   elif isinstance(value,int):
      valueType = "numberValue"
      numberFormatType = "NUMBER"
   elif isinstance(value,bool):
      valueType = "boolValue"
      numberFormatType = "NUMBER"
      
   if color.lower() == "yellow":
      blue = 0
   elif color.lower() == "red":
      green = 0
      blue = 0
   elif color.lower() == "cyan":
      red = 0
   elif color.lower() == "green":
      red = 0
      green = 0.5
      blue = 0
   elif color.lower() == "lime":
      red = 0
      blue = 0

   formattedUpdateValReq = {
    "updateCells": {
        "rows": [{
            "values": [{
                "userEnteredValue": {
                    valueType: value
                },
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": red,
                        "green": green,
                        "blue": blue
                    },
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "numberFormat": {
                        "pattern": "\""+formattedVal+"\"",
                        "type": numberFormatType
                    }
                }
            }]
        }],
        "fields": "*",
        "range": {
            "sheetId": sheetId,
            "startRowIndex": row,
            "endRowIndex": row+1,
            "startColumnIndex": column,
            "endColumnIndex": column+1
        }
    }
   }  
   return formattedUpdateValReq



def updateSpreadSheetTitle():
   #Retrieving the current month's name as plain text
   curMonth = datetime.now().strftime("%B")
   #Retrieving the current year's name
   curYear = str(datetime.now().year)
   req.append(spreadSheetTitleUpdateCommand("BRACU CF RANK "+curMonth.upper()+curYear))
   
def insertNames(names):
   global fullSpreadSheet
   global rankingSheetID
   global req
   req.append(valueUpdateCommand(rankingSheetID,"Names",0,0,"white"))      #sheetId,value,row,column,color
   #This command will adjust the column width
   req.append(widthHeightAdjustCommand(rankingSheetID,"COLUMNS",0,1,175))     #sheetId,dimension,startIndex,endIndex,pixelSize
   #This command will bold the title
   req.append(boldCommand(rankingSheetID,0,1,0,1))      #sheetId,rowStart,rowEnd,colStart,colEnd
   for rowCount in range(0,len(names)):
      if len(names[rowCount])>0:    #Checking if it's empty cell or not
         req.append(valueUpdateCommand(rankingSheetID,names[rowCount][0],rowCount+3,0,"white"))           #sheetId,value,row,column,color
   
def addAContestant(name, handle):
   global fullSpreadSheet
   global contestantInfoSheet
   global rankingSheetID
   global contestantHandles
   errorOccurred = False
   rowCount = len(contestantHandles)+3
   req = []
   contestantReq = []
   contestantInfoSheetID = contestantInfoSheet.get_worksheet(0).id
   if cf.handleValidityCheck(handle):
      contestantReq.append(valueUpdateCommand(contestantInfoSheetID,name,rowCount,0,"white"))      #sheetId,value,row,column,color
      contestantReq.append(valueUpdateCommand(contestantInfoSheetID,handle,rowCount,1,"white"))    #sheetId,value,row,column,color
      
      req.append(valueUpdateCommand(rankingSheetID,name,rowCount,0,"white"))      #sheetId,value,row,column,color
      req.append(valueUpdateCommand(rankingSheetID,handle,rowCount,1,"white"))    #sheetId,value,row,column,color
      #Retrieving the previously calculated contest ID's
      previousContests = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!G1:1",{"valueRenderOption" : "UNFORMATTED_VALUE"})
      if len(previousContests["valueRanges"][0])>2:
         previousContests = previousContests["valueRanges"][0]["values"][0]
         previousContestsParticipantsCount = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!G2:2")
         previousContestsParticipantsCount = previousContestsParticipantsCount["valueRanges"][0]["values"][0]
         contestantsParticipation = 0
         contestantsTotalPoints = 0
         for x in range(0,len(previousContests)):
            cf.setContestID(previousContests[x])
            cf.setContestantHandle(handle)
            contestantsPoint = cf.getTotalPoints()
            if contestantsPoint == -2:
               errorOccurred = True
               print("There was some problem with the network or the server of the CodeForces")
               #req.append(valueUpdateCommand(rankingSheetID,"Error",rowCount,x+6,"red"))    #sheetId,value,row,column,color
            else:
               if contestantsPoint == -1:
                  req.append(valueUpdateCommand(rankingSheetID,-1,rowCount,x+6,"yellow"))      #sheetId,value,row,column,color
               else:
                  contestantsParticipation += 1
                  contestantsTotalPoints += contestantsPoint
                  req.append(valueUpdateCommand(rankingSheetID,contestantsPoint,rowCount,x+6,"lime"))      #sheetId,value,row,column,color
                  req.append(valueUpdateCommand(rankingSheetID,int(previousContestsParticipantsCount[x])+1,1,x+6,"white"))      #sheetId,value,row,column,color
         req.append(valueUpdateCommand(rankingSheetID,contestantsParticipation,rowCount,3,"white"))      #sheetId,value,row,column,color
         req.append(valueUpdateCommand(rankingSheetID,contestantsTotalPoints,rowCount,4,"white"))      #sheetId,value,row,column,color
      if not errorOccurred:
         fullSpreadSheet.batch_update({"requests": req})
         #Keeping the original contestant list updated
         contestantInfoSheet.batch_update({"requests": contestantReq})
   else:
      print("Invalid handle or some network/CF server problem occurred")
   
def removeAContestant(handle):
   global fullSpreadSheet
   global contestantInfoSheet
   global contestantHandles
   global rankingSheetID
   req = []
   contestantExist = False
   contestantRow = 0
   contestantInfoSheetID = contestantInfoSheet.get_worksheet(0).id
   #Forced to iterate like this becuase of the letter case and how the data is stored
   for x in range(0,len(contestantHandles)):
      if handle.upper() == contestantHandles[x][0].upper():
         #Updating the handle as it is stored in the sheet
         handle = contestantHandles[x][0]
         contestantExist = True
         contestantRow = x+3
         break
   if contestantExist:
      previousContestsParticipantsCount = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!G2:2")
      if len(previousContestsParticipantsCount["valueRanges"][0])>2:
         previousContestsParticipantsCount = previousContestsParticipantsCount["valueRanges"][0]["values"][0]
         contestsParticipation = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!G"+str(contestantRow+1)+":"+str(contestantRow+1))
         contestsParticipation = contestsParticipation["valueRanges"][0]["values"][0]
         for x in range(0,len(contestsParticipation)):
            if int(contestsParticipation[x]) >= 0:
               req.append(valueUpdateCommand(rankingSheetID,int(previousContestsParticipantsCount[x])-1,1,6+x,"white"))      #sheetId,value,row,column,color
      req.append(deleteDimensionCommand(rankingSheetID,"ROWS",contestantRow,contestantRow+1))      #sheetId,dim,startIndex,endIndex
      fullSpreadSheet.batch_update({"requests" : req})
      req = []
      '''
      Retrieving the handles of the contestants from the contestant info sheet
      Forced to do this because the index of the contestant handle won't same
      in both RankingSheet and ContestantInfoSheet because of sorting.
      '''
      handles = contestantInfoSheet.values_batch_get("'"+contestantInfoSheetTitle+"'!B4:B")
      handles = handles["valueRanges"][0]["values"]      #If there is no values in that column then it would give an error!
      contestantRow = handles.index([handle])+3
      req.append(deleteDimensionCommand(contestantInfoSheetID,"ROWS",contestantRow,contestantRow+1))      #sheetId,dim,startIndex,endIndex
      contestantInfoSheet.batch_update({"requests":req})
      #Removing the handle from the current handle list
      del contestantHandles[contestantRow-3]
   else:
      print("Given handle doesn't exist in the ranklist")
   
def insertHandlesAndLinks(handles):
   global fullSpreadSheet
   global rankingSheetID
   global contestantHandles
   global req
   req.append(valueUpdateCommand(rankingSheetID,"Handles",0,1,"white"))      #sheetId,value,row,column,color
   #This command will bold the value of cell(0,1)
   req.append(boldCommand(rankingSheetID,0,1,1,2))      #sheetId,rowStart,rowEnd,colStart,colEnd
   #This command will set the width of column 1 to given pixelSize
   req.append(widthHeightAdjustCommand(rankingSheetID,"COLUMNS",1,2,120))     #sheetId,dimension,startIndex,endIndex,pixelSize
   req.append(valueUpdateCommand(rankingSheetID,"Profile Links",0,2,"white"))      #sheetId,value,row,column,color
   #This command will bold the value of cell(0,2)
   req.append(boldCommand(rankingSheetID,0,1,2,3))      #sheetId,rowStart,rowEnd,colStart,colEnd
   for rowCount in range(0,len(handles)):
      if len(handles[rowCount])>0:      #Checking if it's empty cell or not
         #Doing this maintain the 5 requests per second rule of CF API
         time.sleep(0.1)
         if cf.handleValidityCheck(handles[rowCount][0]):
            #Keeping a copy of the handles in a global list
            contestantHandles.append(handles[rowCount])  
            req.append(valueUpdateCommand(rankingSheetID,handles[rowCount][0],rowCount+3,1,"white"))       #sheetId,value,row,column,color
            req.append(valueUpdateCommand(rankingSheetID,"https://codeforces.com/profile/"+handles[rowCount][0],rowCount+3,2,"white"))         #sheetId,value,row,column,color
         else:
            #Keeping a copy of the handles in a global list
            handles[rowCount][0]="Invalid Handle!!"
            contestantHandles.append(handles[rowCount])
            req.append(valueUpdateCommand(rankingSheetID,"Invalid Handle!!",rowCount+3,1,"red"))      #sheetId,value,row,column,color
            req.append(valueUpdateCommand(rankingSheetID,"Invalid Handle!!",rowCount+3,2,"red"))      #sheetId,value,row,column,color
   #This command hides column 3
   req.append(hideCommand(rankingSheetID,2,3,"COLUMNS"))     #sheetId,start,end,dimension

def setCurrentRankingSheet(sheetURL):
   global client
   global contestantHandles
   global contestantInfoSheet
   global contestantInfoSheetTitle
   global fullSpreadSheet
   global rankingSheetID
   global rankingSheetTitle
   global req
   try:
      fullSpreadSheet = client.open_by_url(sheetURL)
   except:
      print(sys.exc_info()[0])
      return False
   rankingSheetTitle = fullSpreadSheet.get_worksheet(0).title
   rankingSheetID = fullSpreadSheet.get_worksheet(0).id
   handles = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!B:B")
   #Checking if this is a new sheet! If it is then storing the initial stuffs in it.
   if len(handles["valueRanges"][0])<3:
      req = []
      updateSpreadSheetTitle()
      req.append(sheetPropertiesUpdate(rankingSheetID,"Ranking",500,500))       #sheetId,title,rowCount,colCount
      rankingSheetTitle = "Ranking"
      #Adding the point system details sheet
      req.append(addSheetUpdate(1,"Details", 1, 1))     #sheetId,title,rowCount,colCount
      #Adding the point system data
      req.append(valueUpdateCommand(1,contestantInfoSheet.get_worksheet(1).cell(1,1).value,0,0,"cyan"))    #sheetId,value,row,column,color
      req.append(widthHeightAdjustCommand(1,"COLUMNS",0,1,700))       #sheetId,dimension,startIndex,endIndex,pixelSize
      req.append(widthHeightAdjustCommand(1,"ROWS",0,1,300))       #sheetId,dimension,startIndex,endIndex,pixelSize
      #Retrieving the names of the contestants from the contestant info sheet
      names = contestantInfoSheet.values_batch_get("'"+contestantInfoSheetTitle+"'!A4:A")
      names = names["valueRanges"][0]["values"]      #If there is no values in that column then it would give an error!
      #Retrieving the handles of the contestants from the contestant info sheet
      handles = contestantInfoSheet.values_batch_get("'"+contestantInfoSheetTitle+"'!B4:B")
      handles = handles["valueRanges"][0]["values"]      #If there is no values in that column then it would give an error!
      #Since req is global, all the requests will be appended to it
      insertNames(names)
      insertHandlesAndLinks(handles)
      req.append(valueUpdateCommand(rankingSheetID,"Participation",0,3,"white"))     #sheetId,value,row,column,color
      #This command will set the width of column 4 to given pixelSize
      req.append(widthHeightAdjustCommand(rankingSheetID,"COLUMNS",3,4,80))     #sheetId,dimension,startIndex,endIndex,pixelSize
      #This command will bold the value of cell(0,3)
      req.append(boldCommand(rankingSheetID,0,1,3,4))      #sheetId,rowStart,rowEnd,colStart,colEnd
      req.append(repeatValueUpdateCommand(rankingSheetID,0,3,len(names)+3,3,4,"white"))    #sheetId,value,startRow,endRow,startCol,endCol,color
      req.append(valueUpdateCommand(rankingSheetID,"Total Points",0,4,"white"))     #sheetId,value,row,column,color
      #This command will set the width of column 5 to given pixelSize
      req.append(widthHeightAdjustCommand(rankingSheetID,"COLUMNS",4,5,80))     #sheetId,dimension,startIndex,endIndex,pixelSize
      req.append(boldCommand(rankingSheetID,0,1,4,5))      #sheetId,rowStart,rowEnd,colStart,colEnd
      req.append(repeatValueUpdateCommand(rankingSheetID,0,3,len(names)+3,4,5,"white"))    #sheetId,value,startRow,endRow,startCol,endCol,color
      #This command will set the width of column 6 to given pixelSize
      req.append(widthHeightAdjustCommand(rankingSheetID,"COLUMNS",5,6,40))     #sheetId,dimension,startIndex,endIndex,pixelSize
      fullSpreadSheet.batch_update({"requests": req})
   else:
      contestantHandles = handles["valueRanges"][0]["values"]     #If there is no values in that column then it would give an error!
      #Slicing to remove the first three values which are not handles
      contestantHandles = contestantHandles[3:len(contestantHandles)]
   return True

def initializeRankingUpdate():
   global contestantHandles
   global rankingSheetID
   global req
   errorOccurred = False
   req = []
   contestCalculated = False
   participantsCount = 0
   contestColumn = 0
   contestID = cf.getContestID()
   contestName = cf.getContestName()
   #Retrieving the previously calculated contest ID's
   previousContests = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!G1:1",{"valueRenderOption" : "UNFORMATTED_VALUE"})
   if len(previousContests["valueRanges"][0])>2:
      previousContests = previousContests["valueRanges"][0]["values"][0]
      #Checking if it has already been calculated
      contestColumn = 6+len(previousContests)
      if contestID in previousContests:
         contestCalculated = True         
         print("This contest has already been calculated")
   else:
      contestColumn = 6
   
   if not contestCalculated:
      req.append(formattedValueUpdateCommand(rankingSheetID,contestID,contestName,0,contestColumn,"white"))      #sheetId,value,formattedVal,row,column,color
      req.append(boldCommand(rankingSheetID,0,1,contestColumn,contestColumn+1))     #sheetId,rowStart,rowEnd,colStart,colEnd
      req.append(widthHeightAdjustCommand(rankingSheetID,"COLUMNS",contestColumn,contestColumn+1,85))       #sheetId,dimension,startIndex,endIndex,pixelSize
      participationList = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!D4:D")
      participationList = participationList["valueRanges"][0]["values"]
      totalPointsList = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!E4:E")
      totalPointsList = totalPointsList["valueRanges"][0]["values"]
      for rowCount in range(0,len(contestantHandles)):
         if contestantHandles[rowCount][0] != "Invalid Handle!!":
            print(contestantHandles[rowCount][0])
            cf.setContestantHandle(contestantHandles[rowCount][0])
            contestantsPoint = cf.getTotalPoints()
            #Doing this maintain the 5 requests per second rule of CF API
            time.sleep(0.1)
            if contestantsPoint == -2:
               errorOccurred = True
               print("There was some problem with the network or the server of the CodeForces")
               #req.append(valueUpdateCommand(rankingSheetID,"Error",rowCount+3,contestColumn,"red"))    #sheetId,value,row,column,color
            else:
               if contestantsPoint == -1:
                  req.append(valueUpdateCommand(rankingSheetID,-1,rowCount+3,contestColumn,"yellow"))      #sheetId,value,row,column,color
               else:
                  participantsCount += 1
                  req.append(valueUpdateCommand(rankingSheetID,contestantsPoint,rowCount+3,contestColumn,"lime"))      #sheetId,value,row,column,color
                  req.append(valueUpdateCommand(rankingSheetID,int(participationList[rowCount][0])+1,rowCount+3,3,"white"))       #sheetId,value,row,column,color
                  req.append(valueUpdateCommand(rankingSheetID,int(totalPointsList[rowCount][0])+contestantsPoint,rowCount+3,4,"white"))       #sheetId,value,row,column,color
         else:
            req.append(valueUpdateCommand(rankingSheetID,-2,rowCount+3,contestColumn,"red"))      #sheetId,value,row,column,color
      req.append(valueUpdateCommand(rankingSheetID,participantsCount,1,contestColumn,"white"))     #sheetId,value,row,column,color
      if not errorOccurred:
         fullSpreadSheet.batch_update({"requests" : req})

def removeAContest(contestID):
   global fullSpreadSheet
   global rankingSheetID
   global req
   req = []
   contestColumn = 0
   contestCalculated = False
   #Retrieving the previously calculated contest ID's
   previousContests = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!G1:1",{"valueRenderOption" : "UNFORMATTED_VALUE"})
   if len(previousContests["valueRanges"][0])>2:
      previousContests = previousContests["valueRanges"][0]["values"][0]
      #Checking if it has already been calculated
      if contestID in previousContests:
         contestCalculated = True
         contestColumn = 6+previousContests.index(contestID)
      
   if contestCalculated:
      curContestValues = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!"+A1Notation(contestColumn)+"4:"+A1Notation(contestColumn))
      curContestValues = curContestValues["valueRanges"][0]["values"]
      participationList = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!D4:D")
      participationList = participationList["valueRanges"][0]["values"]
      totalPointsList = fullSpreadSheet.values_batch_get("'"+rankingSheetTitle+"'!E4:E")
      totalPointsList = totalPointsList["valueRanges"][0]["values"]
      for rowCount in range(0,len(curContestValues)):
         if int(curContestValues[rowCount][0]) >=0:
            req.append(valueUpdateCommand(rankingSheetID,int(participationList[rowCount][0])-1,rowCount+3,3,"white"))       #sheetId,value,row,column,color
            req.append(valueUpdateCommand(rankingSheetID,int(totalPointsList[rowCount][0])-int(curContestValues[rowCount][0]),rowCount+3,4,"white"))       #sheetId,value,row,column,color
      req.append(deleteDimensionCommand(rankingSheetID,"COLUMNS",contestColumn,contestColumn+1))      #sheetId,dim,startIndex,endIndex
      fullSpreadSheet.batch_update({"requests":req})
   else:
      print("No contest exist in the ranking sheet with ID = ",contestID)
      
def sortRankSheet():
   global fullSpreadSheet
   global req
   req = []
   req.append(sortCommand(rankingSheetID,3,0,"DESCENDING",4,"ASCENDING",3))     #sheetId,row,col,sortOrderOne,dimensionIndexOne,sortOrderTwo,dimensionIndexTwo
   fullSpreadSheet.batch_update({"requests": req})
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("BRACUCFRankCreds.json",scope)
client = gspread.authorize(creds)

contestantInfoSheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1xgQUh8lIJQJTxiLcJBsSWOAkZqGk3W6KrybgCku2Ghg/edit?usp=sharing")
contestantInfoSheetTitle = contestantInfoSheet.get_worksheet(0).title
fullSpreadSheet = ""
rankingSheetID = ""
rankingSheetTitle = ""
contestantHandles = []
req = []
#setCurrentRankingSheet("https://docs.google.com/spreadsheets/d/1xgQUh8lIJQJTxiLcJBsSWOAkZqGk3W6KrybgCku2Ghg/edit?usp=sharing")

