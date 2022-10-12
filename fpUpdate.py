from __future__ import print_function

import os.path
import requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1T1EJ7XJhqG9dniVW5jk38p96p8GDdEdQoGOGOTw7v7g'

fpDict = {}
rowDict = {}
profitDict = {}
profitList = []
lossList = []

def updateCurrentHoldings(service):
    req = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range="ProfitLog!D3:I1000")
    rows = req.execute()['values']

    ch = {}

    for row in rows:
        currentValue = row[0]
        sellingFee = row[2]
        if currentValue != '':
            currentValue = float(currentValue)
            if sellingFee != '':
                currentValue -= float(sellingFee)
            collectionName = row[5]
            if collectionName not in ch.keys():
                ch[collectionName] = [1, currentValue]
            else:
                ch[collectionName][1] += currentValue
                ch[collectionName][0] += 1

    t = [[x]+ch[x] for x in ch.keys()]
    t = sorted(t,key=lambda x: x[2],reverse=True)

    service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID,
                                          range="CurrentHoldings!A3:D1000").execute()
    
    print("Updating current holdings")
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'valueInputOption': "USER_ENTERED",
              'data': [{'range': "CurrentHoldings!A3:D" + str(3 +len(t)),
                        'values': t}]}).execute()


def updateFloorPrice(service, rowData, rowNum):
    colName = rowData[7]
    print(colName)
    if colName in fpDict:
        service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        valueInputOption="USER_ENTERED",
        range="ProfitLog!D" + str(rowNum),
        body={"values":[[fpDict[colName]]]}).execute()
        return
    r = requests.get('https://api.opensea.io/api/v1/collection/' + colName + '/stats')
    x = r.json()
   
    fp = x['stats']['floor_price']
    fpDict[colName] = fp
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        valueInputOption="USER_ENTERED",
        range="ProfitLog!D" + str(rowNum),
        body={"values":[[fp]]}).execute()

def updateProfitTable(service):
    global profitList
    global lossList
    global rowDict

    
    for p in range(len(profitList)):
        profitList[p][1] = "=SUM("
        for i in rowDict[profitList[p][0]]:
            profitList[p][1] += "G" + str(i) + ","
        profitList[p][1] += ")"

    for p in range(len(lossList)):
        lossList[p][1] = "=SUM("
        for i in rowDict[lossList[p][0]]:
            lossList[p][1] += "G" + str(i) + ","
        lossList[p][1] += ")"

    service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID,
                                          range="ProfitLog!L8:O1000").execute()
   
    print("Updating profit and loss tables")
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'valueInputOption': "USER_ENTERED",
              'data': [{'range': "ProfitLog!L8:M" + str(7 +len(profitList)),
                        'values': profitList}]}).execute()
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'valueInputOption': "USER_ENTERED",
              'data': [{'range': "ProfitLog!N8:O" + str(7 +len(lossList)),
                        'values': lossList}]}).execute()

def updateConditionalFormatting(service):
    global profitList
   
    request = [
      {
        "addConditionalFormatRule": {
          "rule": {
            "gradientRule": {
              "maxpoint": {
                "type": "MAX",
                "color": {'red':87/255,'green':187/255,'blue':138/255}
              },
              "minpoint": {
                "type": "MIN",
                "color":{'red':1,'green':1,'blue':1}
              }
            },
            "ranges": [{'sheetId':0,
                       'startRowIndex':7,
                       'endRowIndex':7+len(profitList),
                       'startColumnIndex':12,
                       'endColumnIndex':13}]
                      
        },
        "index": 0
        }
      }, {
        "addConditionalFormatRule": {
          "rule": {
            "gradientRule": {
              "minpoint": {
                "type": "MIN",
                "color": {'red':230/255,'green':124/255,'blue':115/255}
              },
              "maxpoint": {
                "type": "MAX",
                "color":{'red':1,'green':1,'blue':1}
              }
            },
            "ranges": [{'sheetId':0,
                       'startRowIndex':7,
                       'endRowIndex':7+len(lossList),
                       'startColumnIndex':14,
                       'endColumnIndex':15}]
                      
        },
        "index": 0
        }
      }, {
      "repeatCell": {
          "range": {
              "sheetId": 0,
              "startRowIndex": 7,
              "endRowIndex": 7+len(lossList),
              "startColumnIndex": 13,
              "endColumnIndex": 14
              },
          "cell": {
              "userEnteredFormat": {
                  "backgroundColor": {
                      "red": 244/255,
                      "green": 204/255,
                      "blue": 204/255
                      },
                  "textFormat": {
                      "bold": True
                      }
                  }
              },
         
          "fields": "userEnteredFormat(textFormat, backgroundColor)"
          }
      }, {
      "repeatCell": {
          "range": {
              "sheetId": 0,
              "startRowIndex": 7,
              "endRowIndex": 7+len(profitList),
              "startColumnIndex": 11,
              "endColumnIndex": 12
              },
          "cell": {
              "userEnteredFormat": {
                  "backgroundColor": {
                      "red": 217/255,
                      "green": 234/255,
                      "blue": 211/255
                      },
                  "textFormat": {
                      "bold": True
                      }
                  }
              },
         
          "fields": "userEnteredFormat(textFormat, backgroundColor)"
          }
      }
     
     
    ]
    body = {'requests': request}
    service.spreadsheets() \
    .batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute() 

def Main():
    global profitList
    global profitDict
    global lossList
    global rowDict
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range='A3:I1000').execute()
        rows = result.get('values', [])

        for rowNum in range(len(rows)):
            transaction = rows[rowNum]
            collection = ""
            if len(transaction) >= 9:
                collection = transaction[8]
            profit = transaction[6]
            if rows[rowNum][4] == '' and rows[rowNum][7]!='':
                updateFloorPrice(service, transaction, rowNum+3)
                pass
            if collection == "":
                continue
            if collection not in profitDict.keys():
                profitDict[collection] = float(profit)
                rowDict[collection] = [rowNum+3]
            else:
                profitDict[collection] += float(profit)
                rowDict[collection].append(rowNum+3)
        profitDict = {k:v for k,v in sorted(profitDict.items(), key=lambda item: item[1], reverse=True)}
        profitList = [[k[0], k[1]] for k in profitDict.items()]
        for p in range(len(profitList)):
            if profitList[p][1] < 0:
                lossIndex = p
                break
           
        lossList = profitList[lossIndex:][::-1]
        profitList = profitList[:lossIndex]
        
        updateProfitTable(service)
        updateConditionalFormatting(service)
        updateCurrentHoldings(service)
               
           
    except HttpError as err:
        print(err)

if __name__ == "__main__":
    Main()
