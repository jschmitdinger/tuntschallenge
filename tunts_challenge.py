#!/usr/bin/env python
"""
This program was written to be submitted to the Tunts Challenge, one of the
stages of the TuntsCorp Internship selection process. The program connects to
an online worksheet and reads data from it. The data are fictional grades and
absence rates from a fictional classroom. The program must then calculate the
final situation of each student based on predetermined criteria and write the
results on the online worksheet, thus completing the grade form.
"""

# Importing required Sheets API funtions to connect with an online sheet
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# Importing required mathematical functions from the math.py library
from math import ceil

__author__ = "Johann Schmitdinger Vieira"
__version__ = "1.0"
__maintainer__ = "Johann Schmitdinger Vieira"
__email__ = "johann.schmitdinger@outlook.com"

# Authentication scope from Sheets API for read and write
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Worksheet ID and cell Ranges for reading and writing
worksheetID = '1qTotMIq_6HZ_lSXCl7ABaU-mZxWFaeLCXllWGAuhwVA'
rangeRead = 'engenharia_de_software!A4:F'
rangeWrite = 'engenharia_de_software!G4:H'

# This function receives the students data in matrix format and generates
# another matrix containing the resulting "Situation" and final exame grade
def situation(studentData):

    results = [[], []] # results will be saved in this matrix
    gradeAverage = 0

    for row in studentData:
        gradeAverage = (int(row[3]) + int(row[4]) + int(row[5]))/30
        if int(row[2]) > 15: # evaluating the number of absences
            results[0].append("Reprovado por Falta")
            results[1].append(0)
        elif gradeAverage < 5: # evaluating grade averages < 5
            results[0].append("Reprovado por Nota")
            results[1].append(0)
        elif gradeAverage < 7: # evaluating 5 =< grade averages and < 7
            results[0].append("Exame Final")
            results[1].append(ceil(10-gradeAverage))
        else: # evaluating grade averages > 7
            results[0].append("Aprovado")
            results[1].append(0)

    print("Results were succesfully calculated.")
    return results

# The main function of this program, establishes connection with the worksheet
# through Sheets API, reads the ranges containing the students data and writes
# the results on the respective range
def main():

    print("Initializing application.")
    creds = None

    # The token.pickle file stores the authentication data and directs the user
    # to the login page in case no previous data is located.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
        print("Token was located and login data was extracted.")
    # If no token file is found, the user must login.
    if not creds or not creds.valid:
        print("No Token was found and the user must login to continue.")
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Token credentials are saved for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
        print("Token updated and saved.")

    # Runs Sheets API to establish a connection with the online worksheet
    print("Attempting to establish connection with online worksheet.")
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Result reads worksheet cells with the .get function and trasnfers them
    # to "values"
    result = sheet.values().get(spreadsheetId=worksheetID,
                                range=rangeRead).execute()
    values = result.get('values', [])

    # Testing for blank data
    if not values:
        print("No data was found.")
    else:
        print("Values succesfully read from the online worksheet.")

        # Running situation function to determine the students situation
        results = situation(values)

        # Formatting the results in the format as required by Sheets API
        body = {
            'values' : results,
            'majorDimension' : 'COLUMNS'
        }

        # Updating the online sheet with generated data.
        result = service.spreadsheets().values().update(
            spreadsheetId=worksheetID, range=rangeWrite,
            valueInputOption='RAW', body=body).execute()

        print('{0} cells were updated in the online worksheet.'.format(
            result.get('updatedCells')))
        print("Exiting program.")

# Runs the main() function
if __name__ == '__main__':
    main()
