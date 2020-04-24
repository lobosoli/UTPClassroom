# ListarCursos.py
# Adrian Ramirez <adr@utp.edu.co>
# Available subject to the Apache 2.0 License
# https://www.apache.org/licenses/LICENSE-2.0

from __future__ import print_function

import pickle
import os.path
import json
import time
import sys
import openpyxl

from pprint import pprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError



# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/classroom.courses.readonly', 
        'https://www.googleapis.com/auth/classroom.rosters',
        'https://www.googleapis.com/auth/classroom.rosters.readonly',
        'https://www.googleapis.com/auth/classroom.profile.emails',
        'https://www.googleapis.com/auth/classroom.profile.photos',
	'https://www.googleapis.com/auth/drive', 
        'https://www.googleapis.com/auth/spreadsheets',
	'https://www.googleapis.com/auth/spreadsheets.readonly']

def check_auth(api, version):
    creds = None
    
    """Shows basic usage of the Classroom API.
    Prints the names of the first 10 courses the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'estadisticasCR_id.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build(api, version, credentials=creds)   
    return service

def main():        
    # Call the Classroom API
    classroom_service = check_auth('classroom','v1')  
    # Call Google Spreadsheets API
    SHEETS = check_auth('sheets','v4')
    # SpreadSheet properties
    spreadsheet_body = {
        'properties': {
            'title': 'ListadoCursos'
            
        },
        'sheets': [
            {'properties': {
                'sheetType': 'GRID',                
                'gridProperties': {
                    'rowCount': 20000, 
                    'columnCount': 9
                }
            }
            }
        ]
    }
    # Create a new SpreadSheet
    spreadsheet = SHEETS.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
    # Get Spreadsheet ID
    spreadsheets_Id = spreadsheet.get('spreadsheetId')
    #spreadsheets_Id = '1H1iNJlrbRG45KMG7EnacP6rlCj6LszJtgIm6bk3Zr9I'
    range_ = 'Sheet1!A1:J1'    
    values =[ 
        ['ID Curso'],
        ['Nombre Curso'],
        ['ID Profesor'],
        ['Nombre Profesor'],
        ['Correo Profesor'],
        ['Fecha de Creación'],
        ['Ultima Actualización'],
        ['Número de Estudintes'],
        ['Estado del Curso'],
    ]
    Body = {
        'values' : values,
        'majorDimension' : 'COLUMNS',
    }    
    # Update the spreadsheet
    result = SHEETS.spreadsheets().values().update(spreadsheetId = spreadsheets_Id, range = range_, valueInputOption = 'RAW', body = Body).execute()   
     
    # Number of rows   
    contar= 1

    courses = []
    page_token = None

    # Get all cousers
    while True:
        try:
            response = classroom_service.courses().list(pageToken=page_token, pageSize=500).execute()
            courses.extend(response.get('courses', []))
            page_token = response.get('nextPageToken', None)
            if not page_token:
                break
        except HttpError as err:
            # If the error is a rate limit or connection error, wait and try again.
            if err.resp.status in [400, 403, 500, 503]:
                time.sleep(5)
            else: raise

    if not courses:
        print('No courses found.')
    else:      
        #print('ID Curso;Nombre Curso;ID Profesor;Nombre Profesor;Correo Profesor;Fecha de Creación;Ultima Actualización;Número de Estudiantes;Estado del Curso')  
        for course in courses:            
            try: 
                teacherID = classroom_service.userProfiles().get(userId = course.get('ownerId'), x__xgafv='2', alt="json").execute()
                Estudiantes = classroom_service.courses().students().list(courseId = course['id'], pageSize = 0, x__xgafv='2', alt="json").execute()
                datosEstudiantes = Estudiantes.get('students',[])
                estuCont = 0
                contar=contar + 1
                #Get the numbers of students
                if not datosEstudiantes:
                    estuCont = 0
                else:
                    for listaEstudiantes in datosEstudiantes:                            
                        estuCont = estuCont + 1
                #print(course['id'],';', course['name'],';', course['ownerId'],';', teacherID['name']['fullName'],';', teacherID['emailAddress'],';', course['creationTime'],';', course['updateTime'],';', estuCont ,';', course['courseState'])                                                  
                values =[ 
                        [course['id']],
                        [course['name']],
                        [course['ownerId']],
                        [teacherID['name']['fullName']],
                        [teacherID['emailAddress']],
                        [course['creationTime']],
                        [course['updateTime']],
                        [estuCont],
                        [course['courseState']],
                ]
                Body = {
                    'values' : values,
                    'majorDimension' : 'COLUMNS',
                }
                range_ = 'Sheet1!A' + str(contar) + ':J' + str(contar)
                result = SHEETS.spreadsheets().values().update(spreadsheetId = spreadsheets_Id, range = range_,valueInputOption = 'RAW', body = Body).execute() 
            except HttpError as err:
                # If the error is a rate limit or connection error, wait and try again.
                if err.resp.status in [400, 403, 500, 503]:
                    time.sleep(5)
                else: raise      

if __name__ == '__main__':
    main()
