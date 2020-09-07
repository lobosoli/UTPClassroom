from __future__ import print_function

import pickle
import os.path
import json
import time
import sys
import openpyxl

from datetime import date
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
	    'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/admin.reports.usage.readonly']

def check_auth(user, api):
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

    service = build(user, api, credentials=creds)   
    return service

def profesoresActivos(fecha):        
    # Call the Reports API
    reports_service = check_auth('admin','reports_v1')  
    num_7day_teachers = reports_service.customerUsageReports().get(date = fecha ,parameters = 'classroom:num_7day_teachers', x__xgafv='2', alt="json").execute()    
        
    teachers = num_7day_teachers.get('usageReports', [])

    for x in teachers:
        lista = x['parameters']
        for y in lista:
            return int(y['intValue'])              

def estudiantesActivos(fecha):        
    # Call the Reports API
    reports_service = check_auth('admin','reports_v1')  
    num_7day_students = reports_service.customerUsageReports().get(date = fecha ,parameters = 'classroom:num_7day_students', x__xgafv='2', alt="json").execute()
        
    students = num_7day_students.get('usageReports', [])        

    for x in students:
        lista = x['parameters']
        for y in lista:
            return int(y['intValue'])

def cursosCreados(fecha_):        
    # Call the Reports API
    reports_service = check_auth('admin','reports_v1')  
    courses_created = reports_service.customerUsageReports().get(date = str(fecha_) ,parameters = 'classroom:num_courses_created', x__xgafv='2', alt="json").execute()    
    courses = courses_created.get('usageReports', [])
    for x in courses:
        lista = x['parameters']
        for y in lista:
            return int(y['intValue'])

def cursosActivos(fecha1):        
    # Call the Reports API
    reports_service = check_auth('admin','reports_v1')  
    num_14day = reports_service.customerUsageReports().get(date = fecha1 ,parameters = 'classroom:num_14day_active_courses', x__xgafv='2', alt="json").execute()
    
    active_courses = num_14day.get('usageReports', []) 

    for x in active_courses:
        lista = x['parameters']
        for y in lista:
            return int(y['intValue'])

def listarCursos(fecha1_, fecha2_):        
    # Call the Classroom API
    classroom_service = check_auth('classroom','v1')  
    # Call Google Spreadsheets API
    SHEETS = check_auth('sheets','v4')
    drive_service = check_auth('drive','v3')

    # Get date
    today = str(date.today())
    NombreHC = "ReporteClassroom-"+today

    # SpreadSheet properties
    spreadsheet_body = {
        'properties': {
            'title': NombreHC            
        },
        'sheets': [
            {
                'properties': 
                    {
                        'title': "CursosActivos",
                        'sheetType': 'GRID',                
                        'gridProperties': 
                        {
                            'rowCount': 20000, 
                            'columnCount': 9
                        }
                }
            },
            {
                'properties': 
                    {
                        'title': "InfoConsolidada",
                        'sheetType': 'GRID',                
                        'gridProperties': 
                        {
                            'rowCount': 6, 
                            'columnCount': 2
                        }
                    }
            }
        ]
    }

    # Create a new SpreadSheet
    spreadsheet = SHEETS.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
    # Get Spreadsheet ID
    spreadsheets_Id = spreadsheet.get('spreadsheetId')    
    # Move the file to the new folder
    folder_id = '1nPFQPf3WOYMqmHxjrGQDAYaCvg-SDt3K'
    res = drive_service.files().update(fileId = spreadsheet.get('spreadsheetId'), addParents = folder_id, removeParents = 'root').execute()       

    # Set Columns Names
    range_ = 'CursosActivos!A1:J1'    
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
    
    # Set Rows names
    range_ = 'InfoConsolidada!A1:A6'
    values =[ 
        ['Cursos Creados desde 2020-08-18'],        
        ['Curso Activos Últimos 14 días'],
        ['Cursos con 0 Estudiantes'],
        ['Profesores Activos Últimos 14 días'],
        ['Estudiantes Activos Últimos 14 días'],        
    ]
    Body = {
        'majorDimension' : 'ROWS',
        'values' : values,
    }    

    # Update the spreadsheet
    result = SHEETS.spreadsheets().values().update(spreadsheetId = spreadsheets_Id, range = range_, valueInputOption = 'RAW', body = Body).execute()

    # Generate consolidate data
    cursos_ = cursosCreados('2020-08-18')    
    activos_ = cursosActivos( fecha1_ )
    profesores_ = profesoresActivos( fecha1_ ) + profesoresActivos( fecha2_ )
    estudiantes_ = estudiantesActivos( fecha1_ ) + estudiantesActivos( fecha2_ )    

    range_ = 'InfoConsolidada!B1:B6'
    values =[ 
        [cursos_],        
        [activos_],
        ["=countif(CursosActivos!H2:H20000,\"=0\")"],
        [profesores_],
        [estudiantes_],        
    ]
    Body = {        
        'majorDimension' : 'ROWS',
        'values' : values,
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
        for course in courses:            
            try:
                
                teacherID = classroom_service.userProfiles().get(userId = course.get('ownerId'), x__xgafv='2', alt="json").execute()
                Estudiantes = classroom_service.courses().students().list(courseId = course['id'], pageSize = 0, x__xgafv='2', alt="json").execute()
                datosEstudiantes = Estudiantes.get('students',[])
                estuCont = 0
                contar = contar + 1
                #Get the numbers of students
                if not datosEstudiantes:
                    estuCont = 0
                else:
                    for listaEstudiantes in datosEstudiantes:                            
                        estuCont = estuCont + 1                    
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
                range_ = 'CursosActivos!A' + str(contar) + ':J' + str(contar)
                result = SHEETS.spreadsheets().values().update(spreadsheetId = spreadsheets_Id, range = range_,valueInputOption = 'RAW', body = Body).execute() 
            except HttpError as err:
                # If the error is a rate limit or connection error, wait and try again.
                if err.resp.status in [400, 403, 500, 503]:
                    time.sleep(500)
                else: raise      
                  

def main():
    print("Por favor digite fecha para informe 14 días (YYYY-MM-DD): ")
    fecha1_ = input()
    print("Por favor digite fecha para informe 7 días (YYYY-MM-DD): ")
    fecha2_ = input()
    listarCursos(fecha1_ ,fecha2_ )  

if __name__ == '__main__':    
    main()
