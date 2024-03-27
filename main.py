import requests
import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# API Info
API_KEY = 'OGJmYzE1Y2ItZTk3ZS00M2E2LThhZGUtZWZjMDQ0NzE0M2I0'
WORKSPACE_ID = '65a9019f3d6c0f48fa86a4f5'
BASE_URL = 'https://api.clockify.me/api/v1'

class ClockifyConnect:
    def __init__(self, API_KEY, WORKSPACE_ID, BASE_URL):
        self.api = API_KEY
        self.workspace_id = WORKSPACE_ID
        self.base_url = BASE_URL

    # Method to take projects from a [properly formatted]
    # Excel file and transfer them into a lis  adst
    def read_projects_from_excel(self, file): 
        # open file
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        projects = []
        for row in sheet.iter_rows(min_row = 1):
            if row[6].value == 'PD1':
                name = str(row[1].value) + ' - ' + str(row[2].value)
                archived = True if row[6].value == 'PD0' else False
                entry = {'name': name,
                        'archived': archived}
                projects.append(entry)

        self.updated_projects = projects

    def get_projects(self):
        url = f'{self.base_url}/workspaces/{WORKSPACE_ID}/projects'
        response = requests.get(url, headers={'X-API-Key': self.api})
        self.data = response.json()
        
        if response.status_code != 200:
            print('request failure', response.status_code)

    def update_projects(self):
        url = f'{self.base_url}/workspaces/{WORKSPACE_ID}/projects'
        for project in self.updated_projects:
            response = requests.post(url, headers= {'X-Api-Key': self.api},json=project)
            if(response.status_code != 200):
                print(response.status_code,response.text)

if __name__ =='__main__':
    clockify = ClockifyConnect(API_KEY, WORKSPACE_ID,BASE_URL)
    Tk().withdraw()
    filename = askopenfilename()
    clockify.read_projects_from_excel(filename)
    # clockify.get_projects()
    clockify.update_projects()
