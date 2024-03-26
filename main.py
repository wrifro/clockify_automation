import requests
import openpyxl

# API Info
API_KEY = 'ZDQ2YzA4ODktNWM4Ny00MGZhLWFhNGQtMzZhYWM2NjU5OGMx'
WORKSPACE_ID = '65a9019f3d6c0f48fa86a4f5'

BASE_URL = 'https://api.clockify.me/api/v1'

def read_projects_from_excel(file):
    # open file
    wb = openpyxl.load_workbook(file)
    sheet = wb['projects']
    projects = []
    for row in sheet.iter_rows(min_row = 2):
        name = str(row[0]) + ' - ' + str(row[1])
        entry = {'name': name}
        projects.append(entry)
    return projects

def create_project_in_clockify(project_data):
    url = f'{BASE_URL}/workspaces/{WORKSPACE_ID}/projects'
    headers = {'X-Api-Key': API_KEY}
    response = requests.post(url, headers=headers,json=project_data)
    print(response.status_code)
    response.raise_for_status # raise error for non-200 status codes

projects = read_projects_from_excel("C:\\Users\wright.frost\\OneDrive - V-Nova Services Ltd\\Clockify\\filtered_export_objectives_03-19-2024.xlsx")

for project in projects:
    create_project_in_clockify(project)

print("Projects created successfully in Clockify!")