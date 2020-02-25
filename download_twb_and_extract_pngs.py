"""
AUTHOR: jaowens
DESCRIPTION: Enter a project name and download pngs from Tableau.
"""
import os
import tableauserverclient as TSC
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ADD YOUR USER NAME AND PASSWORD HERE:
USER_UN = "rampvenu"
USER_PW = "Ram3@hcl2019"

AUTH_TABLEAU = TSC.TableauAuth(USER_UN, USER_PW, site_id='Compass')
SERVER_GTCBI = TSC.Server('https://gtcbi-stage.cisco.com')
SERVER_GTCBI.add_http_options({'verify': False})
SERVER_GTCBI.version = "2.6"

# Project ID for 'Compass Prime' folder:
TABLEAU_PROJECT_ID = "e159b352-0643-44ae-a298-004c097b1c1b"
TABLEAU_PROJECT_NAME = input("\nEnter a project name: ")
DIR_PNG_FILES = "./png_files"


def get_project_id():
    """ Docs """
    for this_project in TSC.Pager(SERVER_GTCBI.projects):
        if this_project.name == TABLEAU_PROJECT_NAME:
            print(this_project.name, this_project.id)
            return this_project.id


def get_project_workbooks(project_id):
    """ Docs """
    list_workbook_ids = []

    for this_workbook in TSC.Pager(SERVER_GTCBI.workbooks):

        if this_workbook.project_id == project_id:
            print(this_workbook.name, this_workbook.id, this_workbook.project_id)
            list_workbook_ids.append(this_workbook.id)

    return list_workbook_ids


def extract_png(this_project_name, this_workbook_id):
    """ Extract .png files for all the plots and save to ./png_files/<customer_folder>. """
    this_workbook = SERVER_GTCBI.workbooks.get_by_id(this_workbook_id)
    this_workbook_name = this_workbook.name
    this_image_folder = f"{DIR_PNG_FILES}/{this_project_name}"
    this_workbook_folder = f"{this_image_folder}/{this_workbook_name}"

    message_extract_png = (f"Extracting image files for {this_workbook_name}...")
    print(message_extract_png)

    try:
        os.mkdir(this_image_folder)
    except FileExistsError:
        pass

    try:
        os.mkdir(this_workbook_folder)
    except FileExistsError:
        pass

    for this_view in this_workbook.views:
        SERVER_GTCBI.views.populate_image(this_view)
        with open(f'{this_workbook_folder}/{this_view.name}.png', 'wb') as this_file:
            this_file.write(this_view.image)


def main():
    """ Docs """
    with SERVER_GTCBI.auth.sign_in(AUTH_TABLEAU):

        project_id = get_project_id()
        list_workbook_ids = get_project_workbooks(project_id)
        for workbook_id in list_workbook_ids:
            extract_png(TABLEAU_PROJECT_NAME, workbook_id)


main()
