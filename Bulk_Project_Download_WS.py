import requests
import re
import io
import zipfile
import os
import xlwings
import sys
import tkinter
import datetime
from tkinter import messagebox
from tkinter import simpledialog
from xlwings import books
from pywintypes import *

checkbox_value = []
project_name_list = []
comment_list = []
json_list = []
project_data = {}
project_scoping = {}
cost_list = []
scoping_json_list = []
proj_download = []
proj_missing = set()
AID_data = {}
missing_proj_dict = {}
project_name_list = []
deadline_dict = {}
invalid_proj = []
file_deadline_dict = {}
missing_link = []
wwbfile_rgx = re.compile('input name="wwbfile"  value="' +
                         '[a-z0-9-/\\\.]+', re.I)

'''Each freelancer receives a group of projects called "assignment"
and each of those have an ID. This function downloads all
projects per ID and checks that the name of the projects
downloaded and the ones in the Excel sheet are the same'''
try:
    def download_WS_project(token,
                            openertoken,
                            checkbox_proj,
                            scheme,
                            assig_ID,
                            proj_name_list):
        params_1 = {'scheme': scheme,
                    'url_of_origin': 'https://example-worldserver.com',
                    'tm_export_type': 'CONTENT',
                    'backwardCompatibleKit': 'on',
                    'tmFilterId': '-1',
                    'tdFilterId': '-1',
                    'checkbox': checkbox_proj,
                    'project': '0',
                    'cur_step': 'assetExportStep2',
                    'mode': 'projects',
                    'formAction': '"export?&token=' +
                    str(token) +
                    '&mode=projects&project=0&openertoken=' +
                    str(openertoken) + '"',
                    'submittedBy': 'ok', 'methodUsed': 'POST'}
        r = requests.post(
            'http://example-worldserver.com/ws/export?&token=' + str(token) +
            '&mode=projects&project=0&openertoken=' +
            str(openertoken), verify=False, data=params_1)
        first_response = r.text
        find_wwbfile = re.search(wwbfile_rgx, first_response)
        try:
            wwbfile_value = find_wwbfile.group(0)
        except AttributeError:
            missing_proj_dict[assig_ID] = proj_name_list
        else:
            wwbfile_value = wwbfile_value.split('value="')[1]
            params_2 = {'download': 'yes', 'wwbfile': wwbfile_value,
                        'suggestedname': 'xliff_projects.zip'}
            headers = {'Host': 'example-worldserver.com',
                       'User-Agent': 'Mozilla/5.0 (Windows NT 6.1;\
                       WOW64; rv:45.0) Gecko/20100101 Firefox/45.0',
                       'Accept': 'text/html,application/xhtml+xml,\
                       application/xml;q=0.9,*/*;q=0.8',
                       'Accept-Language': 'en-US,en;q=0.5',
                       'Accept-Encoding': 'gzip, deflate',
                       'Referer': 'https://example-worldserver.com/\
                       export?&token=' +
                       str(token) + '&mode=projects&project=0&openertoken=' +
                       str(openertoken), 'Connection': 'keep-alive'}
            r_2 = requests.post(
                'http://example-worldserver.com/export?&token=' +
                str(token) +
                '&mode=projects&project=0&openertoken=' +
                str(openertoken), verify=False, data=params_2, headers=headers)
            fp = io.BytesIO(r_2.content)
            print("project downloaded")
            zfp = zipfile.ZipFile(fp, "r")
            userpro = os.environ['USERPROFILE']
            destDir = str(userpro) + '\\Downloads'
            os.chdir(destDir)
            z = zipfile.ZipFile('Assignment ID ' +
                                str(assig_ID) + ".zip", "w")
            zfp.extractall(destDir)
            file_names = zfp.namelist()
            project_names = proj_name_list
            # test if all projects were downloaded
            for name in project_names:
                for proj in file_names:
                    if name in proj:
                        proj_download.append(name)
                        deadline_str = deadline_dict.get(name)
                        file_deadline_dict[proj] = deadline_str
            proj_missing = set(project_names).difference(set(proj_download))
            for item in file_names:
                file_deadline = file_deadline_dict.get(item)
                os.rename(str(destDir) + "\\" + str(item), str(destDir) +
                          "\\" + str(file_deadline) + "_" + str(item))
                z.write(str(file_deadline) + "_" + str(item))
                os.remove(str(destDir) + "\\" +
                          str(file_deadline) + "_" + str(item))
            z.close()
            if len(proj_missing) != 0:
                missing_proj_dict[assig_ID] = proj_missing

    def check_links(unique_AID, get_AID_table):
        for index, value in enumerate(get_AID_table):
            if value == unique_AID:
                indeces_set.add(int(index) + 2)
        for i in indeces_set:
            try:
                link = xlwings.Range((i, wsproj_pos)).hyperlink
            except Exception:
                missing_link.append(str(i))
            else:
                project_name = str(xlwings.Range((i, wsproj_pos)).value)
                project_id_search = re.search(project_id_rgx, link)
                try:
                    project_id = project_id_search.group(1)
                except AttributeError:
                    missing_link.append(str(i))
        indeces_set.clear()

    ''' In the Excel sheet with the project information,
    iterate through the assignment ID column to get all
    the information of all the projects with that ID'''
    def get_info_for_AID(unique_AID, get_AID_table):
        for index, value in enumerate(get_AID_table):
            if value == unique_AID:
                indeces_set.add(int(index) + 2)
        # get project names, WS id of projects,
        # comments and scoping
        for i in indeces_set:
            try:
                link = xlwings.Range((i, wsproj_pos)).hyperlink
            except Exception:
                missing_link.append(str(i))
            else:
                project_name = str(xlwings.Range((i, wsproj_pos)).value)
                project_id_search = re.search(project_id_rgx, link)
                try:
                    project_id = project_id_search.group(1)
                except AttributeError:
                    missing_link.append(str(i))
                else:
                    project_name = project_name.replace('\xa0', '')
                    project_deadline = datetime.datetime.strptime(
                        str(xlwings.Range(
                            (i, projdeadline_pos)).value),
                        "%Y-%m-%d %H:%M:%S").strftime(
                        '%d-%m-%Y %H:%M:%S').split(" ")[0]
                    project_name_list.append(project_name)
                    deadline_dict[project_name] = project_deadline
                    checkbox_value.append(project_id)
                    comment_list.append(
                        str(xlwings.Range((i, comment_pos)).value))
                    freelancer_name = xlwings.Range((i, freelance_pos)).value
        try:
            freelance_PO_pos = email_list.index(freelancer_name) + 2
        except UnboundLocalError:
            indeces_set.clear()
        else:
            projects_list = project_name_list[:]
            checkbox_list = checkbox_value[:]
            software_value = str(wb.sheets['PO'].range(
                (freelance_PO_pos, software_pos)).value).lower()
            if "trados" in software_value or (
                    "memoq" in software_value or "memo q" in software_value):
                scheme_value = 'trados_studio'
            else:
                scheme_value = 'xliff'
            AID_data[unique_AID] = [checkbox_list, scheme_value, projects_list]
            del project_name_list[:]
            del checkbox_value[:]
            del comment_list[:]
            indeces_set.clear()

        '''Check if Excel is active and if the correct
        Excel file is open/no one changed the column names'''

    try:
        wb = books.active
    except (IndexError, OSError):
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="Excel is not active. \
            Please open the Freelancer Tracker you want to use")
        sys.exit()

    try:
        headers = xlwings.Range('A1').expand('right').value
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="This file does not contain a 'Tracker' tab. \
            Sure this is the Freelance tracker?")
        sys.exit()

    try:
        AID_pos = headers.index('Assignment ID') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Assignment ID' in this file.\
            Sure this is the tracker?")
        sys.exit()

    try:
        wsproj_pos = headers.index('WorldServer project') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'WorldServer Project' in this file. \
            Sure this is the tracker?")
        sys.exit()

    try:
        projdeadline_pos = headers.index('Due date') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Due date' in this file. \
            Sure this is the tracker?")
        sys.exit()

    try:
        comment_pos = headers.index('Comment') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Comment' in this file. \
            Sure this is the tracker?")
        sys.exit()

    try:
        freelance_pos = headers.index('Name') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Name' in this file. \
            Sure this is the tracker?")
        sys.exit()
    try:
        headers_PO = wb.sheets['PO'].range('A1').expand('right').value
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="This file does not contain a 'PO' tab. \
            Sure this is the Freelance tracker?")
        sys.exit()

    try:
        software_pos = headers_PO.index('Software') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Software' in this file. \
            Sure this is the tracker?")
        sys.exit()
    try:
        name_pos = headers_PO.index('Name') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Name' in this file. \
            Sure this is the tracker?")
        sys.exit()
    try:
        email_pos = headers_PO.index('email') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'email' in this file. \
            Sure this is the tracker?")
        sys.exit()

    '''The user selects a list of projects in Excel
    We get the address of those rows in order to gather
    all the information'''

    freelancer_pos = wb.selection.address

    if ":" in freelancer_pos:
        first_row = freelancer_pos.split("$")[2][:-1]
        last_row = freelancer_pos.split("$")[4]
        current_AID = xlwings.Range(
            (first_row, AID_pos), (last_row, AID_pos)).value
    else:
        freelancer_row = freelancer_pos.split("$")[2]
        current_AID = [xlwings.Range((freelancer_row, AID_pos)).value]

    unique_AID = set(current_AID)
    AID_table = xlwings.Range((2, AID_pos)).expand('down').value
    indeces_set = set()
    project_id_rgx = re.compile('project=' + '([0-9]+)' + '(&)?', re.I)
    email_list = wb.sheets['PO'].range((2, name_pos)).expand('down').value

    for AID in unique_AID:
        check_links(AID, AID_table)

    if len(missing_link) != 0:
        missing_link.sort()
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="Missing or wrong link in the \
            following row(s): " + ", ".join(missing_link))
        sys.exit()

    for AID in unique_AID:
        get_info_for_AID(AID, AID_table)
    AID_data_values = list(AID_data.values())
    sample_proj_id = AID_data_values[0][0][0]

    print("Information gathered from Excel")

    ''' download the projects. For that we need the
    user token'''

    root = tkinter.Tk()
    root.withdraw()
    token = simpledialog.askstring(
        'Token required', 'Please enter the number after "token="')

    if token is None or len(token) == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title='Error', message="You didn't introduce a token, \
            so farewell!")
        sys.exit()
    test = token.isdigit()
    if test is False:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title='Error', message='The token you pasted is incorrect. \
            Please paste only the number after "token=" in the address \
            bar of your browser.')
        sys.exit()

    print("Getting openertoken")
    get_openertoken = requests.get('http://example-worldserver.com/ws/\
    assignments_tasks?&project=' +
                                   str(sample_proj_id) +
                                   '&token=' + str(token_value), verify=False)
    r_foropener = get_openertoken.text
    opener_rgx = re.compile('openertoken=' + '([0-9]+)' + '(&)?', re.I)
    find_opener = re.search(opener_rgx, r_foropener)
    try:
        openertoken_value = find_opener.group(1)
    except AttributeError:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="Please refresh the WS tab, \
            the token is outdated")
        sys.exit()

    number_of_AIDs = len(AID_data)
    counter = 0
    print("Starting download, this might take a few minutes.")
    for key, value in AID_data.items():
        counter += 1
        download_WS_project(token_value, openertoken_value,
                            value[0], value[1], key, value[2])
        print("Downloaded Assignment " + str(counter) +
              " out of " + str(number_of_AIDs))

    userpro = os.environ['USERPROFILE']

    if len(missing_proj_dict) != 0:
        report_file = open(userpro + '\\Downloads\\Download_report.txt', 'w')
        report_file.write("Assignment ID    Missing projects\n")
        for k, v in missing_proj_dict.items():
            report_file.write(
                str(k) + "  " + " ,".join([value for value in v]) + "\n")
        report_file.close()

    missing_link.sort()

    '''Once the download is finished, we inform of the result'''

    if len(missing_link) != 0 and len(missing_proj_dict) == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(
            title="Success!", message="Download is finished. There was no hyperlink or \
            an invalid one in the following row(s): " +
            ", ".join(missing_link))
        sys.exit()
    elif len(missing_link) == 0 and len(missing_proj_dict) != 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(
            title="Success!", message="Download is finished. \
            Some projects could not be downloaded, see \
            'Download_report.txt' in the Downloads folder")
        sys.exit()
    elif len(missing_link) != 0 and len(missing_proj_dict) != 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(title="Success!", message="Download is finished.\
        Some projects could not be downloaded (see 'Download_report.txt') \
        and there was no hyperlink or an invalid one in the following \
        row(s): " + ", ".join(missing_link))
        sys.exit()
    elif len(missing_link) == 0 and len(missing_proj_dict) == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(
            title="Success!", message="Download is finished, \
            all projects were downloaded and there are no missing hyperlinks.")
        sys.exit()
except Exception as e:
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showerror(title="Error", message=str(e))
    sys.exit()
=======
import requests
import re
import io
import zipfile
import os
import xlwings
import sys
import tkinter
import datetime
from tkinter import messagebox
from tkinter import simpledialog
from xlwings import books
from pywintypes import *

checkbox_value = []
project_name_list = []
comment_list = []
json_list = []
project_data = {}
project_scoping = {}
cost_list = []
scoping_json_list = []
proj_download = []
proj_missing = set()
AID_data = {}
missing_proj_dict = {}
project_name_list = []
deadline_dict = {}
invalid_proj = []
file_deadline_dict = {}
missing_link = []
wwbfile_rgx = re.compile('input name="wwbfile"  value="' +
                         '[a-z0-9-/\\\.]+', re.I)

'''Each freelancer receives a group of projects called "assignment"
and each of those have an ID. This function downloads all
projects per ID and checks that the name of the projects
downloaded and the ones in the Excel sheet are the same'''
try:
    def download_WS_project(token,
                            openertoken,
                            checkbox_proj,
                            scheme,
                            assig_ID,
                            proj_name_list):
        params_1 = {'scheme': scheme,
                    'url_of_origin': 'https://example-worldserver.com',
                    'tm_export_type': 'CONTENT',
                    'backwardCompatibleKit': 'on',
                    'tmFilterId': '-1',
                    'tdFilterId': '-1',
                    'checkbox': checkbox_proj,
                    'project': '0',
                    'cur_step': 'assetExportStep2',
                    'mode': 'projects',
                    'formAction': '"export?&token=' +
                    str(token) +
                    '&mode=projects&project=0&openertoken=' +
                    str(openertoken) + '"',
                    'submittedBy': 'ok', 'methodUsed': 'POST'}
        r = requests.post(
            'http://example-worldserver.com/ws/export?&token=' + str(token) +
            '&mode=projects&project=0&openertoken=' +
            str(openertoken), verify=False, data=params_1)
        first_response = r.text
        find_wwbfile = re.search(wwbfile_rgx, first_response)
        try:
            wwbfile_value = find_wwbfile.group(0)
        except AttributeError:
            missing_proj_dict[assig_ID] = proj_name_list
        else:
            wwbfile_value = wwbfile_value.split('value="')[1]
            params_2 = {'download': 'yes', 'wwbfile': wwbfile_value,
                        'suggestedname': 'xliff_projects.zip'}
            headers = {'Host': 'example-worldserver.com',
                       'User-Agent': 'Mozilla/5.0 (Windows NT 6.1;\
                       WOW64; rv:45.0) Gecko/20100101 Firefox/45.0',
                       'Accept': 'text/html,application/xhtml+xml,\
                       application/xml;q=0.9,*/*;q=0.8',
                       'Accept-Language': 'en-US,en;q=0.5',
                       'Accept-Encoding': 'gzip, deflate',
                       'Referer': 'https://example-worldserver.com\
                       /export?&token=' +
                       str(token) + '&mode=projects&project=0&openertoken=' +
                       str(openertoken), 'Connection': 'keep-alive'}
            r_2 = requests.post(
                'http://example-worldserver.com/export?&token=' +
                str(token) +
                '&mode=projects&project=0&openertoken=' +
                str(openertoken), verify=False, data=params_2, headers=headers)
            fp = io.BytesIO(r_2.content)
            print("project downloaded")
            zfp = zipfile.ZipFile(fp, "r")
            userpro = os.environ['USERPROFILE']
            destDir = str(userpro) + '\\Downloads'
            os.chdir(destDir)
            z = zipfile.ZipFile('Assignment ID ' +
                                str(assig_ID) + ".zip", "w")
            zfp.extractall(destDir)
            file_names = zfp.namelist()
            project_names = proj_name_list
            # test if all projects were downloaded
            for name in project_names:
                for proj in file_names:
                    if name in proj:
                        proj_download.append(name)
                        deadline_str = deadline_dict.get(name)
                        file_deadline_dict[proj] = deadline_str
            proj_missing = set(project_names).difference(set(proj_download))
            for item in file_names:
                file_deadline = file_deadline_dict.get(item)
                os.rename(str(destDir) + "\\" + str(item), str(destDir) +
                          "\\" + str(file_deadline) + "_" + str(item))
                z.write(str(file_deadline) + "_" + str(item))
                os.remove(str(destDir) + "\\" +
                          str(file_deadline) + "_" + str(item))
            z.close()
            if len(proj_missing) != 0:
                missing_proj_dict[assig_ID] = proj_missing

    def check_links(unique_AID, get_AID_table):
        for index, value in enumerate(get_AID_table):
            if value == unique_AID:
                indeces_set.add(int(index) + 2)
        for i in indeces_set:
            try:
                link = xlwings.Range((i, wsproj_pos)).hyperlink
            except Exception:
                missing_link.append(str(i))
            else:
                project_name = str(xlwings.Range((i, wsproj_pos)).value)
                project_id_search = re.search(project_id_rgx, link)
                try:
                    project_id = project_id_search.group(1)
                except AttributeError:
                    missing_link.append(str(i))
        indeces_set.clear()

    ''' In the Excel sheet with the project information,
    iterate through the assignment ID column to get all
    the information of all the projects with that ID'''
    def get_info_for_AID(unique_AID, get_AID_table):
        for index, value in enumerate(get_AID_table):
            if value == unique_AID:
                indeces_set.add(int(index) + 2)
        # get project names, WS id of projects,
        # comments and scoping
        for i in indeces_set:
            try:
                link = xlwings.Range((i, wsproj_pos)).hyperlink
            except Exception:
                missing_link.append(str(i))
            else:
                project_name = str(xlwings.Range((i, wsproj_pos)).value)
                project_id_search = re.search(project_id_rgx, link)
                try:
                    project_id = project_id_search.group(1)
                except AttributeError:
                    missing_link.append(str(i))
                else:
                    project_name = project_name.replace('\xa0', '')
                    project_deadline = datetime.datetime.strptime(
                        str(xlwings.Range(
                            (i, projdeadline_pos)).value),
                        "%Y-%m-%d %H:%M:%S").strftime(
                        '%d-%m-%Y %H:%M:%S').split(" ")[0]
                    project_name_list.append(project_name)
                    deadline_dict[project_name] = project_deadline
                    checkbox_value.append(project_id)
                    comment_list.append(
                        str(xlwings.Range((i, comment_pos)).value))
                    freelancer_name = xlwings.Range((i, freelance_pos)).value
        try:
            freelance_PO_pos = email_list.index(freelancer_name) + 2
        except UnboundLocalError:
            indeces_set.clear()
        else:
            projects_list = project_name_list[:]
            checkbox_list = checkbox_value[:]
            software_value = str(wb.sheets['PO'].range(
                (freelance_PO_pos, software_pos)).value).lower()
            if "trados" in software_value or (
                    "memoq" in software_value or "memo q" in software_value):
                scheme_value = 'trados_studio'
            else:
                scheme_value = 'xliff'
            AID_data[unique_AID] = [checkbox_list, scheme_value, projects_list]
            del project_name_list[:]
            del checkbox_value[:]
            del comment_list[:]
            indeces_set.clear()

        '''Check if Excel is active and if the correct
        Excel file is open/no one changed the column names'''

    try:
        wb = books.active
    except (IndexError, OSError):
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="Excel is not active. \
            Please open the Freelancer Tracker you want to use")
        sys.exit()

    try:
        headers = xlwings.Range('A1').expand('right').value
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="This file does not contain a 'Tracker' tab. \
            Sure this is the Freelance tracker?")
        sys.exit()

    try:
        AID_pos = headers.index('Assignment ID') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Assignment ID' in this file.\
            Sure this is the tracker?")
        sys.exit()

    try:
        wsproj_pos = headers.index('WorldServer project') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'WorldServer Project' in this file. \
            Sure this is the tracker?")
        sys.exit()

    try:
        projdeadline_pos = headers.index('Due date') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Due date' in this file. \
            Sure this is the tracker?")
        sys.exit()

    try:
        comment_pos = headers.index('Comment') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Comment' in this file. \
            Sure this is the tracker?")
        sys.exit()

    try:
        freelance_pos = headers.index('Name') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Name' in this file. \
            Sure this is the tracker?")
        sys.exit()
    try:
        headers_PO = wb.sheets['PO'].range('A1').expand('right').value
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="This file does not contain a 'PO' tab. \
            Sure this is the Freelance tracker?")
        sys.exit()

    try:
        software_pos = headers_PO.index('Software') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Software' in this file. \
            Sure this is the tracker?")
        sys.exit()
    try:
        name_pos = headers_PO.index('Name') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'Name' in this file. \
            Sure this is the tracker?")
        sys.exit()
    try:
        email_pos = headers_PO.index('email') + 1
    except Exception:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="There is not a header called 'email' in this file. \
            Sure this is the tracker?")
        sys.exit()

    '''The user selects a list of projects in Excel
    We get the address of those rows in order to gather
    all the information'''

    freelancer_pos = wb.selection.address

    if ":" in freelancer_pos:
        first_row = freelancer_pos.split("$")[2][:-1]
        last_row = freelancer_pos.split("$")[4]
        current_AID = xlwings.Range(
            (first_row, AID_pos), (last_row, AID_pos)).value
    else:
        freelancer_row = freelancer_pos.split("$")[2]
        current_AID = [xlwings.Range((freelancer_row, AID_pos)).value]

    unique_AID = set(current_AID)
    AID_table = xlwings.Range((2, AID_pos)).expand('down').value
    indeces_set = set()
    project_id_rgx = re.compile('project=' + '([0-9]+)' + '(&)?', re.I)
    email_list = wb.sheets['PO'].range((2, name_pos)).expand('down').value

    for AID in unique_AID:
        check_links(AID, AID_table)

    if len(missing_link) != 0:
        missing_link.sort()
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="Missing or wrong link in the \
            following row(s): " + ", ".join(missing_link))
        sys.exit()

    for AID in unique_AID:
        get_info_for_AID(AID, AID_table)
    AID_data_values = list(AID_data.values())
    sample_proj_id = AID_data_values[0][0][0]

    print("Information gathered from Excel")

    ''' download the projects. For that we need the
    user token'''

    root = tkinter.Tk()
    root.withdraw()
    token = simpledialog.askstring(
        'Token required', 'Please enter the number after "token="')

    if token is None or len(token) == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title='Error', message="You didn't introduce a token, \
            so farewell!")
        sys.exit()
    test = token.isdigit()
    if test is False:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title='Error', message='The token you pasted is incorrect. \
            Please paste only the number after "token=" in the address \
            bar of your browser.')
        sys.exit()

    print("Getting openertoken")
    get_openertoken = requests.get('http://example-worldserver.com/ws/\
    assignments_tasks?&project=' +
                                   str(sample_proj_id) +
                                   '&token=' + str(token_value), verify=False)
    r_foropener = get_openertoken.text
    opener_rgx = re.compile('openertoken=' + '([0-9]+)' + '(&)?', re.I)
    find_opener = re.search(opener_rgx, r_foropener)
    try:
        openertoken_value = find_opener.group(1)
    except AttributeError:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror(
            title="Error", message="Please refresh the WS tab, \
            the token is outdated")
        sys.exit()

    number_of_AIDs = len(AID_data)
    counter = 0
    print("Starting download, this might take a few minutes.")
    for key, value in AID_data.items():
        counter += 1
        download_WS_project(token_value, openertoken_value,
                            value[0], value[1], key, value[2])
        print("Downloaded Assignment " + str(counter) +
              " out of " + str(number_of_AIDs))

    userpro = os.environ['USERPROFILE']

    if len(missing_proj_dict) != 0:
        report_file = open(userpro + '\\Downloads\\Download_report.txt', 'w')
        report_file.write("Assignment ID    Missing projects\n")
        for k, v in missing_proj_dict.items():
            report_file.write(
                str(k) + "  " + " ,".join([value for value in v]) + "\n")
        report_file.close()

    missing_link.sort()

    '''Once the download is finished, we inform of the result'''

    if len(missing_link) != 0 and len(missing_proj_dict) == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(
            title="Success!", message="Download is finished. There was no hyperlink or \
            an invalid one in the following row(s): " +
            ", ".join(missing_link))
        sys.exit()
    elif len(missing_link) == 0 and len(missing_proj_dict) != 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(
            title="Success!", message="Download is finished. \
            Some projects could not be downloaded, see \
            'Download_report.txt' in the Downloads folder")
        sys.exit()
    elif len(missing_link) != 0 and len(missing_proj_dict) != 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(title="Success!", message="Download is finished.\
        Some projects could not be downloaded (see 'Download_report.txt') \
        and there was no hyperlink or an invalid one in the following \
        row(s): " + ", ".join(missing_link))
        sys.exit()
    elif len(missing_link) == 0 and len(missing_proj_dict) == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo(
            title="Success!", message="Download is finished, \
            all projects were downloaded and there are no missing hyperlinks.")
        sys.exit()
except Exception as e:
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showerror(title="Error", message=str(e))
    sys.exit()
