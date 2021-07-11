'''
author: fharookshaik
email: fharookshaik.5@gmail.com
GitHub username: fharookshaik
'''

# Makesure you've stored the outlook credentials in credentials.json file

import ClointFusion as cf
import os
import time
import shutil as sh
import sys

cf.OFF_semi_automatic_mode()

# GLOBAL vARIABLES
WORKSPACE_DIR = os.getcwd()
DOWNLOADS_DIR = os.path.join(WORKSPACE_DIR,'DOWNLOADS')
REPORTS_DIR = os.path.join(WORKSPACE_DIR,'REGIONAL_REPORTS')
REPORTS_ZIP = '.'.join([REPORTS_DIR,'zip'])

WHO_LINK = 'http://www.who.int/'
WHO_COVID_EXPLORER_LINK = 'https://worldhealthorg.shinyapps.io/covid/'

GLOBAL_REGIONS = ['African Region','Region of the Americas','Eastern Mediterranean Region','European Region','South-East Asia Region','Western Pacific Region']

def instantiate_website():
    browser_state = False
    try:
        browser_state = cf.launch_website_h(URL=WHO_COVID_EXPLORER_LINK,files_download_path=DOWNLOADS_DIR)
    except Exception as e:
        print('Error in instantiate_website = %s' % e)
    finally:
        return browser_state

def _accept_terms():
    try:
        cf.browser_wait_until_h(text='I ACCEPT')
        time.sleep(2)
        cf.browser_mouse_click_h(User_Visible_Text_Element='I ACCEPT',element='d')
    
    except Exception as e:
        print('Error in accept_terms = ', str(e))

def _redirect_to_regional_overview():
    try:
        cf.browser_mouse_hover_h(User_Visible_Text_Element='Regional Overview')
        cf.browser_mouse_click_h(User_Visible_Text_Element='Regional Overview',element='d',to_left_of='Country/Area/Territory',to_right_of='Global Overview')

        cf.browser_wait_until_h(text='Regional Epidemic Curve')

    except Exception as e:
        print('Error in _redirect_to_regional_overview = ' + str(e))

def _organize_regional_data(option='',type=''):
    time.sleep(2)
    try:
        NEW_FOLDER = os.path.join(REPORTS_DIR,f'{option}')
        
        if not os.path.exists(NEW_FOLDER):
            cf.folder_create(NEW_FOLDER)
        
        files = cf.folder_get_all_filenames_as_list(strFolderPath=DOWNLOADS_DIR)
        # print(files)
        for file in files:
            file_name,ext = file.split('\\')[-1].split('.')
            old_file_path = os.path.join(DOWNLOADS_DIR,file)
            if type == 'CASES':
                new_file_name = '.'.join([f'{file_name}_CASES',ext])
            else:
                new_file_name = '.'.join([f'{file_name}_DEATHS',ext])
                         
            cf.file_rename(old_file_path=old_file_path, new_file_name=new_file_name,ext=True)
            sh.move(src=os.path.join(DOWNLOADS_DIR,new_file_name),dst=os.path.join(NEW_FOLDER,new_file_name))
            print(f'Moved {new_file_name}')

    except Exception as e:
        print('Error in _organize_regional_data = ',str(e))
        sys.exit()

def _download_regional_cases_data(option=''):
    try:
        cf.browser_mouse_click_h('CASES',element='d')
        cf.browser_mouse_click_h('DOWNLOAD PLOT',element='d',)
        _organize_regional_data(option,type='CASES')

    except Exception as e:
        print('Error in _download_regional_cases_data = ',str(e))

def _download_regional_deaths_data(option=''):
    try:
        cf.browser_mouse_click_h('DEATHS',element='d')
        cf.browser_mouse_click_h('DOWNLOAD PLOT',element='d')
        _organize_regional_data(option,type='DEATHS')


    except Exception as e:
        print('Error in _download_regional_deaths_data = ',str(e))

def download_regional_overview_reports():
    try:
        _redirect_to_regional_overview()
        
        for option in GLOBAL_REGIONS:
            if option == 'African Region':
                prev_option = option
            
            else:
                prev_option = GLOBAL_REGIONS[GLOBAL_REGIONS.index(option) - 1]
            
            cf.browser_mouse_click_h(prev_option,element='d')
            cf.browser_mouse_click_h(option,element='d')
            time.sleep(2)

            _download_regional_cases_data(option)
            _download_regional_deaths_data(option)
            time.sleep(2)

    except Exception as e:
        print('Error in download_regional_overview_reports = ',str(e))

def zip_reports_dir():
    try:
        base = os.path.basename(REPORTS_ZIP)
        name,format = base.split('.')
        archive_from = os.path.dirname(REPORTS_DIR)
        archive_to = os.path.basename(REPORTS_DIR.strip(os.sep))
        sh.make_archive(name,format, archive_from,archive_to)
   
    except Exception as e:
        print('Error in _zip_reports_dir = ',str(e))

def send_outlook_email():
    try:
        CREDENTIALS_JSON = os.path.join(WORKSPACE_DIR,'credentials.json')
        outlook_details = cf.file_get_json_details(path_of_json_file=CREDENTIALS_JSON,section='Outlook')
        outlook_username = outlook_details.get('username')
        outlook_password = outlook_details.get('password')

        cf.browser_navigate_h('outlook.com')
        cf.browser_mouse_click_h('Sign in',element='d')

        cf.browser_write_h(outlook_username,User_Visible_Text_Element='Email, phone, or Skype')
        cf.browser_mouse_click_h('Next',element='d')

        cf.browser_write_h(outlook_password,User_Visible_Text_Element='Password')
        cf.browser_mouse_click_h('Sign in',element='d')

        cf.browser_mouse_click_h('New message',element='d')

        time.sleep(2)
        to_email = cf.gui_get_any_input_from_user('Enter to email: ')
        cf.browser_write_h(to_email,User_Visible_Text_Element='To')

        cf.browser_write_h('WHO COVID-19 Global Regional Reports',User_Visible_Text_Element='Add a subject')

        body_elem = cf.browser_locate_element_h("//*[@aria-label='Message body']")
        cf.browser_write_h('Please find the attached Reports.\n\n\nThanks & Regards\nWHO COVID-19 BOT',User_Visible_Text_Element=body_elem)

        cf.browser_mouse_click_h(User_Visible_Text_Element='Attach',element='d')
        cf.browser_mouse_click_h(User_Visible_Text_Element='Browse this computer',element='d')
        cf.key_write_enter(strMsg=REPORTS_ZIP)

        time.sleep(5)

        cf.browser_mouse_click_h('Send',element='d')

    except Exception as e:
        print('Error in Sending Outlook Email = ',str(e))


if __name__ == '__main__':
    try:
        # Emptying Previous Download Files
        if not os.path.exists(DOWNLOADS_DIR):
            cf.folder_create(strFolderPath=DOWNLOADS_DIR)
        else:
            cf.folder_delete_all_files(DOWNLOADS_DIR)
        
        # Removing REPORTS.zip
        if os.path.exists(REPORTS_ZIP):
            os.remove(REPORTS_ZIP)

        # Creating Reports Folder
        if not os.path.exists(REPORTS_DIR):
            cf.folder_create(REPORTS_DIR)
        else:
            cf.folder_delete_all_files(REPORTS_DIR)

        browser_state = instantiate_website()

        if browser_state == True:
            # Accept Terms & Conditions
            _accept_terms()

            # Downloading Regional Overview Reports
            download_regional_overview_reports()

            # Zip the Downloaded Reports
            zip_reports_dir()

            # Send Gmail
            send_outlook_email()

            time.sleep(5)
            cf.browser_quit_h()

    except Exception as e:
        print('Exception Raised = ', str(e))