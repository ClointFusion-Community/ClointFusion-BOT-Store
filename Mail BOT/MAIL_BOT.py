# Author: Murali Manohar Varma
# Project: Mail Sending Pythonic BOT, built using ClointFusion package

import os

import ClointFusion as cf #https://pypi.org/project/ClointFusion/
import cryptocode

# This BOT can be run in 2 ways: Automatic and Semi-Automatic Way

username = ''
password = ''
to_address = ''
subject = ''
body_email = ''
attachments = []
still_files = False

cf.OFF_semi_automatic_mode()

mail_choice = cf.gui_get_dropdownlist_values_from_user(msgForUser=":",dropdown_list=["Gmail","Outlook"], multi_select=False)[0]

user_choice = cf.gui_get_dropdownlist_values_from_user(msgForUser=":",dropdown_list=["Automatic","Semi Automatic"], multi_select=False)[0]

def get_details():
    global username, password, to_address, subject, body_email, attachments, still_files
    
    username = cf.gui_get_any_input_from_user(f"{mail_choice} Username/Login") #GUI pop-up window, first time 

    password = cf.gui_get_any_input_from_user(f"Your {mail_choice} Email Password", True)

    to_address = cf.gui_get_any_input_from_user("To Email Address") #GUI pop-up window, first time 

    to_address = ",".join(to_address.split(",")) + ","
    
    subject = cf.gui_get_any_input_from_user("Email Subject") #GUI pop-up window, first time 

    body_email = cf.gui_get_any_input_from_user("Email Body",multi_line=True) #GUI pop-up window, first time 
    
    want_to_attach_files = cf.gui_get_consent_from_user("Want to attact files.")
    
    if want_to_attach_files == "Yes":
        still_files = True
    while still_files:
        attachments += [str(cf.ClointFusion.Path(cf.gui_get_any_file_from_user("File to be attached")))]
        more_files = cf.gui_get_consent_from_user("Attach more files")
        if more_files == 'Yes':
            still_files = True
        else:
            still_files = False
            
if user_choice == "Automatic":
    cf.message_counter_down_timer("Launching {} BOT in '{}' way.\n\nStored responses will be used in this execution".format(mail_choice, user_choice))

    cf.ON_semi_automatic_mode() #Run BOT in automatic mode, without asking any GUI based questions. GUI questions would be asked first time, to store the responses
    
    get_details()
    

elif user_choice == "Semi Automatic":
    cf.message_counter_down_timer("Launching {} BOT in '{}' way.\n\nYou can change GUI responses".format(mail_choice, user_choice))
    
    get_details()
            
if mail_choice == "Gmail":
    cf.launch_website_h("https://mail.google.com")

    cf.browser_write_h(username,"Email or phone") #Type email address

    cf.browser_mouse_click_h("Next", element="d") #Type BLUE next button

    cf.browser_write_h(password,"enter your password") #Type password

    cf.browser_mouse_click_h("Next", element="d") #Type BLUE next button

    cf.browser_mouse_click_h("Compose", element="d") #Click on Compose button

    cf.time_sleep(3)
    
    cf.browser_write_h(to_address) #Type directly in Recepient's To Email Address

    cf.time_sleep(2)

    cf.browser_write_h(subject,"Subject") #Type the subject

    cf.browser_key_press_h('tab') #Move from Subject field to Body part

    cf.browser_write_h(body_email)

    for file in attachments:
        cf.browser_mouse_click_h(User_Visible_Text_Element="Attach files", element="d")
        cf.time.sleep(1)
        cf.key_write_enter(strMsg=file)

    cf.browser_mouse_click_h("Send", element="d") #Click BLUE Send button

    cf.browser_wait_until_h("Message sent.") #Wait till email is sent

    cf.browser_quit_h()
    
    cf.message_flash(f"{mail_choice} Mail Sent Successfully. Thank You for using ClointFusion.", delay=5)

    # cf.message_counter_down_timer(strMsg="Mail Sent. Exiting now !", start_value=3)

if mail_choice == "Outlook":
    
    cf.launch_website_h("https://outlook.live.com/")

    cf.browser_mouse_click_h(User_Visible_Text_Element="Sign in", element="d")

    cf.browser_write_h(Value=username, User_Visible_Text_Element="Email, phone, or Skype")
    cf.browser_hit_enter_h()


    cf.browser_write_h(Value=password, User_Visible_Text_Element="Password")
    cf.browser_hit_enter_h()
    

    cf.browser_mouse_click_h(User_Visible_Text_Element="New message", element="d")

    cf.time_sleep(3)
    cf.browser_write_h(Value=to_address, User_Visible_Text_Element="To")

    cf.browser_write_h(Value=subject, User_Visible_Text_Element="Add a subject")
    cf.browser_key_press_h("TAB")
    cf.browser_write_h(body_email)

    for file in attachments:
        cf.browser_mouse_click_h(User_Visible_Text_Element="Attach", element="d")
        cf.browser_mouse_click_h(User_Visible_Text_Element="Browse this computer", element="d")
        cf.time_sleep(1)
        cf.key_write_enter(strMsg=file)

    cf.browser_mouse_click_h(User_Visible_Text_Element="Send", element="d")
    cf.browser_quit_h()

    cf.message_flash(f"{mail_choice} Mail Sent Successfully. Thank You for using ClointFusion.", delay=5)
