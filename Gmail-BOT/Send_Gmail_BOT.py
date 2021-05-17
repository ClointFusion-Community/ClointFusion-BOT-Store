# Author: ClointFusion
# Project: Gmail Sending Pythonic BOT, built using ClointFusion package
# Please visit: https://lnkd.in/gh_r9YB

import time
import os

try:
    import ClointFusion as cf #https://pypi.org/project/ClointFusion/
except:
    os.system("pip install ClointFusion")
    import ClointFusion as cf

try:
    import cryptocode #ClointFusion stores your password in an encrypted way. So, you need to decrypt password at run-time. 
except:
    cf.os.system("pip install cryptocode")
    import cryptocode

# This BOT can be run in 2 ways: Automatic and Semi-Automatic Way

cf.OFF_semi_automatic_mode()

user_choice = cf.gui_get_dropdownlist_values_from_user(msgForUser=":",dropdown_list=["Automatic","Semi Automatic"], multi_select=False)[0]

if user_choice == "Automatic":
    cf.message_counter_down_timer("Launching GMail BOT in '{}' way.\n\nStored responses will be used in this execution".format(user_choice))

    cf.ON_semi_automatic_mode() #Run BOT in automatic mode, without asking any GUI based questions. GUI questions would be asked first time, to store the responses

    username = cf.gui_get_any_input_from_user("Email Username/Login") #GUI pop-up window, first time 

    password = cf.gui_get_any_input_from_user("your Email Password",True) #GUI pop-up window, first time 

    password = (str(cryptocode.decrypt(password, "ClointFusion")).strip()) #GUI pop-up window, first time 

    tomail = cf.gui_get_any_input_from_user("To Email Address") #GUI pop-up window, first time 

    subjectmail = cf.gui_get_any_input_from_user("Email Subject") #GUI pop-up window, first time 

    bodypartgmail = cf.gui_get_any_input_from_user("Email Body",multi_line=True) #GUI pop-up window, first time 

elif user_choice == "Semi Automatic":
    cf.message_counter_down_timer("Launching GMail BOT in '{}' way.\n\nYou can change GUI responses".format(user_choice))

    username = cf.gui_get_any_input_from_user("Email Username/Login") #GUI pop-up window, Every Time, (pre-filled with last responses)  

    password = cf.gui_get_any_input_from_user("your Email Password",True) #GUI pop-up window, Every Time

    tomail = cf.gui_get_any_input_from_user("To Email Address") #GUI pop-up window, Every Time

    subjectmail = cf.gui_get_any_input_from_user("Email Subject") #GUI pop-up window, Every Time    

    bodypartgmail = cf.gui_get_any_input_from_user("Email Body",multi_line=True) #GUI pop-up window, Every Time

cf.launch_website_h("https://mail.google.com")

cf.browser_write_h(username,"email") #Type email address

cf.browser_mouse_click_h("next") #Type BLUE next button

cf.browser_write_h(password,"enter your password") #Type password

cf.browser_mouse_click_h("next") #Type BLUE next button

cf.browser_mouse_click_h("compose") #Click on Compose button

time.sleep(4)

cf.browser_write_h(tomail,"To") #Type directly in Recepient's To Email Address

time.sleep(2)

cf.browser_write_h(subjectmail,"Subject") #Type the subject

cf.key_press('tab') #Move from Subject field to Body part

cf.key_write_enter(bodypartgmail)

cf.browser_mouse_click_h("send") #Click BLUE Send button

cf.browser_wait_until_h("Message sent.") #Wait till email is sent

cf.browser_quit_h()

cf.message_counter_down_timer(strMsg="Mail Sent. Exiting now !", start_value=3)
