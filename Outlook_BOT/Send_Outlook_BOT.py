# Author: Murali Manohar Verma

import time
import os

try:
    import ClointFusion as cf  # https://pypi.org/project/ClointFusion/
except ModuleNotFoundError:
    os.system("pip install ClointFusion")
    import ClointFusion as cf

try:
    import cryptocode
    # ClointFusion stores your password in an encrypted way. So, you need to decrypt password at run-time.
except ModuleNotFoundError:
    cf.os.system("pip install cryptocode")
    import cryptocode

# This BOT can be run in 2 ways: Automatic and Semi-Automatic Way

cf.OFF_semi_automatic_mode()

user_choice = cf.gui_get_dropdownlist_values_from_user(msgForUser=":", dropdown_list=["Automatic", "Semi Automatic"],
                                                       multi_select=False)[0]
username = ''
password = ''
to_address = ''
subject = ''
body_email = ''
if user_choice == "Automatic":
    cf.message_counter_down_timer(
        "Launching Outlook BOT in '{}' way.\n\nStored responses will be used in this execution".format(user_choice))

    cf.ON_semi_automatic_mode()  # Run BOT in automatic mode, without asking any GUI based questions. GUI questions
    # would be asked first time, to store the responses

    username = cf.gui_get_any_input_from_user("Email Username/Login")  # GUI pop-up window, first time

    password = cf.gui_get_any_input_from_user("your Email Password", True)  # GUI pop-up window, first time

    password = (str(cryptocode.decrypt(password, "ClointFusion")).strip())  # GUI pop-up window, first time

    to_address = cf.gui_get_any_input_from_user("To Email Address")  # GUI pop-up window, first time

    subject = cf.gui_get_any_input_from_user("Email Subject")  # GUI pop-up window, first time

    body_email = cf.gui_get_any_input_from_user("Email Body", multi_line=True)  # GUI pop-up window, first time

elif user_choice == "Semi Automatic":
    cf.message_counter_down_timer(
        "Launching Outlook BOT in '{}' way.\n\nYou can change GUI responses".format(user_choice))

    username = cf.gui_get_any_input_from_user("Email Username/Login")
    #                                                  GUI pop-up window, Every Time, (pre-filled with last responses)

    password = cf.gui_get_any_input_from_user("your Email Password", True)  # GUI pop-up window, Every Time

    to_address = cf.gui_get_any_input_from_user("To Email Address")  # GUI pop-up window, Every Time

    subject = cf.gui_get_any_input_from_user("Email Subject")  # GUI pop-up window, Every Time

    body_email = cf.gui_get_any_input_from_user("Email Body", multi_line=True)  # GUI pop-up window, Every Time

cf.launch_website_h("https://login.live.com/")

cf.browser_write_h(username, "email")  # Type email address

cf.browser_mouse_click_h("next")  # Click BLUE next button

cf.browser_write_h(password, "Password")  # Type password

cf.key_hit_enter()

time.sleep(3)

cf.key_press("ctrl + l")

cf.key_write_enter("https://outlook.live.com/mail/0/inbox")

cf.browser_mouse_click_h("New message")  # Click on Compose button

time.sleep(4)

cf.browser_write_h(to_address, "To")  # Type directly in Recipient's To Email Address

time.sleep(2)

cf.browser_write_h(subject, "Add a subject")  # Type the subject

time.sleep(4)

cf.key_press('tab')  # Move from Subject field to Body part

cf.key_write_enter(body_email)

time.sleep(4)

cf.browser_mouse_click_h("Send")  # Click BLUE Send button

time.sleep(2)

# cf.browser_wait_until_h("Message sent.") #Wait till email is sent

cf.browser_quit_h()

cf.message_counter_down_timer(strMsg="Mail Sent. Exiting now !", start_value=3)
