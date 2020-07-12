from tkinter import Tk
from tkinter import filedialog
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from getpass import getpass
import codecs
import csv
import webbrowser
from datetime import datetime
import time


def print_welcome():
    os.system('cls')
    print("                  AutoMail Client                   ")
    print("----------------------------------------------------\n")


def menu_select(sender_email, contact_file, html_file, server, curr_port, plain_file, wait_time):
    print_welcome()
    print("Sender email: " + sender_email)
    print("SMTP server: " + server)
    print("SMTP port: " + str(curr_port))
    print("Contact list file: " + contact_file)
    print("HTML Email file: " + html_file)
    print("Plain-text Email file: " + plain_file)
    print("Time between each email: " + str(wait_time))

    print("\nPlease make a selection:")
    print("1. Update Login Credentials")
    print("2. Update SMTP Settings")
    print("3. Select Contact List to Import")
    print("4. Select HTML File to Email")
    print("5. Send Emails")

    ans = input("\nPlease input your selection here or press 'q' to quit: ")
    return ans


# Login screen with email sender confirmation and password retrieval
def get_login_credentials(sender_email, pwd):
    print_welcome()
    print("Current sender email: " + sender_email + "\n")

    updated_email = input("\nPlease enter new sender email: ")
    print("\nPlease enter new login credentials.")

    print("Sender email: " + updated_email)
    get_password = getpass("Password: ")

    input("\nPress ENTER to return to menu")
    return updated_email, get_password


def select_file(kind):
    while True:
        root = Tk()
        #root.lift()
        root.attributes("-topmost", True)
        root.withdraw()
        if kind == "csv":
            root.filename = filedialog.askopenfilename(filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        elif kind == "html":
            root.filename = filedialog.askopenfilename(filetypes=(("HTML files", "*.html"), ("txt files", "*.txt"), ("All files", "*.*")))
        elif kind == "txt":
            root.filename = filedialog.askopenfilename(filetypes=(("txt files", "*.txt"), ("All files", "*.*")))
        elif kind == "dir":
            root.filename = filedialog.askdirectory()

        while True:
            print("\nSelected file: " + root.filename)
            ans = input("\nIs this the correct file path? Press ENTER to continue, or 'N' to select new file. ").lower()
            if ans == 'n':
                break
            else:
                return root.filename


#File selection screen for contact list
def select_contact_list():
    print_welcome()
    print("\nPlease select contact list you wish to import. Please select .CSV file type.")
    selected_file_path = select_file("csv")
    print("\nSelected contact file to import: " + selected_file_path)

    print("Press ENTER to return to menu.")
    input()
    data = import_email_list(selected_file_path)
    return selected_file_path, data


def import_email_list(file_path):
    with open(file_path) as csv_file:
        csv_reader = csv.reader(csv_file)
        data = list(csv_reader)
    return data


#file selection screen for HTML email
def get_html_email():
    print_welcome()
    print("\nPlease select HTML email you wish to send. Please select .html or .txt file type.")
    selected_file_path = select_file("html")
    print("\nSelected HTML file to import: " + selected_file_path)
    email_str = read_file_as_string(selected_file_path)

    print("\nPlease select plain-text email you wish to send. Please select .txt file type.")
    plain_file_path = select_file("txt")
    print("\nSelected plain-text file to import: " + selected_file_path)
    plain_txt = read_file_as_string(plain_file_path)

    print("\nPress ENTER to continue to return to menu.")
    input()
    return selected_file_path, email_str, plain_file_path, plain_txt


def read_file_as_string(file_path):
    file_object = codecs.open(file_path, "r")
    email_str = file_object.read()
    return email_str


def update_smtp_headings(smtp_server, port, name, subj, sender_email, wait_time):
    updated_server = smtp_server
    updated_port = port
    updated_name = name
    updated_subj = subj
    updated_time = wait_time

    while True:
        print_welcome()
        print("1. Current SMTP server: " + updated_server)
        print("2. Current Port: " + str(updated_port))
        print("3. Current Sender Name: " + updated_name)
        print("4. Current Subject: " + updated_subj)
        print("5. Current Wait Time: " + str(updated_time) + "\n")

        sel = input("Please input the number of the setting you wish to modify or 'q' to exit with above settings: ")
        if sel == "1":
            updated_server = input("\nPlease enter new SMTP server: ")
            continue
        elif sel == "2":
            updated_port = input("\nPlease enter new port: ")
            continue
        elif sel == "3":
            u_name = input("\nPlease enter new sender name: ")
            updated_name = u_name + " <" + sender_email + ">"
            continue
        elif sel == "4":
            updated_subj = input("\nPlease enter new subject: ")
            continue
        elif sel == "5":
            updated_time = input("\nPlease enter new wait time: ")
            continue
        elif sel == "q":
            return updated_server, int(updated_port), updated_name, updated_subj, int(updated_time)
        else:
            print("Invalid input. Press ENTER to continue.")
            input()


def confirm_settings(sender_email, smtp_server, port, contact_file, html_file, name, subj, plain_file, wait_time):
    while True:
        print_welcome()
        print("These are the current settings. Are you sure you want to continue?\n")
        print("Sender email: " + sender_email)
        print("Sender Name: " + name)
        print("Subject: " + subj)
        print("HTML Email file: " + html_file)
        print("SMTP server: " + smtp_server)
        print("SMTP port: " + str(port))
        print("Contact list file: " + contact_file)
        print("HTMl email file: " + html_file)
        print("Plain-text email file: " + plain_file)
        print("Wait time: " + str(wait_time))
        ans = input("\nPress 'Y' to continue or 'N' to return to menu or 'P' to preview HTML file in browser.")
        if ans.lower() == 'n':
            return 'n'
        if ans.lower() == 'y':
            return 'y'
        if ans.lower() == 'p':
            webbrowser.open_new_tab(html_file)
        else:
            print("Invalid input. Please select either 'Y' or 'N'. Press ENTER to continue: ")
            input()
            continue


def send_email(receiver_email, sender_email, pwd, email_str, smtp_server, port, subj, from_name, plain_text):
    now = datetime.now()
    dt_sent = now.strftime("%d-%m-%Y %H-%M-%S")


    message = MIMEMultipart("alternative")
    message["Subject"] = subj
    message["From"] = from_name
    message["To"] = receiver_email

    html = email_str
    text = plain_text

    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    message.attach(part1)
    message.attach(part2)

    # create a secure SSL context
    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, pwd)
            server.sendmail(
                sender_email, receiver_email, message.as_string()
            )
        print("Successfully sent email to " + receiver_email)
        return receiver_email, "Success", dt_sent

    except smtplib.SMTPException as error:
        print("\nCould not send email to " + receiver_email)
        print("Error: " + str(error))
        return receiver_email, "Failed", dt_sent


def write_result(data, dest_folder):
    now = datetime.now()
    dt_str = now.strftime("%d-%m-%Y %H-%M-%S")
    out_name = "output " + dt_str + ".csv"
    target_filepath = os.path.join(dest_folder, out_name)
    with open(target_filepath, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)


#------------------------- Main program starts here! -------------------------------------------
#Hard-coded presets
senderEmail = "Not Selected"
password = "Not Selected"
contactFilePath = "Not Selected"
htmlFilePath = "Not Selected"
smtpServer = "Not Selected"
smtpPort = 465
senderName = "Not Selected"
emailSubject = "Not Selected"
waitTime = 60  # seconds
listData = [[]]
htmlStr = "Not Selected"
plainText = "Not Selected"
plainFilePath = "Not Selected"
resultData = []

while True:
    print_welcome()
    usrSelect = menu_select(senderEmail, contactFilePath, htmlFilePath, smtpServer, smtpPort,  plainFilePath, waitTime)
    if usrSelect == '1':
        #function to update login credentials
        loginCredentials = get_login_credentials(senderEmail, password)
        senderEmail = loginCredentials[0]
        password = loginCredentials[1]

    if usrSelect == '2':
        #function to update SMTP settings
        smtpInfo = update_smtp_headings(smtpServer, smtpPort, senderName, emailSubject, senderEmail, waitTime)
        smtpServer = smtpInfo[0]
        smtpPort = smtpInfo[1]
        senderName = smtpInfo[2]
        emailSubject = smtpInfo[3]
        waitTime = smtpInfo[4]

    elif usrSelect == '3':
        #function to select contact list
        contactList = select_contact_list()
        contactFilePath = contactList[0]
        listData = contactList[1]

    elif usrSelect == '4':
        #function to select HTML file
        htmlEmail = get_html_email()
        htmlFilePath = htmlEmail[0]
        htmlStr = htmlEmail[1]
        plainFilePath = htmlEmail[2]
        plainText = htmlEmail[3]

    elif usrSelect == '5':
        exit_flag = False
        send = confirm_settings(senderEmail, smtpServer, smtpPort, contactFilePath, htmlFilePath, senderName, emailSubject, plainFilePath, waitTime)
        if send == 'n':
            exit_flag = True
        elif send == 'y':
            print("Please select folder you wish to output results to.")
            outputPath = select_file("dir")
            for contact in listData[1:]:
                receiverAddress = contact[0]
                #print(receiverAddress)
                result = send_email(receiverAddress, senderEmail, password, htmlStr, smtpServer, smtpPort, emailSubject, senderName, plainText)
                resultData.append(list(result))
                time.sleep(waitTime)
            input("Done! Press ENTER to continue.")
            write_result(resultData, outputPath)

    elif usrSelect == 'q':
        quit()





