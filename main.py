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
from datetime import date
from datetime import datetime



def print_welcome():
    os.system('cls')
    print("                  AutoMail Client                   ")
    print("----------------------------------------------------\n")


def menu_select(sender_email, contact_file, html_file, server, curr_port):
    print_welcome()
    print("Sender email: " + sender_email)
    print("SMTP server: " + server)
    print("SMTP port: " + str(curr_port))
    print("Contact list file: " + contact_file)
    print("Contact list file: " + html_file)

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
    ans = input("Press ENTER to continue or 'q' to quit. ")
    if ans:
        return sender_email, pwd

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

        while True:
            print("\nSelected file: " + root.filename)
            ans = input("\nIs this the file you wish to import? Press ENTER to continue, or 'N' to select new file. ").lower()

            if ans == 'n':
                break
            else:
                return root.filename


#File selection screen for contact list
def select_contact_list():
    print_welcome()
    print("\nPress ENTER to select a contact list file to import or 'q' to quit. Please select .CSV file type.")
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

    #TODO: post processing?
    return data


#file selection screen for HTML email
def get_html_email():
    print_welcome()
    print("\nPress ENTER to select the HTML email to import. Please select .html or .txt file type.")
    selected_file_path = select_file("html")
    print("\nSelected contact file to import: " + selected_file_path)
    email_str = read_html_as_string(selected_file_path)
    print("\nPress ENTER to continue to return to menu.")
    input()
    return selected_file_path, email_str


def read_html_as_string(file_path):
    file_object = codecs.open(file_path, "r")
    email_str = file_object.read()
    return email_str


def update_smtp(smtp_server, port):
    print_welcome()
    print("Current SMTP server: " + smtp_server)
    print("Current port: " + str(port) + "\n")

    ans = input("Press ENTER to continue or 'q' to quit. ")
    if ans:
        return smtp_server, port

    updated_server = input("\nPlease enter new SMTP server: ")
    updated_port = input("\nPlease enter new port: ")

    input("\nPress ENTER to return to menu")
    return updated_server, updated_port


def confirm_settings(sender_email, smtp_server, port, contact_file, html_file):
    while True:
        print_welcome()
        print("These are the current settings. Are you sure you want to continue?\n")
        print("Sender email: " + sender_email)
        print("SMTP server: " + smtp_server)
        print("SMTP port: " + str(port))
        print("Contact list file: " + contact_file)
        print("Contact list file: " + html_file)
        ans = input("\nPress 'Y' to continue or 'N' to return to menu. ")
        if ans.lower() == 'n':
            return 'n'
        if ans.lower() == 'y':
            return 'y'
        else:
            print("Invalid input. Please select either 'Y' or 'N'. Press ENTER to continue: ")
            input()
            continue


def confirm_email(subject, email_from, html_file,):
    up_name = email_from
    up_subj = subject
    e_flag = False
    while True:
        print_welcome()
        print("These are the current settings. Are you sure you want to continue?\n")
        print("Sender Name: " + up_name)
        print("Subject: " + up_subj)
        print("HTML Email file: " + html_file)
        ans = input("\nPress 'Y' to continue, 'E' to edit name/subject, 'P' to preview HTML document in web browser, or 'q' to quit to menu. ")
        if ans.lower() == 'y':
            return up_subj, up_name, e_flag
        elif ans.lower() == 'e':
            up_name = input("\n Please input new sender name (not email): ")
            up_subj = input("Please input new subject: ")
        elif ans.lower() == 'p':
            webbrowser.open_new_tab(html_file)
            q_code = input("If you would like to change this, press 'Q' to exit to menu. or ENTER to continue.")
            if q_code.lower() == 'q':
                e_flag = True
                return up_subj, up_name, e_flag
        elif ans.lower() == 'q':
            return up_subj, up_name, e_flag


def send_email(receiver_email, sender_email, pwd, email_str, smtp_server, port, subj, from_name):
    today = date.today()
    date_sent = today.strftime("%m/%d/%Y")

    message = MIMEMultipart("alternative")
    message["Subject"] = subj
    message["From"] = from_name
    message["To"] = receiver_email

    html = email_str

    part1 = MIMEText(html, "html")

    message.attach(part1)

    # create a secure SSL context
    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, pwd)
            server.sendmail(
                sender_email, receiver_email, message.as_string()
            )
        print("Successfully sent email to " + receiver_email)
        return receiver_email, "Success", date_sent

    except smtplib.SMTPException as error:
        print("\nCould not send email to " + receiver_email)
        print("Error: " + str(error))
        return receiver_email, "Failed", date_sent


def write_result(data):
    now = datetime.now()
    dt_str = now.strftime("%d-%m-%Y %H-%M-%S")
    out_name = "output " + dt_str + ".csv"
    with open(out_name, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)


'''
 Secure SSL/TLS Settings (Recommended)
Username: 	susanlawson@lawsonmyside.com
Password: 	Use the email account’s password.
Incoming Server: 	mail.lawsonmyside.com

    IMAP Port: 993
    POP3 Port: 995

Outgoing Server: 	mail.lawsonmyside.com

    SMTP Port: 465

PW: lC@2sQ^9r&C6Eg
'''
# TODO: create separate program that reads in big csv file and separates into smaller ones

#-------------------------Main program starts here! -------------------------------------------
senderEmail = "jubox79@gmail.com"
password = "W4HS7KO2"
contactFilePath = "Not Selected"
htmlFilePath = "Not Selected"
smtpServer = "smtp.gmail.com"
smtpPort = 465
#email settings
senderName = "Susan Lawson, Esq."
emailSubject = "Attorney Referral Fees"
listData = [[]]
htmlStr = ""
resultData = []

while True:
    print_welcome()
    usrSelect = menu_select(senderEmail, contactFilePath, htmlFilePath, smtpServer, smtpPort)
    if usrSelect == '1':
        #function to update login credentials
        loginCredentials = get_login_credentials(senderEmail, password)
        senderEmail = loginCredentials[0]
        password = loginCredentials[1]

    if usrSelect == '2':
        #function to update SMTP settings
        smtpInfo = update_smtp(smtpServer, smtpPort)
        smtpServer = smtpInfo[0]
        smtpPort = smtpInfo[1]

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

    elif usrSelect == '5':
        exit_flag = False
        send = confirm_settings(senderEmail, smtpServer, smtpPort, contactFilePath, htmlFilePath)
        if send == 'n':
            exit_flag = True
        elif send == 'y':
            add_email_info = confirm_email(emailSubject, senderName, htmlFilePath)
            emailSubject = add_email_info[0]
            senderName = add_email_info[1]
            exit_flag = add_email_info[2]
        if exit_flag:
            continue
        else:
            for contact in listData[1:]:
                receiverAddress = contact[0]
                #print(receiverAddress)
                result = send_email(receiverAddress, senderEmail, password, htmlStr, smtpServer, smtpPort, emailSubject, senderName)
                resultData.append(list(result))
        input("Done! Press ENTER to continue.")
        write_result(resultData)

    elif usrSelect == 'q':
        quit()





