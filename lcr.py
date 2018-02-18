#!/usr/bin/env python3
"""A script to help automate the task of emailing LCR reports to parents."""

import os
import xlrd
import sys
import smtplib
import shutil
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders


class Error(Exception):
    """Base error class for this module"""
    pass


class MissingFileError(Error):
    """Raised when a directory is accessed but does not contain the necessary
    files/data

    Attributes:
        file -- the name of the missing file
        path -- the filepath of the directory being accessed
    """

    def __init__(self, file, path):
        self.file = file
        self.path = path


class MissingColumnError(Error):
    """Raised when a spreadsheet is missing a column required by the program

    Attributes:
        column -- the name of the required column
        sheetName -- the name of the spreadsheet with the missing column
    """

    def __init__(self, col, sheet):
        self.column = col
        self.sheetName = sheet


def lcr():
    # Initialize relevant paths
    cwd = os.getcwd()
    template_email = os.path.join(cwd, "TemplateEmail.html")
    test_total_ss = os.path.join(cwd, "AchievementTestData.xls")

    print("Please start by exporting end of level reports from CMS."
          "\n* Include all columns when generating the spreadsheet."
          "\n* Ensure that the corresponding spreadsheets are named math.xls"
          " and reading.xls")
    working_directory = getLCRDirectory()
    math_ss = os.path.join(working_directory, "math.xls")
    reading_ss = os.path.join(working_directory, "reading.xls")

    test_totals = {}
    """
       Contains all achievement test data pulled from test_total_ss
        -- Key is a string with the subject and test level (eg. "Math 4A")
        -- Value is a dictionary: {'totalMarks': ,'suggestedTime': }
    """

    tests_taken = {}
    """
       Contains all testing data from math_ss and reading_ss
        -- Key is a string with the subject and row number (eg. "Math 3")
           corresponding to a single test entry in the relevant spreadsheet
        -- Value is a dictionary relating all REQUIRED_COLUMNS to their
           relevant values
    """

    # Grab all the data contained in the spreadsheets and arrange it in the
    # test_totals and tests_taken dictionaries
    loadSpreadsheet(test_total_ss, 0, ["level"], test_totals)
    loadSpreadsheet(test_total_ss, 1, ["level"], test_totals)
    loadSpreadsheet(math_ss, 0, ["Subject"], tests_taken,
                    include_row_in_key=True)
    loadSpreadsheet(reading_ss, 0, ["Subject"], tests_taken,
                    include_row_in_key=True)

    # Add the relevant total score and suggested time for every test taken
    assignTestTotals(test_totals, tests_taken)

    # Create folders for every student's test, each with it's own corresponding
    # email
    prepEmailFolders(tests_taken, template_email, working_directory)

    # PDFs must now be manually imported from Kumon's CMS into their
    # corresponding folder.
    input("Please import all LCR PDFs into their respective folders. "
          "Press Enter when ready...")

    # Create a folder for managing reports which could not be emailed
    print_folder = os.path.join(working_directory, "To Print")
    try:
        os.mkdir(print_folder)
    except OSError:
        pass

    # Ensure each student's folder contains only one PDF file
    checkPDFs(working_directory, [print_folder])

    # Prompt user for login credentials
    user, smtp = emailLogin()

    # Walk each student's folder and attempt to send an email
    for folder, subfolders, filenames in os.walk(working_directory):
        if folder == working_directory or folder == print_folder:
            continue
        # The student's key in tests_taken has been formatted into the name
        # of each folder after ' --- '
        student_info = tests_taken[folder.split(' --- ')[1]]
        msg, attachment = assembleEmail(user, student_info,
                                        folder, filenames)
        try:
            smtp.sendmail(user, msg['To'], msg.as_string())
        except smtplib.SMTPException:
            # If an email cannot be sent, the report must be printed
            print("Failed to send email to " + msg['To'])
            print("LCR report for " + student_info['FirstName'] + " "
                  + student_info['LastName'] + " must be printed. "
                  + "It has been moved to " + dest)
            dest = print_folder + os.path.split(attachment)[1]
            src = attachment
            shutil.copyfile(src, dest)
            continue
        print("Successfully sent email for " + student_info['FirstName']
              + " " + student_info['LastName'] + " to " + msg['To'])

    smtp.quit()
    print("Finished! :)")
    input("Press Enter to exit")


def getLCRDirectory():
    """
       Returns the path of the directory containing math.xls and reading.xls
       Takes in, as user input, the path of a folder and checks its validity
    """
    cwd = os.getcwd()
    while True:
        target_directory = input("Enter the filepath for the folder containing"
                                 " LCR spreadsheets: ")
        try:
            os.chdir(target_directory)
            os.chdir(cwd)
            checkTarget(target_directory)
        except OSError:
            print("Invalid directory. Please try again.")
            continue
        except MissingFileError as err:
            print("*** ERROR Could not find " + err.file + " in "
                  + err.path + "\nPlease include it and try again.")
            continue
        except MissingColumnError as err:
            print("*** ERROR Missing column " + err.column + " in spreadsheet "
                  + err.sheetName + "\nPlease include it and try again.")
            continue
        break
    return target_directory


def checkTarget(target_directory):
    """
       Throws a MissingFileError if the target directory does not contain the
       REQUIRED_SPREADSHEETS
    """
    REQUIRED_SPREADSHEETS = ['math.xls', 'reading.xls']
    for ss in REQUIRED_SPREADSHEETS:
        path = os.path.join(target_directory, ss)
        if not os.access(path, os.F_OK):
            raise MissingFileError(ss, target_directory)
        checkLCRSpreadsheet(path)


def checkLCRSpreadsheet(spreadsheet_path):
    """
       Throws a MissingColumnError if the spreadsheet does not contain all of
       the columns necessary to the rest of the script.
    """
    REQUIRED_COLUMNS = ['FirstName', 'LastName', 'Subject', 'Type', 'Time',
                        'Score', 'FatherEmail', 'MotherEmail', 'Passing']
    spreadsheet_name = os.path.split(spreadsheet_path)[1]
    print("Checking columns in " + spreadsheet_name + "...")
    workbook = xlrd.open_workbook(spreadsheet_path)
    sheet = workbook.sheet_by_index(0)
    columns = []
    for col in range(sheet.ncols):
        columns.append(sheet.cell_value(0, col))
    for col in REQUIRED_COLUMNS:
        if col not in columns:
            raise MissingColumnError(col, spreadsheet_name)


def loadSpreadsheet(spreadsheet,
                    sheet_index,
                    keys,
                    dictionary,
                    include_row_in_key=False):
    """
       Extracts the data from the input spreadsheet and arranges it in a
       dictionary. The 'keys' parameter is a list containing the column headers
       used in generating the dictionary keys used to index each row in the
       spreadsheet
    """
    workbook = xlrd.open_workbook(spreadsheet)
    sheet = workbook.sheet_by_index(sheet_index)

    # Generate a list containing the column numbers of the desired keys in
    # the spreadsheet
    key_cols = keys.copy()
    for col in range(sheet.ncols):
        for i in range(len(key_cols)):
            if str(sheet.cell_value(0, col)).strip() == key_cols[i]:
                key_cols[i] = col

    # Generate a dictionary entry for each row in the spreadsheet. Note that
    # columns used to generate the key are included in the value as well.
    # This is because their data may be necessary to other parts of the
    # program. Values extracted from spreadsheet randomly include whitespace,
    # so it is always stripped.
    for row in range(1, sheet.nrows):
        key = key_cols.copy()
        value = {}
        for col in range(sheet.ncols):
            if col in key:
                for i in range(len(key)):
                    if col == key[i]:
                        key[i] = str(sheet.cell_value(row, col)).strip()
            value[str(sheet.cell_value(0, col)).strip()] \
                = str(sheet.cell_value(row, col)).strip()
        if include_row_in_key:
            dictionary[" ".join(key) + " " + str(row)] = value
        else:
            dictionary[" ".join(key).strip()] = value


def assignTestTotals(totals, tests_taken):
    """
       For each test in 'tests_taken', adds the relevant total score and
       suggested time from 'totals'
    """
    for s in tests_taken.keys():
        student = tests_taken[s]
        total_key = student["Subject"].strip() + " " + student["Type"].strip()
        for t in totals[total_key].keys():
            student[t] = totals[total_key][t]


def prepEmailFolders(tests_taken, template_email, target_directory):
    """
       Generates a folder for each test taken and includes the body of the
       corresponding email report
    """
    # Numbers extracted by xlrd are floats by default. For formatting purposes,
    # these columns are treated as integers instead.
    INTEGER_COLUMNS = ['Time', 'suggestedTime', 'Score', 'totalMarks']
    for s in tests_taken.keys():
        student = tests_taken[s]
        if student["Passing"] == "No":
            continue
        # Create the folder which will contain the email and its PDF attachment
        folder_name = student['FirstName'].strip() + " " + \
            student['LastName'].strip() + " Level " + \
            student['Type'].strip() + " --- " + str(s)
        folder_path = os.path.join(target_directory, folder_name)
        try:
            os.mkdir(folder_path)
        except FileExistsError:
            pass
        # Write the email for the current student's test
        for k in INTEGER_COLUMNS:
            student[k] = student[k].split(".")[0]
        template = Template(open(template_email).read())
        email_path = os.path.join(folder_path, "email.html")
        f = open(email_path, 'w')
        f.write(template.substitute(student))
        f.close()


def checkPDFs(target_directory, exceptions):
    """
       Verifies that each student folder in the target directory contains
       exactly one pdf file
    """
    print("Checking that lcr PDF files are sorted reasonably...")
    pdfs_sorted = False
    while not pdfs_sorted:
        pdfs_sorted = True
        for root, dirs, files in os.walk(target_directory):
            if root in exceptions:
                continue
            num_of_pdfs = len([f for f in files if f.endswith(".pdf")])
            if root == target_directory:
                if num_of_pdfs > 0:
                    input("*** ERROR: There are pdf files in the main LCR "
                          "directory. Please move these into the corresponding"
                          " student's folder, or remove them."
                          "\nPress Enter when ready...")
                    pdfs_sorted = False
                    break
                else:
                    continue
            if num_of_pdfs == 0:
                input("*** ERROR: No LCR PDF found in folder:\n" + root
                      + "\nPlease export the student's LCR report from CMS and"
                      + " place it in the folder.\nPress Enter when ready...")
                pdfs_sorted = False
                break
            elif num_of_pdfs > 1:
                input("*** ERROR: Multiple LCR PDFs found in folder:\n" + root
                      + "\nPlease ensure each student folder only contains one"
                      + " PDF file.\nPress Enter when ready...")
                pdfs_sorted = False
                break


def getRecipients(test_data):
    """
       Returns, as a string, the email addresses of the student's parents for
       the given test
    """
    recipients = []
    if test_data["MotherEmail"] != "":
        recipients.append(test_data["MotherEmail"])
    if test_data["FatherEmail"] != "":
        recipients.append(test_data["FatherEmail"])
    # Some entries for mother/father email are identical
    if len(recipients) == 2:
        if recipients[0] == recipients[1]:
            return recipients[0]
    return ", ".join(recipients)


def emailLogin():
    """
       Prompts the user to log in, and if successful, returns the user email
       address and an smpt connection
    """
    smtp = smtplib.SMTP("smtp.gmail.com", 587)
    smtp.ehlo()
    smtp.starttls()
    while True:
        try:
            user = input("LCR Email account address: ")
            password = input("Password: ")
            smtp.login(user, password)
        except smtplib.SMTPAuthenticationError:
            print("Failed to log in. Please try again")
            continue
        break
    return [user, smtp]


def assembleEmail(FROM, student_info, folder, filenames):
	"""
	   Returns the message and attachment of an email based on the files in
	   'folder'
	"""
    to = getRecipients(student_info)
    attachment = ''
    emailBody = ''
    subject = student_info['FirstName'] + "'s Level Completion Report"

    for f in filenames:
        if f.endswith('.pdf'):
            newName = folder + "\\" + student_info['LastName'] + ", " \
                      + student_info['FirstName'] + " - " \
                      + student_info['Subject'] + " " \
                      + student_info['Type'] + " level completion report.pdf"
            os.rename(os.path.join(folder, f), newName)
            attachment = newName
        if f.endswith('.html'):
            email = open(os.path.join(folder, f))
            emailBody = email.read()
            email.close()

    msg = MIMEMultipart()
    msg['From'] = FROM
    msg['To'] = to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(emailBody, 'html'))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',
                    'attachment; filename="{0}"'
                    .format(os.path.basename(attachment)))
    msg.attach(part)

    return [msg, attachment]


if __name__ == '__main__':
    lcr()
