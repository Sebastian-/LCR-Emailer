#!/usr/bin/env python3
"""A script to help automate the task of sending out LCR reports to parents."""

import os
import xlrd
import sys
import smtplib
import getpass
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
    baseDirectory = os.getcwd()
    templateEmailPath = baseDirectory + "\\TemplateEmail.html"
    test_total_ss = baseDirectory + "\\AchievementTestData.xls"

    working_directory = getLCRDirectory()
    math_ss = working_directory + "\\math.xls"
    reading_ss = working_directory + "\\reading.xls"

    testTotals = {}
    """Contains all achievement test data pulled from test_total_ss
        -- Key is a string with the subject and test level (eg. "Math 4A")
        -- Value is a dictionary: {'totalMarks': , 'suggestedTime': }
    """
    
    lcrData = {}
    """Contains all LCR testing data from math_ss and reading_ss
        -- Key is a string with the subject and row number (eg. "Math 3")
           corresponding single test text entry in the relevant spreadsheet
        -- Value is a dictionary relating all REQUIRED_COLUMNS to their
           relevant values
    """

    # Grab all the data contained in spreadsheets
    loadWorksheet(test_total_ss, 0, ['level'], testTotals)
    loadWorksheet(test_total_ss, 1, ['level'], testTotals)
    loadWorksheet(math_ss, 0, ['Subject'], lcrData, rowNumbers=True)
    loadWorksheet(reading_ss, 0, ['Subject'], lcrData, rowNumbers=True)

    # Assign the relevant total score and suggested time for every test taken 
    assignTestTotals(testTotals, lcrData)

    """create a folder for each test taken to contain their respective email
    and pdf attachment"""
    prepEmailFolders(lcrData, templateEmailPath, working_directory)

    """pdfs are imported by hand from CMS, then a check is done to ensure
    each folder has exactly one pdf"""
    input("Please import all LCR PDFs into their respective folders. Press Enter when ready...")
    checkPDFs(working_directory)

    """prepare print folder"""
    try:
        os.mkdir(".\\To Print")
    except OSError:
        pass

    user, smtp = emailLogin()

    # walk each student's folder and attempt to send an email
    for folderName, subfolders, filenames in os.walk(working_directory):
        if folderName == working_directory or folderName == working_directory + "\\To Print":
            continue
        studentInfo = lcrData[folderName.split(' --- ')[1]]

        msg, attachmentPath = assembleEmail(user, studentInfo, folderName, filenames)

        toPrint = False

        try:
            smtp.sendmail(user, msg['To'], msg.as_string())
        except smtplib.SMTPRecipientsRefused:
            print("***ERROR: Invalid recipients")
            toPrint = True
        except Exception:
            print("Failed to send email for " + msg['To'])
            toPrint = True

        if toPrint:
            dest = working_directory + "\\To Print"
            src = attachmentPath
            print(studentInfo['FirstName'] + "'s LCR report must be printed. It has been moved to " + dest)
            shutil.move(src, dest)
            continue

        print("Successfully sent email for " + studentInfo['FirstName'] + " " + studentInfo['LastName'] + " to " + msg['To'])

    smtp.quit()
    print("Finished! :)")
    input("Press Enter to exit")


def getLCRDirectory():
    """TODO write this"""
    current_dir = os.getcwd()
    while True:
        targetDirectory = input("Enter the filepath for the folder containing LCR spreadsheets: ")
        try:
            os.access(targetDirectory, os.F_OK)
            checkTarget(targetDirectory)
            checkLCRSpreadsheets(targetDirectory)
        except OSError:
            print("Invalid directory. Please try again.")
            continue
        except MissingFileError as err:
            print("***Could not find " + err.file + " in "
                  + err.path)
            continue
        except MissingColumnError as err:
            print("***Missing column " + err.column + " in spreadsheet " 
                  + err.sheetName)
            continue
        break
    os.chdir(current_dir)
    return targetDirectory

# Ensure necessary columns are present in each spreadsheet
def checkLCRSpreadsheets(targetDirectory):
    current_dir = os.getcwd()
    os.chdir(targetDirectory)
    REQUIRED_COLUMNS = ['FirstName', 'LastName', 'Subject', 'Type', 'Time',
                        'Score', 'FatherEmail', 'MotherEmail', 'Passing']

    print("Checking spreadsheets...")
    s = [x for x in os.listdir(".") if x.endswith(".xls")]
    for spreadsheet in s:
        workbook = xlrd.open_workbook("." + '\\' + spreadsheet)
        sheet = workbook.sheet_by_index(0)
        columns = []
        for col in range(sheet.ncols):
            columns.append(sheet.cell_value(0, col))
        for col in REQUIRED_COLUMNS:
            if col not in columns:
                os.chdir(current_dir)
                raise MissingColumnError(col, spreadsheet)
    os.chdir(current_dir)


# Check that all lcr pdfs are reasonably sorted into folders
def checkPDFs(targetDirectory):
    os.chdir(targetDirectory)
    print("Checking that lcr PDF files are sorted reasonably...")
    pdfsSorted = False
    while not pdfsSorted:
        pdfsSorted = True
        for folder in [x[0] for x in os.walk(".")]:
            if folder == ".\\To Print":
                continue
            pdfs = len([file for file in os.listdir("." + "\\" + folder) if file.endswith(".pdf")])
            if folder == ".":
                if pdfs == 0:
                    continue
                if pdfs > 0:
                    input("*** ERROR: some pdfs have not been placed in a folder. "
                          + "Press Enter when ready.")
                    pdfsSorted = False
                    break
            if pdfs == 0:
                input("*** ERROR: no lcr report found in folder: " + folder
                      + " Press Enter when ready.")
                pdfsSorted = False
                break
            elif pdfs > 1:
                input("*** ERROR: mulitple lcr reports found in folder: " + folder
                      + " Press Enter when ready.")
                pdfsSorted = False
                break
        if pdfsSorted:
            break


# Ensure target folder contains correct excel files
def checkTarget(targetDirectory):
    # TODO raise MissingFileError
    current_dir = os.getcwd()
    os.chdir(targetDirectory)
    REQUIRED_SPREADSHEETS = ['math.xls', 'reading.xls']
    s = [x for x in os.listdir(".") if x.endswith(".xls")]

    for file in REQUIRED_SPREADSHEETS:
        if file not in s:
            os.chdir(current_dir)
            raise MissingFileError(file, targetDirectory)
    os.chdir(current_dir)


# Loads the data of an excel worksheet into the dictionary 'data'
# Key labels are strings containing the fields in desiredKeys[] in the order in which they are listed
# if rowNumbers is set, the row number of the entry will be appended to the label
def loadWorksheet(workbookPath, sheetIndex, desiredKeys, data, rowNumbers=False):
    workbook = xlrd.open_workbook(workbookPath)
    sheet = workbook.sheet_by_index(sheetIndex)
    keyColumns = desiredKeys.copy()
    for col in range(sheet.ncols):
        for i in range(len(keyColumns)):
            if str(sheet.cell_value(0, col)).strip() == keyColumns[i]:
                keyColumns[i] = col
    for row in range(1, sheet.nrows):
        key = keyColumns.copy()
        value = {}
        for col in range(sheet.ncols):
            if col in key:
                for i in range(len(key)):
                    if col == key[i]:
                        key[i] = str(sheet.cell_value(row, col)).strip()
            value[str(sheet.cell_value(0, col)).strip()] = str(sheet.cell_value(row, col)).strip()
        if rowNumbers:
            data[" ".join(key) + " " + str(row)] = value
        else:
            data[" ".join(key).strip()] = value


def assignTestTotals(totals, students):
    for s in students.keys():
        student = students[s]
        totalsKey = student["Subject"].strip() + " " + student["Type"].strip()
        for t in totals[totalsKey].keys():
            student[t] = totals[totalsKey][t]


def prepEmailFolders(studentData, templateEmailPath, targetDirectory):
    INTEGER_COLUMNS = ['Time', 'suggestedTime', 'Score', 'totalMarks']
    for s in studentData.keys():
        student = studentData[s]
        if student["Passing"] == "No":
            continue
        # create folder to hold the body of the email and its PDF attachment
        folderName = student['FirstName'].strip() + " " + \
            student['LastName'].strip() + " Level " + \
            student['Type'].strip() + " --- " + str(s)
        try:
            os.mkdir("." + '\\' + folderName)
        except OSError:
            continue
        os.chdir("." + "\\" + folderName)
        # prepare email body
        for k in INTEGER_COLUMNS:
            student[k] = student[k].split(".")[0]
        temp = Template(open(templateEmailPath).read())
        f = open('email.html', 'w')
        f.write(temp.substitute(student))
        f.close()
        os.chdir(targetDirectory)


def getRecipients(studentInfo):
    recipients = []
    if studentInfo['MotherEmail'] != '':
        recipients.append(studentInfo['MotherEmail'])
    if studentInfo['FatherEmail'] != '':
        recipients.append(studentInfo['FatherEmail'])
    if len(recipients) == 2:
        if recipients[0] == recipients[1]:
            return recipients[0]
    return ", ".join(recipients)


def emailLogin():
    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.ehlo()
    smtp.starttls()
    while True:
        try:
            user = input("LCR Email account address: ")
            smtp.login(user, getpass.getpass('Password:'))
        except smtplib.SMTPAuthenticationError:
            print("Failed to log in. Please try again")
            continue
        break
    return [user, smtp]


# returns a list of all data required to assemble an email [to, attachment, body, subject]
# from the files of a folder
def collectEmailData(studentInfo, filenames):
    to = getRecipients(studentInfo)
    attachment = ''
    emailBody = ''
    subject = studentInfo['FirstName'] + "'s Level Completion Report"

    for f in filenames:
        if f.endswith('.pdf'):
            newName = folderName + "\\" + studentInfo['LastName'] + ", " + studentInfo['FirstName'] + " - " + \
                      studentInfo['Subject'] + " " + studentInfo['Type'] + " level completion report.pdf"
            os.rename(folderName + "\\" + f, newName)
            attachment = newName
        if f.endswith('.html'):
            email = open(folderName + "\\" + f)
            emailBody = email.read()
            email.close()


# assembles an email based on the files in folder
def assembleEmail(FROM, studentInfo, folderName, filenames):
    to = getRecipients(studentInfo)
    attachment = ''
    emailBody = ''
    subject = studentInfo['FirstName'] + "'s Level Completion Report"

    for f in filenames:
        if f.endswith('.pdf'):
            newName = folderName + "\\" + studentInfo['LastName'] + ", " + studentInfo['FirstName'] + " - " + \
                      studentInfo['Subject'] + " " + studentInfo['Type'] + " level completion report.pdf"
            os.rename(folderName + "\\" + f, newName)
            attachment = newName
        if f.endswith('.html'):
            email = open(folderName + "\\" + f)
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
    part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(os.path.basename(attachment)))
    msg.attach(part)

    return [msg, attachment]


lcr()
