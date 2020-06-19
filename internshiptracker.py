import requests
from bs4 import BeautifulSoup
import datetime
import openpyxl
from openpyxl.styles import Alignment, Side
from openpyxl.styles import PatternFill, Font, Border
import tkinter as tk
from tkinter import *

root = tk.Tk()
root.configure(bg='#31393C')
#window = Canvas(root, width=100, height=100,  bg='#31393C').grid(row=0, column=0)

# URL Input Entryd
Label(root, text='Enter job description URL:', font=("Helvetica", 16, 'bold'),
      fg='#94DDBC', bg='#31393C').grid(row=0, column=0)
urlBox = tk.Entry(root, text='Enter URL').grid(row=0, column=1)
submit = tk.Button(root, text='Submit').grid(row=0, column=2)

# Company Name
Label(root, text='Company:', font=("Helvetica", 10, 'bold'),  bg='#31393C', fg='#94DDBC').grid(row=1, column=0)
company = tk.Entry(root).grid(row=1, column=1)

# Role Title
Label(root, text='Position Title:', font=("Helvetica", 10, 'bold'),  bg='#31393C', fg='#94DDBC').grid(row=2, column=0)
title = tk.Entry(root).grid(row=2, column=1)

# Pay
Label(root, text='Pay:', font=("Helvetica", 10, 'bold'),  bg='#31393C', fg='#94DDBC').grid(row=3, column=0)
pay = tk.Entry(root).grid(row=3, column=1)

# Link to advert
Label(root, text='Advert Link:', font=("Helvetica", 10, 'bold'),  bg='#31393C' , fg='#94DDBC').grid(row=4, column=0)
link = Entry(root).grid(row=4, column=1)

# Date Applied
Label(root, text='Date Applied:', font=("Helvetica", 10, 'bold'), fg='#94DDBC', bg='#31393C').grid(row=5, column=0)
dateApplied = Entry(root).grid(row=5, column=1)

# Location
Label(root, text='Location:', font=("Helvetica", 10, 'bold'),  bg='#31393C' , fg='#94DDBC').grid(row=6, column=0)
location = tk.Entry(root).grid(row=6, column=1)

root.mainloop()
# https://jobs.thomsonreuters.com/ShowJob/Id/353964/Software-Engineer-Internship-Summer-2020/
# https://jobs.jobvite.com/code42/job/oYc1afw5?__jvst=Career+Site
# https://jobs.lever.co/smartthings/88cf7759-c792-4d20-8cc9-b752f883bd03


# Pulls HTML
def pullInfo(url):
    source = requests.get(url)
    text = source.text
    soup = BeautifulSoup(text, features="html.parser")

    # Grabs all information and stores into master list
    info = []
    for item in soup.find_all('title'):
        info.append(item.string)

    for item in soup.find_all('h1'):
        info.append(item.string)

    for item in soup.find_all('h2'):
        info.append(item.string)

    # Images have a separate list for company name
    imgs = soup.find_all('img')
    composs = []
    if len(composs) > 0:
        for image in imgs:
            composs.append(image['alt'])
    else:
        composs.append("Could Not Find")

    # Removes any potential None values
    clean = []
    for val in info:
        if val is not None:
            clean.append(val)
    print(clean)

    # Initialize variables
    company = str(composs[0]).strip('logo')
    title = ""
    date = datetime.datetime.now().strftime('%d' + '/' + '%m' + '/' + '%Y')
    location = "".strip('')


    # Sorts through info and stores information into respected variables
    for i in range(len(clean)):
        if "Internship" in clean[i] or "Intern" in clean[i]:
            title = str(clean[i]).strip(' ')
            break

    return company, title, url, date, location


# Function to add internship info to spreadsheet
def updateXl(comp, role, link, date, location):
    # Opens spreadsheet
    file = 'Internship Tracker.xlsx'
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    # Style formatting
    bg = PatternFill(start_color='595959',
                          end_color='595959',
                          fill_type='solid')

    border = Border(
                    right=Side(border_style='thin'),
                    top=Side(border_style='thin'),
                    bottom=Side(border_style='thin'))

    font = Font(color='ffffff')

    # Imports data into cells
    ws.insert_rows(3)
    ws['A3'] = comp
    ws['A3'].fill = bg
    ws['B3'] = role
    ws['C3'] = 'N/A'
    ws['D3'] = link
    ws['E3'] = date
    ws['F3'] = location
    ws['G3'] = ''
    ws['H3'] = ''
    ws['I3'] = ''
    ws['J3'] = ''

    # Applying style formats to cells
    ws['C3'].alignment = Alignment(horizontal='center')
    ws['E3'].alignment = Alignment(horizontal='right')
    ws['E3'].font = Font(bold=True)
    ws['A3'].fill = ws['B3'].fill = ws['C3'].fill = ws['D3'].fill = ws['E3'].fill = ws['F3'].fill = ws['G3'].fill\
        = ws['H3'].fill = ws['I3'].fill = ws['J3'].fill = bg
    ws['A3'].border = ws['B3'].border = ws['C3'].border = ws['D3'].border = ws['E3'].border = ws['F3'].border\
        = ws['G3'].border = ws['H3'].border = ws['I3'].border = ws['J3'].border = border
    ws['A3'].font = ws['B3'].font = ws['C3'].font = ws['D3'].font = ws['E3'].font = ws['F3'].font \
        = ws['G3'].font = ws['H3'].font = ws['I3'].font = ws['J3'].font = font

    wb.save(file)

root.mainloop()
#url = str(input("Enter job page URL: ")).strip(' ')
company, title, url, date, location = pullInfo(url)
updateXl(company, title, url, date, location)
