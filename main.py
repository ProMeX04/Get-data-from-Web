import subprocess
import requests
import pandas as pd
from io import StringIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
import subprocess
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from ttkthemes import ThemedTk
import os

# Constant Variable
#create list lesson 
LESSON = (29, 20, 27, 0, 7, 17, 0, 6, 26, 13, 5, 10, 19, 0, 20, 6, 6, 10, 0, 6, 0, 0, 11, 0, 5)
TOTAL = 258

# list admin
LIST_ADMIN = ["laihieu2714", "Tran_Nhi0306","superadmin", "Ninhkfc72", "Nguyen_Thi_Lua_FullHouse"]


def login(username, password) -> str:
    """Login and retrieve session ID"""
    session = requests.Session()
    
    # Get CSRF token
    response = session.get("https://laptrinh24h.vn/accounts/login/?next=" )
    csrf_token = response.cookies.get("csrftoken")

    payload = {
        "username": username,
        "password": password,
        "csrfmiddlewaretoken": csrf_token
    }
    session.post("https://laptrinh24h.vn/accounts/login/?next=" , data=payload)
    
    session_id = session.cookies.get("sessionid")
    return session_id


def request(url, username, password) -> requests.Response:
    """Get response html web"""
    session = requests.Session()
    session.cookies['sessionid'] = login(username, password)
    response = session.get(url)
    return response


def getDataFrame(response, CurrentLesson) -> pd.DataFrame:
    """
    Parses the response text and returns a pandas DataFrame with the extracted data.

    Parameters:
    response (requests.Response): The response object containing the HTML data.

    Returns:
    pandas.DataFrame: The DataFrame containing the parsed data.
    """
    df = pd.read_html(StringIO(response.text))[0]
    # Rest of the code...
    
    df = df.drop(columns=["Organization"])
    df = df.drop(columns=["Points"])

    # Delete column Organization
    for i in range(1, TOTAL-1):
        head = str(i) + "  10"
        df[head] = df[head].astype(str).str[0:2]

    new_df = pd.DataFrame()
    so_hoc_vien = df.shape[0]

    for i in range(so_hoc_vien):
        start_lesson, total = 2, 0
        row = [i+1, df.iloc[i, 1].split()[0]]

        if row[1] in LIST_ADMIN:
            continue

        for index, lesson in enumerate(LESSON, start=1):
            accepted = (
                df.iloc[i, start_lesson: start_lesson + lesson] == "10").sum()
            total += accepted
            if accepted:
                row.append(f"{accepted}/{lesson}")
            else:
                if index < CurrentLesson and lesson != 0:
                    row.append(f"{accepted}/{lesson}")
                else:
                    row.append("")

            start_lesson += lesson
        row.append(f"{total}/{TOTAL}")
        new_df = pd.concat(
            [new_df, pd.DataFrame([row])], ignore_index=True)

    new_df = new_df.rename(columns={0: "Rank", 1: "Name", 27: "Total"})

    return new_df


def format_cell(dataFrame, ExcelPath) -> None:
    """
    Format cell in excel.

    Args:
        dataFrame (pandas.DataFrame): The DataFrame containing the data to be formatted.

    Returns:
        None
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    for row in dataframe_to_rows(dataFrame, index=False, header=True):
        sheet.append(row)

    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['A'].width = 5

    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    custom_font = Font(name='Calibri', size=10, bold=True,
                       italic=True, color='000000')
    custom_alignment = Alignment(
        horizontal='center', vertical='center', wrap_text=False)

    fill_green = PatternFill(start_color='00FF00',
                             end_color='00FF00', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFFF00',
                              end_color='FFFF00', fill_type='solid')
    fill_red = PatternFill(start_color='FF0000',
                           end_color='FF0000', fill_type='solid')

    for row in sheet.iter_rows():
        for cell in row:
            # Rest of the code...
            cell.border = thin_border
            cell.font = custom_font
            cell.alignment = custom_alignment
            if cell.value:
                if len(str(cell.value).split('/')) == 2:
                    # Rest of the code...
                    pass
                    left, right = cell.value.split('/')
                    if int(left) >= int(right):
                        cell.fill = fill_green
                    elif int(left) >= int(right)//2:
                        cell.fill = fill_yellow
                    else:
                        cell.fill = fill_red

    workbook.save(ExcelPath)


def main():
    """main"""
    file_path = os.path.dirname(os.path.abspath(__file__))
    
    window = ThemedTk(theme = "material")
    window.iconbitmap(os.path.join(file_path, 'ico.ico'))
    window.title("GET DATA")
    window.geometry("380x180")
    window.resizable(False, False)
    
    
    
    Username = tk.StringVar()
    Password = tk.StringVar()
    Class = tk.StringVar()
    ExcelPath = tk.StringVar()
    CurrentLesson = tk.IntVar()   
    Status = tk.StringVar()
    
    
    ttk.Label(window, text="Username ",font="calibri",foreground = "Green").grid(row=0, column=0)
    ttk.Label(window, text="Password ",font="calibri",foreground = "Green").grid(row=1, column=0)
    ttk.Label(window, text="Class ",font="calibri",foreground = "Green").grid(row=2, column=0)
    ttk.Label(window, text="Current lesson ",font="calibri",foreground = "Green").grid(row=3, column=0)
    ttk.Label(window, text="Path ",font="calibri",foreground = "Green").grid(row=4, column=0)
    ttk.Label(window, textvariable=Status, foreground="RED", font= "consolas").grid(row=6, columnspan=3)
     

    with open(os.path.join(file_path, 'user.txt'), "r") as file:
        lines = file.readlines()
        if len(lines) >= 5:
            Username.set(lines[0].strip())
            Password.set(lines[1].strip())
            Class.set(lines[2].strip())
            CurrentLesson.set(int(lines[3].strip()))
            ExcelPath.set(lines[4].strip())
        else:
            Username.set("")
            Password.set("")
            Class.set("")
            CurrentLesson.set(0)
            ExcelPath.set("")
        

    ttk.Entry(window, textvariable=Username , font = "consolas").grid(row=0, column=1)
    ttk.Entry(window, textvariable=Password, show="*" , font="consolas").grid(row=1, column=1)
    ttk.Entry(window, textvariable=Class , font="consolas").grid(row=2, column=1)
    ttk.Entry(window, textvariable=CurrentLesson , font="consolas").grid(row=3, column=1)
    ttk.Entry(window, textvariable=ExcelPath, font="consolas").grid(row=4, column=1)
    
    
    def browse_folder():
        """Browse for Excel path"""
        path = filedialog.askdirectory()
        ExcelPath.set(path+"/data.xlsx")
    
    
    def button():
        """Button"""
        
        with open(os.path.join(file_path, 'user.txt'), "w") as file:
            file.write(Username.get() + "\n" + Password.get() + "\n" + Class.get() + "\n" + str(CurrentLesson.get()) + "\n" + ExcelPath.get())
            
        response = request("https://laptrinh24h.vn/contest/cpp{}/ranking/".format(Class.get()), Username.get(), Password.get())
        if response.status_code == 200:
            Status.set("Login Success")
            dataframe = getDataFrame(response , CurrentLesson.get())
            format_cell(dataframe, ExcelPath.get())
            subprocess.Popen(["start", ExcelPath.get()], shell=True)
        else:
            Status.set("Username or Password is wrong")
    
    
    def focus_next_widget(event):
        """Focus the next widget"""
        event.widget.tk_focusNext().focus()
        return("break")


    def focus_previous_widget(event):
        """Focus the previous widget"""
        event.widget.tk_focusPrev().focus()
        return("break")


    # Bind the keys to the focus change functions
    window.bind('<Down>', focus_next_widget)
    window.bind('<Up>', focus_previous_widget)
    ttk.Button(window, text="Browse", command=browse_folder).grid(row=4, column=2)
    ttk.Button(window, text="GET",command = button ).grid(row=5, column=1)

    # Create a button with the new style
    button = ttk.Button(window, text="GET", command=button, style='TButton')
    button.grid(row=5, column=1)


    window.mainloop()

main()

