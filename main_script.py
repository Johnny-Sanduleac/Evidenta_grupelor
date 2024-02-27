# -*- coding: utf-8 -*-
"""
Created on Tue Jan 23 21:28:11 2024

@author: user
""" 
# Imports

# 1. Base libraries
import openpyxl, sys, re, os, git
import tkinter as tk
from tkinter import scrolledtext, Menu, ttk
from tkinter.filedialog import askopenfilename

# 2. Server communication libraries
import smtplib # libraria pentru trimitere e-mail
from email.mime.multipart import MIMEMultipart # libraria pentru formatarea mesajului
from email.mime.text import MIMEText # pentru textul scrisorii

""" ************************** Main Functions ******************************"""


""" ************************** Back END  ***********************************"""
def connect_server(from_addr, password):
    """ Functia care stabileste conexiunea cu serverul"""
    global server
    # Vom verifica ce fel de posta utilizeaza user-ul
    tail = from_addr.split("@")[-1] # coada adresei (yahoo.com, gmail.com, mt.utm.md, mail.ru)
    flag = None
    if tail == "yahoo.com":
        server_adress = 'smtp.mail.yahoo.com'; server_port = 465
        flag = 1
        
    elif tail == "gmail.com":
        server_adress = 'smtp.gmail.com'; server_port = 465
        flag = 1
        
    elif ".".join(tail.split(".")[1:3]) == "utm.md":
        server_adress = 'smtp-mail.outlook.com'; server_port = 587
        flag = 2
    
    elif tail == "mail.ru":
        server_adress = 'smtp.mail.ru'; server_port = 465
        flag = 1

    else:
        info_msg = "Your email is not available now, please contact the developper"
        popup_msg(info_msg)
        flag = 0
        
    server = None
    if flag == 1: # Pentru alte servere de mail
        try:
            server = smtplib.SMTP_SSL(server_adress, server_port)
        except:
            popup_msg("Failed to connect server. Check your login or internet connection")
    if flag == 2: # Pentru serverul outlook
        try:
            server = smtplib.SMTP(server_adress, server_port)
        except:
            popup_msg("Failed to connect server. Check your internet connection")
            quit
        try:
            server.ehlo() # Trimitem un fel de hello serverului, sa vedem daca raspunde
        except:
            popup_msg("Server not responding ")
            quit
        try:
            server.starttls()
        except:
            popup_msg("Server not responding ")
            quit
    # Daca am reusit conexiunea cu serverul, atunci mergem mai departe
    if server:
        try:
            server.login(from_addr, password) # Ne logam pe account-ul nostru
            b2.configure(bg='light blue',font=fnt, text = "Connected...")
            b4['state'] = 'normal'; b4['bg'] = 'yellow'
            b5['state'] = 'normal'; b5['bg'] = 'green'
        except:
            popup_msg("Failed to login. Check your login and password")
    


def open_excel(path_to_excel):
    try:
        excel_obj = openpyxl.load_workbook(path_to_excel, read_only=True, data_only=True)
    except:
        popup_msg("Unable to read excel file. \ Please close the program and try again")
    # Then, return the excel_object
    return excel_obj


def read_excel(excel_obj, sheet_name):
    # Extract first sheet (where names and emails are stored)
    first_sheet = excel_obj.worksheets[0]
    # Get maximum number of rows with data in first sheet (all the rest will have the same max number of rows)
    rows = 0
    for max_row, row in enumerate(first_sheet,1):
        if not all(col.value is None for col in row):
            rows +=1
    # Extract email and position from the first sheet. A dictoinary with key-values. Keys are names, values are list[email, info]
    name_email_info_dict = {}
    # Declare a regex expression for emails
    regex = re.compile(r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])") 
    for row in range(rows):
        # extragem valoarea din coloana C (care trebuie sa contina nume)
        #verificam daca este ceva in celula
        if first_sheet[f"C{row}"].value:
            student_name = str(first_sheet[f"C{row}"].value).strip()
            # verificam daca e-mail-ul corespunde regex
            if first_sheet[f"D{row}"].value and re.fullmatch(regex, first_sheet[f"D{row}"].value.strip()):
                student_email = str(first_sheet[f"D{row}"].value).strip()
                # Si adaugam la dictionar
                name_email_info_dict.update({student_name: [student_email]})
    # Now, we'll take the name chosen in combobox and we'll extract info from that sheet
    #sheet_name = sheets_combo.get()
    sheet_name = "Examen"
    sheet = excel_obj[sheet_name]
    # Now, we'll collect all the information 
    for row in range(rows):
        # extragem valoarea din coloana B (care trebuie sa contina nume)
        cell_val = str(sheet[f"B{row}"].value)
        # Iteram prin cheile dictionarului si daca gasim in sheet-ul dat numele din dictionar
        for name in name_email_info_dict.keys():
            if name.strip() == cell_val.strip():
                # Daca am gasit numele studentului, incepem formarea mesajului
                # Daca sunt indicate coloanele din care se extrage nota
                if e2.get():
                    punctaj_string = ""
                    for i in e2.get().split(','):
                        antet_punctaj = f"{sheet[f'{i.strip()}1'].value}".replace("_x000D_", " ")
                        punctaj = f"{sheet[f'{i.strip()}{row}'].value}"
                        if punctaj != "None":
                            punctaj_string += f"<br>{antet_punctaj} -- {punctaj} puncte "
                        else:
                            punctaj_string += f"<br>{antet_punctaj} -- 0 puncte "
                    
                if e3.get():
                    nota_string = ""
                    for i in e3.get().split(','):
                        antet_nota = f"{sheet[f'{i.strip()}1'].value}"
                        nota = f"{sheet[f'{i.strip()}{row}'].value}"
                        nota_string += (f"<br>{antet_nota} -- {nota}<br>" if nota != "None" else f"<br>{antet_nota} -- n/p<br>")
                        
                # In final, formam un string student_info, pregatit pentru html
                student_info = "".join(punctaj_string + nota_string)
                name_email_info_dict[name].append(student_info)
    return name_email_info_dict    
   
def format_message_content(student_email, student_info):
    """ Functia care formeaza continutul e-mail-ului"""
    # Mesajul inclus in email:
    msg = MIMEMultipart() # cream obiectul msg
    # 1. Antetetul scrisorii
    msg['Subject'] = e6.get()
    msg['From'] = e4.get()
    msg['To'] = student_email
    content = "<br> ".join(e7.get('1.0','end-1c').split('\n')) + '<br>'
    content += student_info
    body = MIMEText(content, 'html') # creaza obiectul html - body
    msg.attach(body)  
    return msg

""" ************************** GUI  ********************************"""

""" GUI-linked functions"""
def browse_excel():
    global excel_obj
    global path_to_excel
    path_to_excel = askopenfilename(title="Select a File",\
                                 filetype=(("Excel", "*.xlsx"), ("Excel", "*.xls"))) 
    # Transmitem path-ul catre open_excel
    excel_obj = open_excel(path_to_excel)
    # If Excel file was read successfully, read all sheet names
    sheet_names = excel_obj.sheetnames
    # And insert these names in combobox
    c0['values'] = sheet_names

def combo_command(event):
    # Get sheet name from combo
    global name_email_info_dict
    sheet_name = c0.get()
    name_email_info_dict = read_excel(excel_obj, sheet_name)
    
    
    
def preview():
    # Now, we have excel_obj, path_to_excel and sheet_name
    with open(f"{os.path.dirname(path_to_excel)}\\{os.path.basename(path_to_excel).split('.')[:-1][0]}_Preview.html", mode = 'w', encoding="utf-8") as file:
        file.write('Informatia care va fi expediata la fiecare student: <br>')
        file.write(f"Subiect: {str(e6.get())}  <br>")
        content = "<br> ".join(e7.get('1.0','end-1c').split('\n'))
        file.write(f"Continut: <br> {content} <br>")
        student_number = 0
        for student in name_email_info_dict.keys():
            student_number += 1
            file.write(f"<br>{student_number}. {student}\n <br>")
            file.write(name_email_info_dict[student][0])
            file.write(name_email_info_dict[student][1])
    popup_msg("Verificati in directoriul excel-ului, s-a generat un fisier html")
    
    
def send_mail_to_yourself():
    """Functia pentru a expedia mesaj pe propriul account"""
    # Luam prima cheie din dictionar - primul student
    first_key = list(name_email_info_dict.keys())[0]
    student_info = name_email_info_dict[first_key][1]
    msg = format_message_content(e4.get(), student_info)
    try:
        server.sendmail(e4.get(), e4.get(), msg.as_string()) # Trimitem e-mail de pe account-ul nostru
        popup_msg("Expediat cu succes pe adresa Dvs.")
    except:
        popup_msg("Connection with mail server lost")

def send_mail_to_recipients():
    """Functia pentru a expedia mesaje la fiecare email din excel"""
    for name in name_email_info_dict.keys():
        student_email = name_email_info_dict[name][0]
        student_info = name_email_info_dict[name][1]
        msg = format_message_content(student_email, student_info)
        try:
            server.sendmail(e4.get(), student_email, msg.as_string()) # Trimitem e-mail de pe account-ul nostru
            print(f"Successfully sent to {name}")
        except:
            print(f"Message not sent to {name}")
            pass
    popup_msg("Finished sending emails...")
            
def exit_app():
    # Garbage collector
    # Close all connections
    try:
        if excel_obj:
            excel_obj.close()
    except:
        pass
    for obj in dir():
        del globals()[obj]
    
    root.destroy()
    
def popup_msg(message):
    popup = tk.Tk()
    popup.wm_title("Info")
    label = tk.Label(popup, text=message, fg = 'blue' )
    label.pack(side="top", fill="x", pady=20)
    label.config(font=("Times New Roman", 14))
    B1 = tk.Button(popup, text="Ok", padx = 30, pady = 5, borderwidth = 5, \
                   command = popup.destroy)
    B1.pack()   

def check_for_updates():
    parent_path = os.path.abspath(os.path.join(sys.executable, os.pardir))
    repo_path = os.path.abspath(os.path.join(parent_path,"MyScripts\\Evidenta_grupelor" ))
    repo = git.Repo(repo_path)
    repo.remotes.origin.fetch()
    diff = repo.git.diff('origin/main')
    if len(diff) !=0:
        print('diff = ', diff)
        popup_msg("New updates are avaialble!")
    
def update():
    parent_path = os.path.abspath(os.path.join(sys.executable, os.pardir))
    repo_path = os.path.abspath(os.path.join(parent_path,"MyScripts\\Evidenta_grupelor" ))
    repo = git.Repo(repo_path)
    repo.git.reset('--hard','origin/main')
    origin = repo.remote(name='origin') 
    origin.pull()
    popup_msg("Successfully updated. Please restart the program!")


""" GUI graphical elements """
root = tk.Tk()
root.title("Email Agent")
root.geometry("900x650")

# When start GUI, we'll check for updates
check_for_updates()


# Font specifications
fnt = ("Arial", 12, "bold")

# Menu
menubar = Menu(root) # generam un obiect menubar in root
filemenu = Menu(menubar, tearoff=0) # Definim obiectul pentru File
filemenu.add_command(label="New", command=lambda: popup_msg('under process...')) # Adaugam comanda New
filemenu.add_separator()
filemenu.add_command(label = 'Exit', command = exit_app)
menubar.add_cascade(label="File", menu=filemenu) #In bara de meniuri includem obiectul filemenu

editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Delete", command=lambda: popup_msg('under process...'))
menubar.add_cascade(label="Edit", menu=editmenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Update", command=update)
helpmenu.add_command(label="About...", command=lambda: popup_msg('A simple application for sending emails to students \n with marks and points'))
menubar.add_cascade(label="Help", menu=helpmenu)

root.config(menu = menubar) # Configuram root-ul ca sa stie ca in menu avem obiectul menubar

# Main widgets

l0 = tk.Label(root, text = "Expedierea notelor si altor informatii catre studenti ", font = fnt, pady = 10)

l1 = tk.Label(root, text = "Excel-ul cu note: ",font=fnt, pady = 10, padx = 10)
b1 = tk.Button(root, text = "Browse ", fg = 'black', font=fnt,\
              padx = 30, pady = 5, borderwidth = 5,\
              bg = 'light blue',command = browse_excel)
# Combobox
c0 = ttk.Combobox(root, state="readonly", width = 40, height=20, font = fnt)
c0.bind("<<ComboboxSelected>>", combo_command)

l2 = tk.Label(root, text = "Pagina cu note: ", font=fnt, pady = 10, padx = 10)

l3 = tk.Label(root, text = "Coloanele (coloana) cu punctaj\n separate prin virgulă (optional)",font=fnt, pady = 10, padx = 10)
e2 = tk.Entry(root,  width = 50, borderwidth = 5, font=fnt)
e2.insert(tk.END,'C, D, E, F, G, H, I')


l4 = tk.Label(root, text = "Coloanele (coloana) cu note\n separate prin virgulă",font=fnt, pady = 10)
e3 = tk.Entry(root,  width = 50, borderwidth = 5, font=fnt)
e3.insert(tk.END,'K')

    
l5 = tk.Label(root, text = "Adresa expeditorului", font=fnt, pady = 10)
e4 = tk.Entry(root,  width = 50, borderwidth = 5, font=fnt)
e4.insert(0,'your_email')


l6 = tk.Label(root, text = "Parola expeditorului", font=fnt, pady = 10)
e5 = tk.Entry(root,  width = 50, borderwidth = 5, font=fnt)
e5.insert(0,'app_password' )

l7 = tk.Label(root, text = "Subiect ", font=fnt, pady = 10)
e6 = tk.Entry(root,  width = 50, borderwidth = 5, font=fnt)
# Temporar, inseram ceva content
e6.insert(0,'Rezultatele evaluării Nr....')
    
l8 = tk.Label(root, text = "Conținutul mesajului \n (un text opțional, \n de salut, semnătura etc.) ", font=fnt)
e7 = scrolledtext.ScrolledText(root, wrap = tk.WORD, height = 12, width = 48,  borderwidth = 5, font=fnt)

# Temporar, inseram ceva content
e7.insert('1.0','Bună ziua, \nmai jos găsiți notele de la evaluarea ...')

b2 = tk.Button(root, text = "Check \nserver \nconnection ", fg = 'black', font=fnt,\
              padx = 10, pady = 5, borderwidth = 5,\
              bg = 'light blue',command = lambda:connect_server(e4.get(), e5.get()))
    
b3 = tk.Button(root, text = "Preview", fg = 'black', font=fnt,\
              padx = 25, pady = 8, borderwidth = 5,\
              bg = 'yellow',command = preview)
    
b4 = tk.Button(root, state = 'disabled', text = "Send \ntest email", fg = 'black', font=fnt,\
              padx = 20, pady = 8, borderwidth = 5,\
              bg = 'gray',command = send_mail_to_yourself)


b5 = tk.Button(root, state = 'disabled', text = "SEND",fg = 'black',font=fnt,\
               padx = 35, pady = 20, borderwidth = 5,\
               bg = 'gray',command = send_mail_to_recipients)
    
b6 = tk.Button(root, text = "EXIT", fg = 'black', font=fnt,\
              padx = 40, pady = 10, borderwidth = 5,\
              bg = 'red',command = exit_app)
    
    
 
"""************* Window layout **************"""
    
r = 0
l0.grid(sticky = "W", row = r, column = 0, columnspan = 5)

r+=1
l1.grid(row = r, column = 0, columnspan = 2)
b1.grid(sticky = "W",row = r, column = 2, columnspan = 2, rowspan = 1)

r+=1
l2.grid(row = r, column = 0, columnspan = 2)
c0.grid(sticky = "W",row = r, column = 2, columnspan = 4)

r+=1
l3.grid(row = r, column = 0, columnspan = 2)
e2.grid(row = r, column = 2, columnspan = 4)
b2.grid(row = r, column = 6, columnspan = 2, rowspan = 2)

r+=1
l4.grid(row = r, column = 0, columnspan = 2)
e3.grid(row = r, column = 2, columnspan = 4)

r+=1
l5.grid(row = r, column = 0, columnspan = 2)
e4.grid(row = r, column = 2, columnspan = 4)
b3.grid(row = r, column = 6, columnspan = 2)


r+=1
l6.grid(row = r, column = 0, columnspan = 2)
e5.grid(row = r, column = 2, columnspan = 4)
b4.grid( row = r, column = 6, columnspan = 2, rowspan = 2)

r+=1
l7.grid(row = r, column = 0, columnspan = 2)
e6.grid(row = r, column = 2, columnspan = 4)



r+=1
l8.grid(row = r, column = 0, columnspan = 2)
e7.grid(row = r, column = 2, columnspan = 4, rowspan = 6)
b5.grid(row = r, column = 6, columnspan = 2, rowspan = 2)

r+=4
b6.grid(row = r, column = 6, columnspan = 2, rowspan = 2)

root.mainloop()
