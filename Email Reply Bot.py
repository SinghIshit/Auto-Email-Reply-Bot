import os
import smtplib
import ssl
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import email
from email.message import EmailMessage
from email.utils import formataddr
from langchain.chat_models import ChatOpenAI
from langchain import PromptTemplate, LLMChain
import ttkbootstrap as ttk
from tkinter import font as tkfont
import imaplib
from email.header import decode_header
import json


sender = ""
password = ""
sender_info_window = None
main_app_window =None

def center(win):
    """
    centers a tkinter window
    :param win: the main window or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

def get_sender_info_unread():
    global sender, password, sender_info_window, recent_email_window, options_window
    options_window.destroy()
    open_main_application()

def get_sender_info_recent():
    global sender, password, sender_info_window, recent_email_window, options_window
    recent_email_window.destroy()
    open_main_application()

def on_hover(event): 
    event.widget['background'] = '#C85855'  

def on_leave(event):
    event.widget['background'] = '#EB6864'

def open_main_application():
    global sender, password, sender_info_window, main_app_window
    os.environ["OPENAI_API_KEY"] = #Input OpenAI API Key
    llm = ChatOpenAI(model="gpt-3.5-turbo")

    def send_email():
        global receiver, sender, password
        
        selected_row = tree.selection()
        if not selected_row:
            messagebox.showinfo("Error", "Please select a row.")
            return

        generated_reply = generated_reply_text.get("1.0", tk.END)

        template = """Write a good subject for the following email: {mail}, don't put "Subject:" in front of it."""
        prompt = PromptTemplate(template=template, input_variables=["mail"])
        chain = LLMChain(llm=llm, prompt=prompt)
        subject = chain.run(generated_reply)

        em = EmailMessage()
        em['From'] = formataddr(("Automating", f"{sender}"))
        em['To'] = receiver
        em['Subject'] = subject
        em.set_content(generated_reply)
        send_approval = messagebox.askquestion("Confirmation", "Do you want to send this email?")
        if send_approval == "yes":
            context = ssl.create_default_context()
            try:
                with smtplib.SMTP('smtp.gmail.com', 587) as server:
                    server.starttls(context=context)
                    server.login(sender, password)
                    server.sendmail(sender, receiver, em.as_string())
                messagebox.showinfo("Success", "Email sent successfully!")
                received_email_text.delete("1.0", tk.END)
                generated_reply_text.delete("1.0", tk.END)
            except Exception as e:
                messagebox.showerror("ErroAr", f"An error occurred: {e}")
        else:
            messagebox.showinfo("Info", "Email not sent.")


    def on_tree_select(event):
        selected_row = tree.selection()
        if selected_row:
            selected_item = tree.item(selected_row)
            oldbody = selected_item['values'][1]
            received_email_text.delete("1.0", tk.END)
            received_email_text.insert(tk.END, oldbody)

    def generate_email():

        global receiver
        selected_row = tree.selection()
        if not selected_row:
            messagebox.showinfo("Error", "Please select a row.")
            return

        selected_item = tree.item(selected_row)
        receiver = selected_item['values'][0]
        edited_body = received_email_text.get("1.0", tk.END)

        template = """reply to this email as the person who this mail is meant for:{mail}."""
        prompt = PromptTemplate(template=template, input_variables=["mail"])
        chain = LLMChain(llm=llm, prompt=prompt)
        body = chain.run(edited_body)

        generated_reply_text.delete("1.0", tk.END)
        generated_reply_text.insert(tk.END, body)


    def frame_hover(event):
        event.widget['highlightbackground'] = '#EB6864'
        event.widget['highlightthickness'] = 2

    def frame_leave(event):
        event.widget['highlightbackground'] = 'black'
        event.widget['highlightthickness'] = 1

    def tree_hover(event):
        event.widget['highlightbackground'] = '#EB6864'
        event.widget['highlightthickness'] = 1

    def logout():
        global sender, password, sender_info_window, main_app_window
        sender = ""
        password = ""
        sender_info_window = None
        main_app_window.destroy()
        open_login_window()

    main_app_window = ttk.Window(themename='journal')
    main_app_window.title("Email Automation")
    main_app_window.geometry("+650+0")

    font_size = 11 

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(size=font_size)

    logout_frame = tk.Frame(main_app_window)
    logout_frame.pack(side="top", anchor="nw", padx=20, pady=10)
    logout_button = tk.Button(logout_frame, text="Logout", command=logout, cursor="hand2", width=7, height=1)
    logout_button.pack(side="top", padx=20, pady=10)
    logout_button.bind("<Enter>", on_hover)
    logout_button.bind("<Leave>", on_leave)

    tree_frame = tk.Frame(main_app_window,highlightbackground="black", highlightthickness=1)
    tree_frame.pack(fill="x", expand=True)
    tree = ttk.Treeview(tree_frame, columns=("Receiver", "Content"), show="headings", style="Treeview")
    tree.heading("Receiver", text="Receiver")
    tree.heading("Content", text="Content")
    tree.pack(fill="both", expand=True)  
    tree_frame.bind("<Enter>", tree_hover)
    tree_frame.bind("<Leave>", frame_leave)


    df = pd.read_excel("email_data1.xlsx")
    for index, row in df.iterrows():
        tree.insert("", "end", values=(row["EMAIL"], row["CONTENT"]))

    tree.bind("<<TreeviewSelect>>", on_tree_select)

    received_email_frame = tk.Frame(main_app_window, highlightbackground="black", highlightthickness=1)
    received_email_frame.pack(fill="x", expand=True)
    received_email_text = tk.Text(received_email_frame, height=10, width=60, font=default_font)  
    received_email_text.pack(fill="x", expand=True)
    received_email_frame.bind("<Enter>", frame_hover)
    received_email_frame.bind("<Leave>", frame_leave)

    generated_reply_frame = tk.Frame(main_app_window, highlightbackground="black", highlightthickness=1)
    generated_reply_frame.pack(fill="x", expand=True)
    generated_reply_text = tk.Text(generated_reply_frame, height=15, width=60, font=default_font)  
    generated_reply_text.pack(fill="x", expand=True)
    generated_reply_frame.bind("<Enter>", frame_hover)
    generated_reply_frame.bind("<Leave>", frame_leave)


    button_frame = tk.Frame(main_app_window, highlightbackground="black", highlightthickness=1)
    button_frame.pack(fill="both", expand=True)

    generate_button_frame = tk.Frame(button_frame)
    generate_button_frame.pack(side="left",fill="x", expand=True)
    generate_button = tk.Button(generate_button_frame, text="Generate Reply", command=generate_email, cursor="hand2")
    generate_button.pack(side="right",padx=20,pady=10)
    generate_button.bind("<Enter>", on_hover)
    generate_button.bind("<Leave>", on_leave)


    send_button_frame = tk.Frame(button_frame)
    send_button_frame.pack(side="left",fill="x", expand=True)
    send_button = tk.Button(send_button_frame, text="Send Email", command=send_email, cursor="hand2")
    send_button.pack(side="left", padx=20,pady=10)
    send_button.bind("<Enter>", on_hover)
    send_button.bind("<Leave>", on_leave)

    button_frame.bind("<Enter>", frame_hover)
    button_frame.bind("<Leave>", frame_leave)

    main_app_window.mainloop()



def open_login_window():
    global sender_info_window, sender_entry, password_entry, show_password_var
    login_data = load_login_data()
    sender_info_window = ttk.Window(themename='journal')
    sender_info_window.title("Login")
    sender_info_window.geometry("550x350")
    center(sender_info_window)
    font_size = 12

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(size=font_size)

    email_frame = tk.Frame(sender_info_window)
    email_frame.pack(pady=10, padx=(0, 26))

    sender_label = tk.Label(email_frame, text="Email:")
    sender_label.pack(side="left", padx=5)

    sender_entry = tk.Entry(email_frame)
    sender_entry.pack(side="left", padx=5)

    password_frame = tk.Frame(sender_info_window)
    password_frame.pack(pady=10, padx=(0, 20), anchor="center")

    password_label = tk.Label(password_frame, text="Password:")
    password_label.pack(side="left", padx=5)

    password_entry = tk.Entry(password_frame, show="*")
    password_entry.pack(side="left", padx=5)

    show_password_var = tk.BooleanVar()
    show_password_var.set(False)

    def toggle_password_visibility():
        if show_password_var.get():
            password_entry.config(show="")
        else:
            password_entry.config(show="*")

    show_password_button = tk.Checkbutton(password_frame, text="", variable=show_password_var,command=toggle_password_visibility)
    show_password_button.pack(side="left")

    remember_me_var = tk.BooleanVar()
    remember_me_check = tk.Checkbutton(sender_info_window, text="Remember Me", variable=remember_me_var)
    remember_me_check.pack(pady=5)

    if "email" in login_data and "password" in login_data:
        sender_entry.insert(0, login_data["email"])
        password_entry.insert(0, login_data["password"])
        remember_me_var.set(True)

    def validate_and_login():
        email = sender_entry.get().strip()
        password = password_entry.get().strip()
        
        if remember_me_var.get():
                save_login_data(email, password)

        try:
            context = ssl.create_default_context()
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls(context=context)
            server.login(email, password)

            server.quit()
            open_options_window()

        except Exception as e:
            messagebox.showerror("Error", f"Login failed: {str(e)}")

    submit_button = tk.Button(sender_info_window, text="Login", command=validate_and_login,width=15)
    submit_button.pack(pady=20)
    submit_button.bind("<Enter>", on_hover)
    submit_button.bind("<Leave>", on_leave)
    sender_info_window.mainloop()

def load_login_data():
    try:
        with open("login_data.json", "r") as file:
            login_data = json.load(file)
    except FileNotFoundError:
        login_data = {}
    return login_data

def save_login_data(email, password):
    login_data = {"email": email, "password": password}
    with open("login_data.json", "w") as file:
        json.dump(login_data, file)

def open_options_window():
    global options_window, sender, password
    sender = sender_entry.get()
    password = password_entry.get()
    sender_info_window.destroy()   

    options_window = ttk.Window(themename='journal')
    options_window.title("Options")
    options_window.geometry("550x350")
    
    center(options_window)

    font_size = 12

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(size=font_size)

    extract_unread_button = tk.Button(options_window, text="Extract Unread Email", command=extract_unread_email)
    extract_unread_button.bind("<Enter>", on_hover)
    extract_unread_button.bind("<Leave>", on_leave)
    extract_unread_button.pack(pady=10)

    extract_recent_button = tk.Button(options_window, text="Extract From User", command=extract_recent_email, width=17)
    extract_recent_button.bind("<Enter>", on_hover)
    extract_recent_button.bind("<Leave>", on_leave)
    extract_recent_button.pack(pady=10)

    none_button = tk.Button(options_window, text="Data Present", command=get_sender_info_unread, width=17)
    none_button.bind("<Enter>", on_hover)
    none_button.bind("<Leave>", on_leave)
    none_button.pack(pady=10)

    options_window.mainloop()

def extract_unread_email():
    global options_window, sender, password,recent_email_window

    imap_url = 'imap.gmail.com'

    try:
        my_mail = imaplib.IMAP4_SSL(imap_url)
        my_mail.login(sender.strip(), password.strip())

        my_mail.select('inbox')

        status, email_ids = my_mail.search(None, "UNSEEN")

        email_addresses = []
        email_bodies = []

        mail_ids = email_ids[0].split()

        for mail_id in mail_ids:
            status, email_data = my_mail.fetch(mail_id, "(RFC822)")
            raw_email = email_data[0][1]
            email_message = email.message_from_bytes(raw_email)

            subject, encoding = decode_header(email_message["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8")
            sender_email = email.utils.parseaddr(email_message["From"])[1]
            email_addresses.append(sender_email)

            if email_message.is_multipart():
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode()
                        email_bodies.append(body)
            else:
                body = email_message.get_payload(decode=True).decode()
                email_bodies.append(body)

        data = {'EMAIL': email_addresses, 'CONTENT': email_bodies}
        df = pd.DataFrame(data)
        excel_file_path = 'email_data1.xlsx'
        df.to_excel(excel_file_path, index=False)
        messagebox.showinfo("Success", "Unread email data saved to Excel file!")
        get_sender_info_unread()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract unread email: {str(e)}")

    
def extract_recent_email():
    global email_entry,options_window,sender,password,recent_email_window
    options_window.destroy()
    recent_email_window = ttk.Window(themename='journal')
    recent_email_window.title("Extract Recent Email")
    recent_email_window.geometry("550x350")

    center(recent_email_window)

    font_size = 13

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(size=font_size)

    email_frame = tk.Frame(recent_email_window)
    email_frame.pack(pady=10)

    email_label = tk.Label(email_frame, text="Enter Email Address:")
    email_label.pack(side="left", padx=5)

    email_entry = tk.Entry(email_frame)
    email_entry.pack(side="left", padx=5)


    def extracting_recent():
        imap_url = 'imap.gmail.com'
        try:
            my_mail = imaplib.IMAP4_SSL(imap_url)
            my_mail.login(sender.strip(), password.strip())

            my_mail.select('inbox')
            email_addre=email_entry.get().strip()
            email_bodies = []
            key = 'FROM'

            _, data = my_mail.search(None, key, email_addre) 

            mail_id_list = data[0].split() 
            mail_id_list.reverse()

            msgs = [] 
            for num in mail_id_list:
                _, data = my_mail.fetch(num, '(RFC822)')
                msgs.append(data)
                break
                

            for msg in msgs[::-1]:
                for response_part in msg:
                    if type(response_part) is tuple:
                        my_msg=email.message_from_bytes((response_part[1]))
                        for part in my_msg.walk():  
                            if part.get_content_type() == 'text/plain':
                                email_addre
                                email_bodies.append(part.get_payload())
                            elif part.get_content_type() == 'text/html':
                                continue
                                
            data = {'EMAIL': email_addre, 'CONTENT': email_bodies}
            df = pd.DataFrame(data)

            excel_file_path = 'email_data1.xlsx'
            df.to_excel(excel_file_path, index=False)

            messagebox.showinfo("Success","Recent email extracted!")
            get_sender_info_recent()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract email: {str(e)}")

    
    def extracting_all():
        imap_url = 'imap.gmail.com'
        try:
            my_mail = imaplib.IMAP4_SSL(imap_url)
            my_mail.login(sender.strip(), password.strip())

            # Select the inbox
            my_mail.select('inbox')
            email_addre=email_entry.get().strip()
            email_bodies = []
            key = 'FROM'

            _, data = my_mail.search(None, key, email_addre) 

            mail_id_list = data[0].split() 

            msgs = [] 
            for num in mail_id_list:
                _, data = my_mail.fetch(num, '(RFC822)')
                msgs.append(data)
                
            for msg in msgs[::-1]:
                for response_part in msg:
                    if type(response_part) is tuple:
                        my_msg=email.message_from_bytes((response_part[1]))
                        for part in my_msg.walk():  
                            if part.get_content_type() == 'text/plain':
                                email_addre
                                email_bodies.append(part.get_payload())
                            elif part.get_content_type() == 'text/html':
                                continue
                                
            data = {'EMAIL': email_addre, 'CONTENT': email_bodies}
            df = pd.DataFrame(data)

            excel_file_path = 'email_data1.xlsx'
            df.to_excel(excel_file_path, index=False)

            messagebox.showinfo("Success","All emails extracted!")
            get_sender_info_recent()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract email: {str(e)}")

    extract_recent = tk.Button(recent_email_window, text="Extract Recent Email", command=extracting_recent)
    extract_recent.bind("<Enter>", on_hover)
    extract_recent.bind("<Leave>", on_leave)
    extract_recent.pack(pady=10)

    extract_all = tk.Button(recent_email_window, text="Extract All Email", command=extracting_all,width=17)
    extract_all.bind("<Enter>", on_hover)
    extract_all.bind("<Leave>", on_leave)
    extract_all.pack(pady=10)
    recent_email_window.mainloop()

open_login_window()