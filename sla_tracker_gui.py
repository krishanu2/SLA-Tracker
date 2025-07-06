import os
import pandas as pd
from datetime import datetime
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox
import smtplib
from email.message import EmailMessage
from fpdf import FPDF
import webbrowser

# === Constants ===
DATA_FILE = "sla_data.xlsx"
if not os.path.exists(DATA_FILE):
    df_init = pd.DataFrame(columns=["Task ID", "Title", "Owner", "Priority", "Start Date", "End Date", "Email",
                                    "Status", "Completion Date", "SLA Breached", "Rating", "Review"])
    df_init.to_excel(DATA_FILE, index=False)

df = pd.read_excel(DATA_FILE)

# === Email Function ===
def send_email(to_email, subject, body, attachment_path=None):
    EMAIL_SENDER = "krishanumahapatra7777@gmail.com"
    APP_PASSWORD = "yijhoqwxpdhmxlgu"

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_SENDER
    msg['To'] = to_email
    msg.set_content(body)

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)
            msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_SENDER, APP_PASSWORD)
            smtp.send_message(msg)
        print("Email sent to:", to_email)
    except Exception as e:
        print("Email error:", e)

# === PDF Generation ===
def generate_pdf_report(task_id, name, rating, review):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=12)

    pdf.cell(0, 10, "Task Completion Report", ln=True, align='C')
    pdf.ln(10)
    pdf.cell(0, 10, f"Task ID: {task_id}", ln=True)
    pdf.cell(0, 10, f"Employee Name: {name}", ln=True)

    if rating:
        pdf.cell(0, 10, f"Performance Rating: {rating}/5", ln=True)
    if review:
        pdf.multi_cell(0, 10, f"Review: {review}")

    file_path = f"{task_id}_Report.pdf"
    pdf.output(file_path)

    try:
        webbrowser.open_new(file_path)
    except Exception as e:
        print("Couldn't open PDF:", e)

    return file_path

# === Task Management Functions ===
def generate_task_id():
    if df.empty:
        return "TSK001"
    last_id = df["Task ID"].dropna().iloc[-1]
    number = int(last_id.replace("TSK", "")) + 1
    return f"TSK{number:03d}"

def save_task():
    title = title_entry.get()
    owner = owner_entry.get()
    priority = priority_combo.get()
    start = start_entry.get()
    end = end_entry.get()
    email = email_entry.get()

    if not all([title, owner, priority, start, end, email]):
        messagebox.showerror("Input Error", "All fields are required.")
        return

    try:
        datetime.strptime(start, "%d-%m-%Y")
        datetime.strptime(end, "%d-%m-%Y")
    except:
        messagebox.showerror("Date Format", "Dates must be in dd-mm-yyyy format")
        return

    task_id = generate_task_id()
    new_row = {
        "Task ID": task_id,
        "Title": title,
        "Owner": owner,
        "Priority": priority,
        "Start Date": start,
        "End Date": end,
        "Email": email,
        "Status": "Pending",
        "Completion Date": "",
        "SLA Breached": "",
        "Rating": "",
        "Review": ""
    }
    global df
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(DATA_FILE, index=False)
    update_table()

    send_email(email, f"New Task Assigned: {task_id}",
               f"Hi {owner},\n\nYou have been assigned a new task:\n\nTask: {title}\nPriority: {priority}\nStart: {start}\nEnd: {end}\n\nSLA Tracker")

    for e in [title_entry, owner_entry, start_entry, end_entry, email_entry]:
        e.delete(0, tk.END)

def update_table():
    for row in table.get_children():
        table.delete(row)
    for _, row in df.iterrows():
        table.insert('', 'end', values=list(row))

def on_row_select(event):
    selected = table.focus()
    if not selected:
        return
    values = table.item(selected, 'values')
    selected_task_id.set(values[0])
    status_combo.set(values[7])
    completion_entry.delete(0, tk.END)
    completion_entry.insert(0, values[8])
    rating_entry.delete(0, tk.END)
    rating_entry.insert(0, values[10])
    review_entry.delete(0, tk.END)
    review_entry.insert(0, values[11])

def update_status():
    tid = selected_task_id.get()
    status = status_combo.get()
    completion = completion_entry.get()
    rating = rating_entry.get()
    review = review_entry.get()

    if not tid or not status:
        messagebox.showerror("Error", "Select a task and status")
        return

    idx = df.index[df["Task ID"] == tid].tolist()
    if not idx:
        messagebox.showerror("Error", "Task not found")
        return
    idx = idx[0]
    df.at[idx, "Status"] = status
    df.at[idx, "Completion Date"] = completion
    df.at[idx, "Rating"] = rating
    df.at[idx, "Review"] = review

    try:
        due = datetime.strptime(df.at[idx, "End Date"], "%d-%m-%Y")
        comp = datetime.strptime(completion, "%d-%m-%Y")
        df.at[idx, "SLA Breached"] = "Yes" if comp > due else "No"
    except:
        df.at[idx, "SLA Breached"] = "Unknown"

    df.to_excel(DATA_FILE, index=False)
    update_table()
    messagebox.showinfo("Updated", f"Task {tid} updated successfully.")

    if status == "Completed" and review:
        pdf_path = generate_pdf_report(tid, df.at[idx, "Owner"], rating, review)
        send_email(
            df.at[idx, "Email"],
            f"Performance Report for Task {tid}",
            f"Hi {df.at[idx, 'Owner']},\n\nYour task is completed. See attached performance report.",
            attachment_path=pdf_path
        )

# === GUI ===
app = tb.Window(themename="darkly")
app.title("SLA Tracker")
app.geometry("1100x650")

notebook = tb.Notebook(app)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# === Add Task Tab ===
add_tab = tb.Frame(notebook)
notebook.add(add_tab, text="➕ Add Task")

labels = ["Title", "Owner", "Priority", "Start Date (dd-mm-yyyy)", "End Date (dd-mm-yyyy)", "Email"]
entries = []
for i, label in enumerate(labels):
    tb.Label(add_tab, text=label).grid(row=i, column=0, sticky="w", pady=5)
    entry = tb.Entry(add_tab)
    entry.grid(row=i, column=1, sticky="ew", pady=5)
    entries.append(entry)

[title_entry, owner_entry, priority_combo, start_entry, end_entry, email_entry] = entries
priority_combo = tb.Combobox(add_tab, values=["Low", "Medium", "High"])
priority_combo.grid(row=2, column=1, sticky="ew", pady=5)

add_tab.columnconfigure(1, weight=1)
tb.Button(add_tab, text="Add Task", bootstyle=SUCCESS, command=save_task).grid(row=6, columnspan=2, pady=20)

# === Dashboard ===
dash_tab = tb.Frame(notebook)
notebook.add(dash_tab, text="📊 Dashboard")

table = tb.Treeview(dash_tab, columns=list(df.columns), show="headings", height=15)
for col in df.columns:
    table.heading(col, text=col)
    table.column(col, width=120)
table.pack(fill="both", expand=True)
table.bind("<<TreeviewSelect>>", on_row_select)

update_frame = tb.Frame(dash_tab)
update_frame.pack(fill="x", pady=10)

selected_task_id = tk.StringVar()
tb.Label(update_frame, text="Selected Task ID:").grid(row=0, column=0, padx=5)
tb.Label(update_frame, textvariable=selected_task_id).grid(row=0, column=1, padx=5)

status_combo = tb.Combobox(update_frame, values=["Pending", "In Progress", "Completed"], width=15)
status_combo.grid(row=0, column=2, padx=5)

completion_entry = tb.Entry(update_frame)
completion_entry.insert(0, "Completion Date")
completion_entry.grid(row=0, column=3, padx=5)

rating_entry = tb.Entry(update_frame, width=5)
rating_entry.insert(0, "Rating")
rating_entry.grid(row=0, column=4, padx=5)

review_entry = tb.Entry(update_frame, width=30)
review_entry.insert(0, "Review")
review_entry.grid(row=0, column=5, padx=5)

tb.Button(update_frame, text="Update Task", bootstyle=PRIMARY, command=update_status).grid(row=0, column=6, padx=5)

update_table()
app.mainloop()
