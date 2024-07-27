import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
from datetime import datetime
from tkfilebrowser import askopenfilenames
import csv
import docx
import PyPDF2
import schedule
import time
import threading
from functools import partial

# User Authentication
users = {'admin': 'password'}  # Simple user dictionary for authentication

def authenticate(username, password):
    return users.get(username) == password

def sign_up(username, password):
    if username in users:
        return False
    users[username] = password
    return True

# Database setup (for demonstration purposes, using SQLite)
def setup_database(db_name='example.db'):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS sales (
        date TEXT,
        product TEXT,
        quantity INTEGER,
        price REAL
    )
    ''')
    conn.commit()
    conn.close()

def insert_dummy_data(db_name='example.db'):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    cursor.executemany('''
    INSERT INTO sales (date, product, quantity, price) VALUES (?, ?, ?, ?)
    ''', [
        ('2023-07-01', 'Product A', 10, 9.99),
        ('2023-07-01', 'Product B', 5, 19.99),
        ('2023-07-02', 'Product A', 7, 9.99),
        ('2023-07-02', 'Product C', 3, 29.99),
        ('2023-07-03', 'Product B', 8, 19.99)
    ])
    conn.commit()
    conn.close()

# Insert data into the database
def insert_data(date, product, quantity, price, db_name='example.db'):
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    cursor.execute('''
    INSERT INTO sales (date, product, quantity, price) VALUES (?, ?, ?, ?)
    ''', (date, product, quantity, price))
    conn.commit()
    conn.close()
    messagebox.showinfo('Data Inserted', 'The data has been inserted successfully.')

# Fetch data from the database
def fetch_data(db_name='example.db'):
    conn = sqlite3.connect(db_name)
    df = pd.read_sql_query('SELECT * FROM sales', conn)
    conn.close()
    return df

# Export data to CSV
def export_data_to_csv(db_name='example.db'):
    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filename:
        df = fetch_data(db_name)
        df.to_csv(filename, index=False)
        messagebox.showinfo('Export Successful', f'Data exported to {filename}')

# Load data from CSV
def load_data_from_csv(db_name='example.db'):
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if filename:
        df = pd.read_csv(filename)
        conn = sqlite3.connect(db_name)
        df.to_sql('sales', conn, if_exists='append', index=False)
        conn.close()
        messagebox.showinfo('Import Successful', f'Data imported from {filename}')
        return filename
    return None

# Convert Word to CSV
def convert_word_to_csv():
    filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if filename:
        doc = docx.Document(filename)
        data = []
        for table in doc.tables:
            for row in table.rows:
                data.append([cell.text for cell in row.cells])
        csv_filename = filename.replace(".docx", ".csv")
        with open(csv_filename, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(data)
        messagebox.showinfo('Conversion Successful', f'Data converted and saved to {csv_filename}')
        return csv_filename
    return None

# Convert PDF to CSV
def convert_pdf_to_csv():
    filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if filename:
        reader = PyPDF2.PdfFileReader(filename)
        data = []
        for page in range(reader.getNumPages()):
            text = reader.getPage(page).extract_text()
            lines = text.split('\n')
            for line in lines:
                data.append(line.split(', '))
        csv_filename = filename.replace(".pdf", ".csv")
        with open(csv_filename, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(data)
        messagebox.showinfo('Conversion Successful', f'Data converted and saved to {csv_filename}')
        return csv_filename
    return None

# Advanced Analytics - Moving Average
def moving_average(df, window):
    df['moving_average'] = df['quantity'].rolling(window=window).mean()
    return df

# Advanced Analytics - Trend Analysis
def trend_analysis(df):
    df['date'] = pd.to_datetime(df['date'])
    df = df.set_index('date')
    trend = df['quantity'].resample('M').sum()
    return trend

# Generate report
def generate_report(data_source, chart_type, db_name='example.db'):
    if data_source == 'database':
        df = fetch_data(db_name)
    else:
        df = fetch_user_data()

    report_date = datetime.now().strftime('%Y-%m-%d')
    report_filename = f'report_{report_date}.png'

    # Group by product and calculate total quantity and revenue
    report = df.groupby('product').agg({'quantity': 'sum', 'price': 'sum'}).reset_index()
    report['revenue'] = report['quantity'] * report['price']

    # Plot the report
    fig, ax = plt.subplots(figsize=(10, 6))

    if chart_type == 'Bar':
        ax.bar(report['product'], report['quantity'], color='skyblue', label='Quantity')
        ax.set_ylabel('Quantity')
        ax2 = ax.twinx()
        ax2.bar(report['product'], report['revenue'], color='salmon', label='Revenue')
        ax2.set_ylabel('Revenue')
    elif chart_type == 'Line':
        ax.plot(report['product'], report['quantity'], marker='o', color='skyblue', label='Quantity')
        ax.set_ylabel('Quantity')
        ax2 = ax.twinx()
        ax2.plot(report['product'], report['revenue'], marker='o', color='salmon', label='Revenue')
        ax2.set_ylabel('Revenue')
    elif chart_type == 'Pie':
        ax.pie(report['quantity'], labels=report['product'], autopct='%1.1f%%', startangle=140)
        ax.set_title('Quantity Distribution')
        plt.figure()
        plt.pie(report['revenue'], labels=report['product'], autopct='%1.1%%', startangle=140)
        plt.title('Revenue Distribution')

    ax.set_title('Sales Report')
    plt.tight_layout()
    plt.savefig(report_filename)
    messagebox.showinfo('Report Generated', f'Report saved as {report_filename}')
    plt.close()

def generate_interactive_report(data_source, chart_type, db_name='example.db'):
    if data_source == 'database':
        df = fetch_data(db_name)
    else:
        df = fetch_user_data()

    report_date = datetime.now().strftime('%Y-%m-%d')
    report_filename = f'interactive_report_{report_date}.html'

    # Group by product and calculate total quantity and revenue
    report = df.groupby('product').agg({'quantity': 'sum', 'price': 'sum'}).reset_index()
    report['revenue'] = report['quantity'] * report['price']

    # Plot the report
    if chart_type == 'Bar':
        fig = go.Figure()
        fig.add_trace(go.Bar(x=report['product'], y=report['quantity'], name='Quantity', marker_color='skyblue'))
        fig.add_trace(go.Bar(x=report['product'], y=report['revenue'], name='Revenue', marker_color='salmon'))
        fig.update_layout(barmode='group', title='Sales Report')
    elif chart_type == 'Line':
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=report['product'], y=report['quantity'], mode='lines+markers', name='Quantity'))
        fig.add_trace(go.Scatter(x=report['product'], y=report['revenue'], mode='lines+markers', name='Revenue'))
        fig.update_layout(title='Sales Report')
    elif chart_type == 'Pie':
        fig = make_subplots(rows=1, cols=2, specs=[[{'type': 'domain'}, {'type': 'domain'}]])
        fig.add_trace(go.Pie(labels=report['product'], values=report['quantity'], name='Quantity'), 1, 1)
        fig.add_trace(go.Pie(labels=report['product'], values=report['revenue'], name='Revenue'), 1, 2)
        fig.update_layout(title_text='Sales Report')

    fig.write_html(report_filename)
    messagebox.showinfo('Interactive Report Generated', f'Interactive report saved as {report_filename}')

# Function to add data from GUI input
def add_data(db_name='example.db'):
    date = date_entry.get()
    product = product_entry.get()
    quantity = quantity_entry.get()
    price = price_entry.get()

    if not date or not product or not quantity or not price:
        messagebox.showerror('Input Error', 'All fields are required.')
        return

    try:
        quantity = int(quantity)
        price = float(price)
    except ValueError:
        messagebox.showerror('Input Error', 'Quantity must be an integer and price must be a float.')
        return

    user_data.append((date, product, quantity, price))
    messagebox.showinfo('Data Added', 'The data has been added to the user input.')

def fetch_user_data():
    columns = ['date', 'product', 'quantity', 'price']
    return pd.DataFrame(user_data, columns=columns)

def choose_data_source(data_source, db_name='example.db'):
    if data_source == 'database':
        choose_chart_type('database', db_name)
    elif data_source == 'user':
        choose_chart_type('user', db_name)

def choose_chart_type(data_source, db_name='example.db'):
    chart_type = chart_type_var.get()
    if not chart_type:
        messagebox.showerror('Selection Error', 'Please select a chart type.')
    else:
        generate_report(data_source, chart_type, db_name)

def choose_interactive_chart_type(data_source, db_name='example.db'):
    chart_type = chart_type_var.get()
    if not chart_type:
        messagebox.showerror('Selection Error', 'Please select a chart type.')
    else:
        generate_interactive_report(data_source, chart_type, db_name)

def schedule_report(interval, data_source, chart_type, db_name='example.db'):
    if data_source == 'database':
        job = partial(generate_report, 'database', chart_type, db_name)
    else:
        job = partial(generate_report, 'user', chart_type, db_name)
    
    if interval == 'Daily':
        schedule.every().day.at("10:00").do(job)
    elif interval == 'Weekly':
        schedule.every().monday.at("10:00").do(job)
    elif interval == 'Monthly':
        schedule.every().month.at("10:00").do(job)
    
    threading.Thread(target=schedule.run_pending).start()
    messagebox.showinfo('Report Scheduled', f'Report scheduled {interval.lower()} at 10:00.')

# GUI setup
def setup_gui():
    global user_data, date_entry, product_entry, quantity_entry, price_entry, chart_type_var

    user_data = []

    def on_start():
        user_choice = messagebox.askquestion('Data Source', 'Would you like to use random data from the database?')
        if user_choice == 'yes':
            choose_data_source('database')
        else:
            input_window()

    def input_window():
        input_root = tk.Toplevel(root)
        input_root.title("Input Data")
        input_root.geometry("400x300")

        frame = ttk.Frame(input_root, padding=10)
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        global date_entry, product_entry, quantity_entry, price_entry

        ttk.Label(frame, text="Date (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        date_entry = ttk.Entry(frame)
        date_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Product:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        product_entry = ttk.Entry(frame)
        product_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Quantity:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        quantity_entry = ttk.Entry(frame)
        quantity_entry.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Price:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        price_entry = ttk.Entry(frame)
        price_entry.grid(row=3, column=1, padx=5, pady=5)

        add_button = ttk.Button(frame, text="Add Data", command=add_data)
        add_button.grid(row=4, column=0, padx=20, pady=20)

        generate_button = ttk.Button(frame, text="Generate Report", command=lambda: choose_data_source('user'))
        generate_button.grid(row=4, column=1, padx=20, pady=20)

    def login_window():
        def login():
            username = username_entry.get()
            password = password_entry.get()
            if authenticate(username, password):
                login_root.destroy()
                main_window()
            else:
                messagebox.showerror('Login Failed', 'Invalid username or password.')

        def sign_up_window():
            def sign_up_user():
                new_username = new_username_entry.get()
                new_password = new_password_entry.get()
                if sign_up(new_username, new_password):
                    messagebox.showinfo('Sign Up Successful', 'Account created successfully.')
                    sign_up_root.destroy()
                else:
                    messagebox.showerror('Sign Up Failed', 'Username already exists.')

            sign_up_root = tk.Toplevel(login_root)
            sign_up_root.title("Sign Up")
            sign_up_root.geometry("300x200")

            frame = ttk.Frame(sign_up_root, padding=10)
            frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

            ttk.Label(frame, text="New Username:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
            new_username_entry = ttk.Entry(frame)
            new_username_entry.grid(row=0, column=1, padx=5, pady=5)

            ttk.Label(frame, text="New Password:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
            new_password_entry = ttk.Entry(frame, show='*')
            new_password_entry.grid(row=1, column=1, padx=5, pady=5)

            sign_up_button = ttk.Button(frame, text="Sign Up", command=sign_up_user)
            sign_up_button.grid(row=2, column=0, columnspan=2, pady=20)

        login_root = tk.Toplevel(root)
        login_root.title("Login")
        login_root.geometry("300x200")

        frame = ttk.Frame(login_root, padding=10)
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="Username:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        username_entry = ttk.Entry(frame)
        username_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Password:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        password_entry = ttk.Entry(frame, show='*')
        password_entry.grid(row=1, column=1, padx=5, pady=5)

        login_button = ttk.Button(frame, text="Login", command=login)
        login_button.grid(row=2, column=0, columnspan=2, pady=20)

        sign_up_button = ttk.Button(frame, text="Sign Up", command=sign_up_window)
        sign_up_button.grid(row=3, column=0, columnspan=2, pady=20)

    def main_window():
        main_root = tk.Toplevel(root)
        main_root.title("Automated Reporting Tool")
        main_root.geometry("500x600")

        frame = ttk.Frame(main_root, padding=10)
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        start_button = ttk.Button(frame, text="Start", command=on_start)
        start_button.grid(row=0, column=0, padx=20, pady=20)

        chart_type_var = tk.StringVar()

        ttk.Label(frame, text="Select Chart Type:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(frame, text="Bar Chart", variable=chart_type_var, value="Bar").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(frame, text="Line Chart", variable=chart_type_var, value="Line").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(frame, text="Pie Chart", variable=chart_type_var, value="Pie").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)

        export_button = ttk.Button(frame, text="Export Data to CSV", command=export_data_to_csv)
        export_button.grid(row=5, column=0, padx=20, pady=20)

        load_button = ttk.Button(frame, text="Load Data from CSV", command=load_data_from_csv)
        load_button.grid(row=6, column=0, padx=20, pady=20)

        word_to_csv_button = ttk.Button(frame, text="Convert Word to CSV", command=convert_word_to_csv)
        word_to_csv_button.grid(row=7, column=0, padx=20, pady=20)

        pdf_to_csv_button = ttk.Button(frame, text="Convert PDF to CSV", command=convert_pdf_to_csv)
        pdf_to_csv_button.grid(row=8, column=0, padx=20, pady=20)

        ttk.Label(frame, text="Schedule Report:").grid(row=9, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(frame, text="Daily", variable=chart_type_var, value="Daily").grid(row=10, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(frame, text="Weekly", variable=chart_type_var, value="Weekly").grid(row=11, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(frame, text="Monthly", variable=chart_type_var, value="Monthly").grid(row=12, column=0, padx=5, pady=5, sticky=tk.W)

        schedule_button = ttk.Button(frame, text="Schedule Report", command=lambda: schedule_report(chart_type_var.get(), 'database', chart_type_var.get()))
        schedule_button.grid(row=13, column=0, padx=20, pady=20)

        interactive_button = ttk.Button(frame, text="Generate Interactive Report", command=lambda: choose_interactive_chart_type('database'))
        interactive_button.grid(row=14, column=0, padx=20, pady=20)

    root = tk.Tk()
    root.title("User Authentication")
    root.geometry("400x200")

    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    login_button = ttk.Button(frame, text="Login", command=login_window)
    login_button.grid(row=0, column=0, padx=20, pady=20)

    root.mainloop()

if __name__ == "__main__":
    setup_database()
    insert_dummy_data()
    setup_gui()
