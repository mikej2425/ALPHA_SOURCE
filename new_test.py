import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont
import pandas as pd
import tkcalendar 
from pandas.tseries.offsets import BDay
import datetime
from collections import defaultdict

# InsertTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/InsertTable.xlsx"
# ResourceTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/ResourcesTable.xlsx"
# MsdtDivisionTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/MsdtDivisionTable.xlsx"
# CyberSecurityDivisionTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/CyberSecurityDivisionTable.xlsx"
# CloudDivisionTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/CloudDivisionTable.xlsx"
# ProjectTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/ProjectTable.xlsx"

InsertTable = "./Alpha Data Project/InsertTable.xlsx"
ResourceTable = "./Alpha Data Project/ResourcesTable.xlsx"
MsdtDivisionTable = "./Alpha Data Project/MsdtDivisionTable.xlsx"
CyberSecurityDivisionTable = "./Alpha Data Project/CyberSecurityDivisionTable.xlsx"
CloudDivisionTable = "./Alpha Data Project/CloudDivisionTable.xlsx"
ProjectTable = "./Alpha Data Project/ProjectTable.xlsx"

df_insert_table = pd.read_excel(InsertTable)
df_resource_table = pd.read_excel(ResourceTable)
df_msdtdivision_table = pd.read_excel(MsdtDivisionTable)
df_cybersecuritydivision_table = pd.read_excel(CyberSecurityDivisionTable)
df_clouddivision_table = pd.read_excel(CloudDivisionTable)
df_project_table = pd.read_excel(ProjectTable)

import os

# Check if running on a headless environment (like Render)
if 'DISPLAY' not in os.environ:
    # If not, set up a virtual display using xvfb
    os.system('Xvfb :0 -screen 0 1024x768x24 &')
    os.environ['DISPLAY'] = ':0'

class App:


    def __init__(self, root):
        self.root = root
        root.title("Forms")
        width = 600
        height = 500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
        
        ft = tkFont.Font(family='Times', size=10)
        self.add_button = tk.Button(root)
        self.add_button["bg"] = "#f0f0f0"
        self.add_button["font"] = ft
        self.add_button["fg"] = "#000000"
        self.add_button["justify"] = "center"
        self.add_button["text"] = "Add"
        self.add_button.place(x=240, y=190, width=148, height=42)
        self.add_button["command"] = self.add_job
        
        self.view_button = tk.Button(root)
        self.view_button["bg"] = "#f0f0f0"
        self.view_button["font"] = ft
        self.view_button["fg"] = "#000000"
        self.view_button["justify"] = "center"
        self.view_button["text"] = "View"
        self.view_button.place(x=240, y=240, width=148, height=42)
        self.view_button["command"] = self.view_jobs



    def add_job(self):
        self.root.withdraw()
        self.add_window = tk.Toplevel(root)
        self.add_window.title("Add Job")
        width = 600
        height = 500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.add_window.geometry(alignstr)
        self.add_window.resizable(width=False, height=False)
        
        tk.Label(self.add_window, text="Name:").place(x=130,y=130,width=148,height=42)
        self.name_entry = tk.Entry(self.add_window)
        self.name_entry.place(x=300,y=130,width=148,height=42)

        tk.Label(self.add_window, text="Date").place(x=130,y=190,width=148,height=42)
        self.cal = tkcalendar.DateEntry(self.add_window, width=12, background='darkblue',
                            foreground='white', borderwidth=2)
        self.cal.place(x=300,y=190,width=148,height=42)

        tk.Label(self.add_window, text="Project Name:").place(x=130,y=250,width=148,height=42)
        self.project_name_combobox = ttk.Combobox(self.add_window, values=df_project_table['Project Name'].values.tolist())
        self.project_name_combobox.place(x=300,y=250,width=148,height=42)

        tk.Label(self.add_window, text="Man-Hour:").place(x=130,y=310,width=148,height=42)
        self.manhour_entry = tk.Entry(self.add_window)
        self.manhour_entry.place(x=300,y=310,width=148,height=42)

        save_button = tk.Button(self.add_window, text="Add", command=self.save_to_excel)
        save_button.place(x=210,y=430,width=148,height=42)

        back_button = tk.Button(self.add_window, text="Back", command=self.back_to_main_form)
        back_button.place(x=20,y=20,width=70,height=25)
        
        self.add_window.mainloop()

    def save_to_excel(self):
        name = self.name_entry.get()
        project_name = self.project_name_combobox.get()
        calendar = self.cal.get_date().strftime('%m-%d-%Y')
        man_hours = self.manhour_entry.get()

        if man_hours == "":
            messagebox.showerror("Error", "Please enter Man-hours.")
            return False

        try:
            man_hours = int(man_hours)
        except ValueError:
            messagebox.showerror("Error", "Man-hours should be an integer.")
            return False

        if not (1 <= man_hours <= 24):
            messagebox.showerror("Error", "Man-hours should be between 1 and 24.")
            return False

        if project_name == "":
            messagebox.showerror("Error", "Select Project Name")
            return False
        
        # Check if the total man-hours for the employee on the given date exceed 24
        existing_insert_df = pd.read_excel(InsertTable)
        existing_insert_df['From Date'] = pd.to_datetime(existing_insert_df['From Date'])
        existing_insert_df = existing_insert_df[existing_insert_df['Name'] == name]
        existing_insert_df = existing_insert_df[existing_insert_df['From Date'] == pd.to_datetime(calendar)]
        total_man_hours_for_date = existing_insert_df['Man-hours'].sum()
        if total_man_hours_for_date + man_hours > 24:
            messagebox.showerror("Error", f"{name} cannot work more than 24 hours in a day.")
            return False

        try:
            # Read the existing Insert table
            existing_insert_df = pd.read_excel(InsertTable)

            # Check if the name already exists in the Insert table
            existing_entry = existing_insert_df[existing_insert_df['Name'] == name]

            # Create a new row
            new_row = {'Name': name, 'Project Name': project_name, 'From Date': calendar, 'Man-hours': man_hours}

            # Append the new row to the Insert table
            existing_insert_df = pd.concat([existing_insert_df, pd.DataFrame([new_row])], ignore_index=True)

            # Calculate cost for new row
            if not existing_entry.empty:
                existing_insert_df['From Date'] = pd.to_datetime(existing_insert_df['From Date'])
                existing_insert_df['Month'] = existing_insert_df['From Date'].dt.month
                existing_insert_df['Year'] = existing_insert_df['From Date'].dt.year

                total_man_hours = existing_insert_df.groupby(['Name', 'Month', 'Year'])['Man-hours'].transform('sum')
                total_cost_by_month = df_resource_table[df_resource_table['Employee Name'] == name]['Total Cost by Month'].values[0]
                existing_insert_df['Cost'] = (total_cost_by_month / total_man_hours) * existing_insert_df['Man-hours']
            else:
                total_cost_by_month = df_resource_table[df_resource_table['Employee Name'] == name]['Total Cost by Month'].values[0]
                existing_insert_df['Cost'] = total_cost_by_month

            existing_insert_df.to_excel(InsertTable, index=False)

            # Update the division tables
            employee_division = df_resource_table.loc[df_resource_table['Employee Name'] == name, 'Division'].iloc[0]
            if employee_division == 'MSDT/AIOT':
                division_file = MsdtDivisionTable
            elif employee_division == 'MSDT/CLOUD':
                division_file = CloudDivisionTable
            elif employee_division == 'MSDT/CYBERSEC':
                division_file = CyberSecurityDivisionTable
            else:
                messagebox.showerror("Error", f"Unknown division for employee: {name}")
                return False

            existing_division_df = pd.read_excel(division_file)
            updated_df = pd.concat([existing_division_df, pd.DataFrame([new_row])], ignore_index=True)

            # Calculate total man-hours and adjust cost for the division table
            updated_df['From Date'] = pd.to_datetime(updated_df['From Date'])
            updated_df['Month'] = updated_df['From Date'].dt.month
            updated_df['Year'] = updated_df['From Date'].dt.year
            total_man_hours_division = updated_df.groupby(['Name', 'Month', 'Year'])['Man-hours'].transform('sum')
            updated_df['Cost'] = (total_cost_by_month / total_man_hours_division) * updated_df['Man-hours']

            updated_df.to_excel(division_file, index=False)

            # Update project table with monthly cost
            updated_insert_df = pd.read_excel(InsertTable)
            project_cost_series = updated_insert_df.groupby('Project Name')['Cost'].sum()
            for project_name, project_cost in project_cost_series.items():
                df_project_table.loc[df_project_table['Project Name'] == project_name, 'Project Cost'] = project_cost
            df_project_table.to_excel(ProjectTable, index=False)

            messagebox.showinfo("Success", f"Data inserted into {employee_division} and Insert table successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

        
        

        
    def view_jobs(self):
        self.root.withdraw()
        self.view_job_window = tk.Toplevel(root)
        self.view_job_window.title("View Job")
        width = 600
        height = 500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.view_job_window.geometry(alignstr)
        self.view_job_window.resizable(width=False, height=False)
        
        tk.Label(self.view_job_window, text="Select").place(x=230,y=60,width=148,height=42)

        employee_button = tk.Button(self.view_job_window, text="Employee", command=self.employee_view_jobs)
        employee_button.place(x=230,y=130,width=148,height=42)

        division_button = tk.Button(self.view_job_window, text="Division", command=self.division_view_jobs)
        division_button.place(x=230,y=190,width=148,height=42)

        project_button = tk.Button(self.view_job_window, text = "Project", command=self.project_view_jobs)
        project_button.place(x=230,y=250,width=148,height=42)

        back_button = tk.Button(self.view_job_window, text="Back", command=self.back_to_main_form1)
        back_button.place(x=20,y=20,width=70,height=25)

        self.view_job_window.mainloop()


        
    def division_view_jobs(self):
        self.view_job_window.withdraw()
        self.division_view_jobs_window = tk.Toplevel(root)
        self.division_view_jobs_window.title("View Division")
        width = 600
        height = 500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.division_view_jobs_window.geometry(alignstr)
        self.division_view_jobs_window.resizable(width=False, height=False)
        
        select_combobox = ttk.Combobox(self.division_view_jobs_window, values=["MsdtDivisionTable", "CyberSecurityDivisionTable", "CloudDivisionTable"])
        select_combobox.set("Select Division")
        select_combobox.place(x=10,y=70,width=148,height=42)

        select_treeview = ttk.Treeview(self.division_view_jobs_window, show="headings")
        select_treeview.place(x=180,y=10,width=400,height=472)
        
        scrollbar = tk.Scrollbar(self.division_view_jobs_window, orient="vertical", command=select_treeview.yview)
        scrollbar.place(x=580, y=10, height=472)
        select_treeview.configure(yscrollcommand=scrollbar.set)

        def load_excel():
            selected_excel = select_combobox.get()
            if selected_excel == "MsdtDivisionTable":
                df = df_msdtdivision_table
            elif selected_excel == "CyberSecurityDivisionTable":
                df = df_cybersecuritydivision_table
            elif selected_excel == "CloudDivisionTable":
                df = df_clouddivision_table
            else:
                messagebox.showerror("Error", "Please select a valid Excel sheet.")
                return

            for col in select_treeview.get_children():
                select_treeview.delete(col)

            select_treeview["columns"] = df.columns.tolist()
            for col in df.columns.tolist():
                select_treeview.heading(col, text=col)
                select_treeview.column(col, width=100, anchor="center")

            for index, row in df.iterrows():
                select_treeview.insert("", "end", values=tuple(row))
            
        load_button = tk.Button(self.division_view_jobs_window, text="Load", command=load_excel)
        load_button.place(x=20,y=140,width=70,height=25)
        
        back_button = tk.Button(self.division_view_jobs_window, text="Back", command=self.back_to_view_job)
        back_button.place(x=20,y=20,width=70,height=25)

        self.division_view_jobs_window.mainloop()
        


    def employee_view_jobs(self):
        
        def search_employee():
            search_query = self.search_var.get().strip().lower()  
            project_name_query = self.search_project_name_combobox.get().strip().lower()

            filtered_df = df_insert_table

            if search_query:
                filtered_df = filtered_df[filtered_df['Name'].str.lower().str.contains(search_query)]
            
            if project_name_query:
                filtered_df = filtered_df[filtered_df['Project Name'].str.lower().str.contains(project_name_query)]
            
            for item in tree.get_children():
                tree.delete(item)
            
            for index, row in filtered_df.iterrows():
                tree.insert("", "end", values=list(row))

        self.view_job_window.withdraw()
        self.employee_view_jobs_window = tk.Toplevel(root)
        self.employee_view_jobs_window.title("View Employee")
        width = 600
        height = 500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.employee_view_jobs_window.geometry(alignstr)
        self.employee_view_jobs_window.resizable(width=False, height=False)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self.employee_view_jobs_window, textvariable=self.search_var)
        self.search_entry.place(x=10, y=70, width=148, height=42)

        self.search_project_name_combobox = ttk.Combobox(self.employee_view_jobs_window, values=df_project_table['Project Name'].values.tolist())
        self.search_project_name_combobox.place(x=10,y=130,width=148,height=42)
        
        self.search_button = tk.Button(self.employee_view_jobs_window, text="Search", command=search_employee)
        self.search_button.place(x=20, y=190, width=70, height=25)

        tree = ttk.Treeview(self.employee_view_jobs_window, show="headings")
        tree["columns"] = list(df_insert_table.columns)
        tree.place(x=180, y=10, width=400, height=472)

        scrollbar = tk.Scrollbar(self.employee_view_jobs_window, orient="vertical", command=tree.yview)
        scrollbar.place(x=580, y=10, height=472)
        tree.configure(yscrollcommand=scrollbar.set)

        for col in df_insert_table.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")

        for index, row in df_insert_table.iterrows():
            tree.insert("", "end", values=tuple(row))

        back_button = tk.Button(self.employee_view_jobs_window, text="Back", command=self.back_to_view_job1)
        back_button.place(x=20,y=20,width=70,height=25)            

        self.employee_view_jobs_window.mainloop()



    def project_view_jobs(self):
        self.view_job_window.withdraw()
        self.project_view_jobs_window = tk.Toplevel(root)
        self.project_view_jobs_window.title("View Project")
        width = 600
        height = 500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.project_view_jobs_window.geometry(alignstr)
        self.project_view_jobs_window.resizable(width=False, height=False)

        tree = ttk.Treeview(self.project_view_jobs_window, show="headings")
        tree["columns"] = list(df_project_table.columns)
        tree.place(x=180, y=10, width=400, height=472)

        scrollbar = tk.Scrollbar(self.project_view_jobs_window, orient="vertical", command=tree.yview)
        scrollbar.place(x=580, y=10, height=472)
        tree.configure(yscrollcommand=scrollbar.set)

        for col in df_project_table.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")

        for index, row in df_project_table.iterrows():
            tree.insert("", "end", values=tuple(row))

        back_button = tk.Button(self.project_view_jobs_window, text="Back", command=self.back_to_view_job2)
        back_button.place(x=20,y=20,width=70,height=25) 


    def back_to_main_form(self):
        self.root.deiconify()
        self.add_window.destroy()

    def back_to_main_form1(self):
        self.root.deiconify()
        self.view_job_window.destroy()

    def back_to_view_job(self):
        self.view_job_window.deiconify()
        self.division_view_jobs_window.destroy()

    def back_to_view_job1(self):
        self.view_job_window.deiconify()
        self.employee_view_jobs_window.destroy()

    def back_to_view_job2(self):
        self.view_job_window.deiconify()
        self.project_view_jobs_window.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
