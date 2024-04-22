import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont
import pandas as pd
import tkcalendar
from pandas.tseries.offsets import BDay
import datetime
from collections import defaultdict

InsertTable = "C:/Users/Jubin/Desktop/Desktop/Alpha Data Project/InsertTable.xlsx"
ResourceTable = "C:/Users/Jubin/Desktop/Desktop/Alpha Data Project/ResourcesTable.xlsx"
MsdtDivisionTable = "C:/Users/Jubin/Desktop/Desktop/Alpha Data Project/MsdtDivisionTable.xlsx"
CyberSecurityDivisionTable = "C:/Users/Jubin/Desktop/Desktop/Alpha Data Project/CyberSecurityDivisionTable.xlsx"
CloudDivisionTable = "C:/Users/Jubin/Desktop/Desktop/Alpha Data Project/CloudDivisionTable.xlsx"
ProjectTable = "C:/Users/Jubin/Desktop/Desktop/Alpha Data Project/ProjectTable.xlsx"

df_insert_table = pd.read_excel(InsertTable)
df_resource_table = pd.read_excel(ResourceTable)
df_msdtdivision_table = pd.read_excel(MsdtDivisionTable)
df_cybersecuritydivision_table = pd.read_excel(CyberSecurityDivisionTable)
df_clouddivision_table = pd.read_excel(CloudDivisionTable)
df_project_table = pd.read_excel(ProjectTable)

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
        name_entry = tk.Entry(self.add_window)
        name_entry.place(x=300,y=130,width=148,height=42)

        tk.Label(self.add_window, text="Date").place(x=130,y=190,width=148,height=42)
        cal = tkcalendar.DateEntry(self.add_window, width=12, background='darkblue',
                            foreground='white', borderwidth=2)
        cal.place(x=300,y=190,width=148,height=42)

        tk.Label(self.add_window, text="Project Name:").place(x=130,y=250,width=148,height=42)
        project_name_combobox = ttk.Combobox(self.add_window, values=df_project_table['Project Name'].values.tolist())
        project_name_combobox.place(x=300,y=250,width=148,height=42)

        tk.Label(self.add_window, text="Manday:").place(x=130,y=310,width=148,height=42)
        manday_entry = tk.Entry(self.add_window)
        manday_entry.place(x=300,y=310,width=148,height=42)


        def count_business_workdays(year, month):
            start_date = datetime.date(year, month, 1)
            end_date = start_date.replace(day=28) + datetime.timedelta(days=4)  
            end_date = end_date - datetime.timedelta(days=end_date.day)  
            business_days = 0
            current_date = start_date
            while current_date <= end_date:
                if current_date.weekday() < 5:  
                    business_days += 1
                current_date += datetime.timedelta(days=1)
            return business_days
        

        def save_to_excel():
            name = name_entry.get()
            project_name = project_name_combobox.get()
            manday = manday_entry.get()
            calender = cal.get_date().strftime('%m-%d-%Y')

            if not manday.isdigit():
                messagebox.showerror("Error", "Please insert a valid integer for Manday")
                return False
            else:
                manday = int(manday)
                if manday < 1:
                    messagebox.showerror("Error", "Manday should be more then 0")
                    return False

            if project_name == "":
                messagebox.showerror("Error", "Select Project Name")
                return False

            manday = int(manday)
            start_date = pd.to_datetime(calender, format='%m-%d-%Y')
            end_date = start_date + pd.offsets.BDay(manday - 1)

            if end_date.month != start_date.month or end_date.year != start_date.year:
                messagebox.showerror("Error", f"Mandays exceeds {start_date.strftime('%B %Y')}")
                return False

            new_row = {'Name': name, 'Project Name': project_name, 'Manday': manday, 'From Date': calender}

            try:
                if name not in df_resource_table['Employee Name'].values:
                    messagebox.showerror("Error", "Not an employee!")
                    return False

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
                        
                if project_name not in df_project_table["Project Name"].values:
                    messagebox.showerror("Error", "Invalid Project Name!")
                    return False
                
                existing_insert_df = pd.read_excel(InsertTable)
                    
                if not existing_insert_df.empty:
                    existing_entry = existing_insert_df[existing_insert_df['Name'] == name]

                    for index, row in existing_entry.iterrows():
                        existing_start_date = pd.to_datetime(row['From Date'])
                        existing_end_date = pd.to_datetime(row['Till Date'])
                        if (start_date >= existing_start_date and start_date <= existing_end_date) or \
                        (end_date >= existing_start_date and end_date <= existing_end_date) or \
                        (start_date <= existing_start_date and end_date >= existing_end_date):
                            messagebox.showerror("Error", f"{name} is already working within the selected date range")
                            return False

                resource_row = df_resource_table[df_resource_table['Employee Name'] == name]

                if not resource_row.empty:
                    total_cost_by_month = resource_row['Total Cost by Month'].sum()
                    start_date_month = start_date.month
                    start_date_year = start_date.year

                    business_workdays = count_business_workdays(start_date_year, start_date_month)
                    if business_workdays != 0:
                        monthly_cost = total_cost_by_month / business_workdays
                    else:
                        messagebox.showerror("Error", "Business workdays count is zero!")
                        return False

                    monthly_cost *= manday
                    new_row['Monthly Cost'] = monthly_cost

                existing_division_df = pd.read_excel(division_file)
                updated_df = pd.concat([existing_division_df, pd.DataFrame([new_row])], ignore_index=True)
                for index, row in updated_df.iterrows():
                    if row['From Date'] == calender:
                        updated_df.at[index, 'Till Date'] = end_date.strftime('%m-%d-%Y')
                    else:
                        updated_df.at[index, 'Till Date'] = (pd.to_datetime(row['From Date'], format='%m-%d-%Y') + pd.offsets.BDay(row['Manday'] - 1)).strftime('%m-%d-%Y')
                updated_df = updated_df[['Name', 'Project Name', 'Manday', 'From Date', 'Till Date', 'Monthly Cost']]
                updated_df.to_excel(division_file, index=False) 

                updated_insert_df = pd.concat([existing_insert_df, pd.DataFrame([new_row])], ignore_index=True)
                updated_insert_df['From Date'] = pd.to_datetime(updated_insert_df['From Date'], errors='coerce').dt.strftime('%m-%d-%Y')
                updated_insert_df['Till Date'] = updated_insert_df.apply(lambda x: end_date.strftime('%m-%d-%Y') if x['From Date'] == calender else (start_date + pd.offsets.BDay(x['Manday'] - 1)).strftime('%m-%d-%Y'), axis=1)
                for index, row in updated_insert_df.iterrows():
                    if row['From Date'] == calender:
                        updated_insert_df.at[index, 'Till Date'] = end_date.strftime('%m-%d-%Y')
                    else:
                        updated_insert_df.at[index, 'Till Date'] = (pd.to_datetime(row['From Date'], format='%m-%d-%Y') + pd.offsets.BDay(row['Manday'] - 1)).strftime('%m-%d-%Y')
                updated_insert_df = updated_insert_df[['Name', 'Project Name', 'Manday', 'From Date', 'Till Date', 'Monthly Cost']]
                updated_insert_df.to_excel(InsertTable, index=False)


                project_cost_series = updated_insert_df.groupby('Project Name')['Monthly Cost'].sum()
                for project_name, project_cost in project_cost_series.items():
                    df_project_table.loc[df_project_table['Project Name'] == project_name, 'Project Cost'] = project_cost
                df_project_table.to_excel(ProjectTable, index=False)


                messagebox.showinfo("Success", f"Data inserted into {employee_division} and Insert table successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

        save_button = tk.Button(self.add_window, text="Add", command=save_to_excel)
        save_button.place(x=210,y=430,width=148,height=42)

        back_button = tk.Button(self.add_window, text="Back", command=self.back_to_main_form)
        back_button.place(x=20,y=20,width=70,height=25)
        
        self.add_window.mainloop()
        

        
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
