from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from datetime import datetime
from collections import defaultdict

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Paths to Excel files
InsertTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/InsertTable.xlsx"
ResourceTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/ResourcesTable.xlsx"
MsdtDivisionTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/MsdtDivisionTable.xlsx"
CyberSecurityDivisionTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/CyberSecurityDivisionTable.xlsx"
CloudDivisionTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/CloudDivisionTable.xlsx"
ProjectTable = "C:/Users/Jubin/Desktop/Desktop/ALPHA_SOURCE/Alpha Data Project/ProjectTable.xlsx"

# Read Excel files
df_insert_table = pd.read_excel(InsertTable)
df_resource_table = pd.read_excel(ResourceTable)
df_msdtdivision_table = pd.read_excel(MsdtDivisionTable)
df_cybersecuritydivision_table = pd.read_excel(CyberSecurityDivisionTable)
df_clouddivision_table = pd.read_excel(CloudDivisionTable)
df_project_table = pd.read_excel(ProjectTable)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add_job', methods=['GET', 'POST'])
def add_job():
    if request.method == 'POST':
        name = request.form['name']
        project_name = request.form['project_name']
        date = request.form['date']
        man_hours = request.form['man_hours']

        # Perform validation and data insertion here
        
        flash('Job added successfully!', 'success')
        return redirect(url_for('add_job'))

    return render_template('add_job.html')

@app.route('/view_job')
def view_job():
    return render_template('view_job.html')

@app.route('/division')
def division():
    return render_template('division.html')

@app.route('/employee')
def employee():
    return render_template('employee.html')

@app.route('/project')
def project():
    return render_template('project.html')



if __name__ == '__main__':
    app.run(debug=True)
