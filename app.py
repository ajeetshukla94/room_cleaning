import pandas as pd
from flask import Flask, render_template, flash, url_for, request, make_response, jsonify, session,send_from_directory
from werkzeug.utils import secure_filename
import os, time
import io
import base64
import json
import datetime
import xlsxwriter
import os, sys, glob
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.styles import Border, Side
from flask import send_file
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from fnmatch import fnmatch
import time
from openpyxl.styles import PatternFill
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = 'file_upload_key'

MYDIR = os.path.dirname(__file__)
app.config['UPLOAD_FOLDER'] = "static/Report/"
equipmet_list = ['Dispensing Booth',
'Dispensing Scoop ( Small)','Dispensing Scoop (Large)',
'Spatula','Spatula' ,'Manufacturing Scoop',
'Blender Bin (350 liter)' ,'Blender Bin (25 liter)' ,
'Vibro Sifter II','sampling rod','sampling rod' 
]





sent_mail = False
server    = 'smtp.gmail.com'
port      =  587
username  =  "aajeetshk@gmail.com"
password  =  "ilbumnmnsnqletdk"
send_from = "aajeetshk@gmail.com"
send_to   = "ashish@pinpointengineers.co.in"

def send_mail(subject,text,files,file_name,isTls=True):
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(files, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename={}.xlsx'.format(file_name))
        msg.attach(part)
        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username,password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
@app.route("/")
def default():
    return make_response(render_template('login_page/login.html'),200)    
    
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == 'POST':
      form_data = request.form
      l_id = form_data['login']
      pwd = form_data['password']
      if(l_id.lower()=='admin'.lower() and pwd == 'admin'):
          print('inside if')
          session['username'] = l_id
          flash('Login Successful')
          return make_response(render_template('cleaning_room.html',equipmet_list  = equipmet_list,
                                msg = True, err = False, warn = False),200)
      else:
          print('inside else')
          flash('Invalid Credentials')
          return make_response(render_template("login_page/login.html", msg = False, err = True, warn = False),403)
    else:
        print('get request')
        
@app.route('/logout')
def logout():
    session.pop('username', None)
    session.pop('selected_file', None)
    global Selected_files
    Selected_file = None
    flash('Logout Successful')
    return make_response(render_template("login_page/login.html",msg = True, err = False, warn = False, message='Logout Successful'),200)

@app.route("/cleaning_room")
def cleaning_room():
    return make_response(render_template('cleaning_room.html',equipmet_list  = equipmet_list),200) 

@app.route("/submit_data")    
def submit_data():
   
    data            = request.args.get('params_data')
    data            = json.loads(data)
     
   
    temp_df         = pd.DataFrame(data)
    new_header      = temp_df.iloc[0] 
    temp_df         = temp_df[1:] 
    temp_df.columns = new_header
    
    temp_df.dropna(axis=1, how='all',inplace=True)    
    wb = Workbook()
    ws = wb.active    
    file_name = "Cleaning_room_report_{}.xlsx".format(str(datetime.datetime.today().strftime('%d_%m_%Y')))    
    end_column =temp_df.shape[1] +2   
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=end_column)
    ws["B" + str(2)] = 'ANNEXURE - I'
    ws["B" + str(2)].fill = PatternFill(start_color='00cc99', end_color='00cc99', fill_type="solid")
    currentCell = ws["B" + str(2)]
    currentCell.alignment = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells(start_row=3, start_column=2, end_row=4, end_column=end_column)
    ws["B" + str(3)] = 'Equipment List '
    ws["B" + str(3)].fill = PatternFill(start_color='ff9933', end_color='ff9933', fill_type="solid")
    currentCell = ws["B" + str(3)]
    currentCell.alignment = Alignment(horizontal='center', vertical='center')
    
    
    ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=end_column)
    ws["B" + str(5)] = 'List of Equipments and their Product Contact Surface Area  '
    ws["B" + str(5)].fill = PatternFill(start_color='ff99ff', end_color='ff99ff', fill_type="solid")
    currentCell = ws["B" + str(5)]
    currentCell.alignment = Alignment(horizontal='center', vertical='center')
    
    row = 6
    j=1
    for row_data in temp_df.columns:
        ws.cell(row=row, column=j+2, value=row_data)
        j=j+1
    row=row+1
    ws['B6']="SR.no"
    for row_data in temp_df.itertuples():
        for j in range(0,len(row_data) ):   
            ws.cell(row=row, column=j+2, value=row_data[j])
        row=row+1    
        
    final_working_directory =MYDIR + "/" + app.config['UPLOAD_FOLDER']+file_name
    wb.save(os.path.join(app.config['UPLOAD_FOLDER'],file_name))

    if sent_mail:
        send_mail(subject,text,final_working_directory,file_name) 
    d = {"error":"none",
         "file_name":file_name,
         "file_path":final_working_directory}
   
    return json.dumps(d)

if __name__ == '__main__':
    #app.debug = True
    app.run()

