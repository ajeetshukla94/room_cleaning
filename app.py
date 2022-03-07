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
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = 'file_upload_key'

MYDIR = os.path.dirname(__file__)
app.config['UPLOAD_FOLDER_REPORT'] = "static/Report/"
equipmet_list = ['Dispensing Booth',
'Dispensing Scoop ( Small)','Dispensing Scoop (Large)',
'Spatula','Spatula' ,'Manufacturing Scoop',
'Blender Bin (350 liter)' ,'Blender Bin (25 liter)' ,
'Vibro Sifter II','sampling rod','sampling rod' 
]

MYDIR = os.path.dirname(__file__)
app.config['UPLOAD_FOLDER_INPUTDATA'] = "static/inputData/"

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
    product_frame  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER_INPUTDATA'],"product_Details.xlsx"))
    product_list   = product_frame.Product_Name.unique().tolist()
    if request.method == 'POST':
      form_data = request.form
      l_id = form_data['login']
      pwd = form_data['password']
      if(l_id.lower()=='admin'.lower() and pwd == 'admin'):
          print('inside if')
          session['username'] = l_id
          flash('Login Successful')
          return make_response(render_template('cleaning_room.html',equipmet_list  = equipmet_list,product_list  = product_list,
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
    product_frame  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER_INPUTDATA'],"product_Details.xlsx"))
    product_list   = product_frame.Product_Name.unique().tolist()
    return make_response(render_template('cleaning_room.html',equipmet_list  = equipmet_list,product_list  = product_list),200) 
    
    
@app.route("/UpdateProductList")
def UpdateProductList():
    product_frame  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER_INPUTDATA'],"product_Details.xlsx"))
    product_list   = product_frame.to_dict('records')
    return make_response(render_template('UpdateProductList.html',product_list  = product_list),200) 
    
@app.route("/submit_UpdateProductList")    
def submit_UpdateProductList():   
    data            = request.args.get('params_data')
    data            = json.loads(data)  
    observation     = data['observation']
    temp_df         = pd.DataFrame.from_dict(observation,orient ='index')
    temp_df         = temp_df[['Product_Name','Generic_Name','Form', 'API_with_strength' ,'Minimum_Batch_size','MRDD','LD50','NOEL']]
      
    store_location = "static/inputData/"+"product_Details.xlsx"
    final_working_directory=MYDIR + "/" +store_location
    temp_df.to_excel(final_working_directory,index=False)
    d = {"error":"none",}
   
    return json.dumps(d)

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
    ws.title = 'Equipment List '    
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
        ws.cell(row=row, column=j+2).fill = PatternFill(start_color='66ccff', end_color='66ccff', fill_type="solid")
        j=j+1
    row=row+1
    ws['B6']="SR.no"
    ws['B6'].fill = PatternFill(start_color='66ccff', end_color='66ccff', fill_type="solid")
    for row_data in temp_df.itertuples():
        for j in range(0,len(row_data) ):   
            ws.cell(row=row, column=j+2, value=row_data[j])
            if j ==0:
                ws.cell(row=row, column=j+2).fill = PatternFill(start_color='669999', end_color='669999', fill_type="solid")
            else :
                ws.cell(row=row, column=j+2).fill = PatternFill(start_color='3399ff', end_color='3399ff', fill_type="solid")
                
        row=row+1    

    

    sheet_ranges = wb.active
    sheet_ranges.column_dimensions["B"].width = 10
    sheet_ranges.column_dimensions["C"].width = 35
    column = 4
    while column < end_column:
        i = get_column_letter(column)
        ws.column_dimensions[i].width = 10
        column += 1   
    #################################start of prodcut sheet##################################################
    product_sheet = wb.create_sheet('Product Details')
    
    product_frame  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER_INPUTDATA'],"product_Details.xlsx"))
    
    end_column =product_frame.shape[1] +2   
    product_sheet.merge_cells(start_row=2, start_column=2, end_row=2, end_column=end_column)
    product_sheet["B" + str(2)] = 'ANNEXURE - II'
    product_sheet["B" + str(2)].fill = PatternFill(start_color='00cc99', end_color='00cc99', fill_type="solid")
    currentCell = product_sheet["B" + str(2)]
    currentCell.alignment = Alignment(horizontal='center', vertical='center')
    
    product_sheet.merge_cells(start_row=3, start_column=2, end_row=4, end_column=end_column)
    product_sheet["B" + str(3)] = 'Product Details '
    product_sheet["B" + str(3)].fill = PatternFill(start_color='ff9933', end_color='ff9933', fill_type="solid")
    currentCell = product_sheet["B" + str(3)]
    currentCell.alignment = Alignment(horizontal='center', vertical='center')    
    row = 5
    j=1        
    for row_data in product_frame.columns:
        product_sheet.cell(row=row, column=j+2, value=row_data)
        product_sheet.cell(row=row, column=j+2).fill = PatternFill(start_color='ffcc99', end_color='ffcc99', fill_type="solid")
        j=j+1
    row=row+1
    product_sheet['B5']="SR.no"
    product_sheet['B5'].fill = PatternFill(start_color='ffcc99', end_color='ffcc99', fill_type="solid")
    for row_data in product_frame.itertuples():
        for j in range(0,len(row_data) ):   
            product_sheet.cell(row=row, column=j+2, value=row_data[j])
            if j ==0:
                product_sheet.cell(row=row, column=j+2).fill = PatternFill(start_color='669999', end_color='669999', fill_type="solid")
            else :
                product_sheet.cell(row=row, column=j+2).fill = PatternFill(start_color='3399ff', end_color='3399ff', fill_type="solid")
        row=row+1   
   
    sheet_ranges = wb['Product Details']
    sheet_ranges.column_dimensions["C"].width = 20
    sheet_ranges.column_dimensions["D"].width = 20
    sheet_ranges.column_dimensions["E"].width = 15
    sheet_ranges.column_dimensions["F"].width = 20
    sheet_ranges.column_dimensions["G"].width = 20
    
    
    store_location = "static/inputData/"+file_name
    final_working_directory=MYDIR + "/" +store_location
    #final_working_directory=store_location
    wb.save(final_working_directory)
    
    ###################################END of prodcut sheet#################################################

    if sent_mail:
        send_mail(subject,text,final_working_directory,file_name) 
    d = {"error":"none",
         "file_name":file_name,
         "file_path":store_location}
   
    return json.dumps(d)

if __name__ == '__main__':
    #app.debug = True
    app.run()

