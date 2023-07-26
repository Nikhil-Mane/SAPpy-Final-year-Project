import sys, os
if sys.executable.endswith('pythonw.exe'):
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.path.join(os.getenv('TEMP'), 'stderr-{}'.format(os.path.basename(sys.argv[0]))), "w")
    
import os
import io
import pandas as pd
import numpy as np
import ipywidgets as widgets
from IPython.display import display
import os
import flask
from flask import Flask, render_template, request, redirect, url_for, json, flash, send_file, send_from_directory, abort, session, jsonify
from pandas.core.frame import DataFrame
from wtforms.widgets.core import HiddenInput
from forms import DataframeForm
from werkzeug.utils import secure_filename
from flask import Response
import re
from random import random
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure
import pymongo
import uuid
from functools import wraps

app = Flask(__name__)
app.secret_key = "abc" 
app.config['SECRET_KEY'] = '9ed3e1b5b9bb74f177e81230412af077'
app.config["SESSION_PERMANENT"] = False
ALLOWED_EXTENSIONS = {"xlsx", "jpg", "csv", "xlsm", "xlsb", "xltx", "xltm"}
app.config["excel_uploads"] = r"C:\Users\Nikhil Mane\Desktop\SAPpy\Final Sappy (1)\static\exel"
app.config["get_result"] = r"C:\Users\Nikhil Mane\Desktop\SAPpy\Final Sappy (1)\static\exel\results\results"


client = pymongo.MongoClient('localhost', 27017)
db = client.user_login_system

url = " "
url_PO = " "
GRN_file = " "
PO_file = " "
buyers_list = []
CSV = pd.DataFrame()

def allowed_file(filename):
    if not "." in filename:
        return False
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    
# Decorators
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user" in session:
            return f(*args, **kwargs)
        else:
            flash(f'Enter Username and Password!', 'danger')
            return redirect('/')
    return decorated_function

@app.route("/", methods=['GET', 'POST'])
def login():
    if request.method == 'POST' and 'username' in request.form and 'password' in request.form:
        
        password = request.form['password']

        user = db.users.find_one({
            "username": request.form.get('username')
        })

        if user and password == user['password']:
            session["user"] = user
            flash(f'You have successfully logged In!', 'success')
            return redirect(url_for('home'))
        else:
            flash(f'Incorrect username/password!', 'danger')

    return render_template("login.html")

@app.route('/pythonlogin/register', methods=['GET', 'POST'])
def register():
    # Check if "username", "password" and "email" POST requests exist (user submitted form)
    if request.method == 'POST' and 'username' in request.form and 'password' in request.form and 'email' in request.form:

        user = {
            "_id": uuid.uuid4().hex,
            "username": request.form.get('username'),
            "email": request.form.get('email'),
            "password": request.form.get('password')
        }

        if db.users.find_one({"email" : user['email']}):
            flash(f'Email already exist!', 'danger')
            return redirect(request.url)
        if db.users.find_one({"username" : user['username']}):
            flash(f'Username already exist!', 'danger')
            return redirect(request.url)
        if db.users.insert_one(user):
            flash(f'You have successfully registered! Please login', 'success')
            return redirect('/')

    return render_template('register.html')


@app.route("/signout", methods=['GET', 'POST'])
def signout():
    session.clear()
    flash(f'You have successfully logged Out! Please login', 'success')
    return redirect('/')


@app.route("/home", methods=['GET', 'POST'])
@login_required
def home():
    form = DataframeForm()
    global GRN_file
    global PO_file
    global buyers_list
    global supplier_list
    
    if request.method == "POST":
        if request.files:
            if ('GRN' not in request.files) and ('PO' not in request.files):
                print('No file part')
                return redirect(request.url)
            excel = request.files["GRN"]
            excel_PO = request.files["PO"]
            if (excel.filename == "") and (excel_PO.filename == ""):
                flash(f'No file selected', 'warning')
                return redirect(request.url)
            if (excel and allowed_file(excel.filename)) and (excel_PO and allowed_file(excel_PO.filename)):
                GRN_file = secure_filename(excel.filename)
                PO_file = secure_filename(excel_PO.filename)
                excel.save(os.path.join(app.config['excel_uploads'], GRN_file))
                excel_PO.save(os.path.join(
                    app.config['excel_uploads'], PO_file))
                flash(
                    f'{GRN_file} & {PO_file} files successfully uploaded!', 'success')
                print(GRN_file)
                global url
                url = "C:\\Users\\Nikhil Mane\\Desktop\\SAPpy\\Final Sappy (1)\\static\exel\\%s" % (GRN_file)
                global url_PO
                url_PO = "C:\\Users\\Nikhil Mane\\Desktop\\SAPpy\\Final Sappy (1)\\static\\exel\\%s" % (PO_file)
                df = pd.read_excel(url)
                buyers = df.Buyer.unique()
                buyers = np.insert(buyers, 0, 'All Buyers')
                suppilers = df.Supplier.unique()
                suppilers = np.insert(suppilers, 0, 'All Suppliers')
                supplier_list = [(val, val) for val in suppilers]
                
                buyers_list = [(val, val) for val in buyers]
                return redirect(request.url)
            else:
                flash(f'this file extension is not allowed', 'warning')
                return redirect(request.url)

    return render_template('home.html', form=form)


def build_csv_data(dataframe):
    csv_data = dataframe.to_excel(
        r"C:\Users\Nikhil Mane\Desktop\SAPpy\Final Sappy (1)\static\exel\results\results.xlsx", index=False, encoding='utf-8')

    return csv_data


def preprocess(i):
    i["month"] = i["GRN Date"].dt.month
    i["day"] = i["GRN Date"].dt.day
    i.rename(columns={'Item No. & Desc.': 'Items'}, inplace=True)
    item = i['Items'].str.split('--', n=1, expand=True)
    i['Item No.'] = item[0]
    i['Item Description'] = item[1]
    i.drop(['Items'], axis=1)
    i.loc[(i['Item Description'] == 'HOTEL BILL') | (i['Item Description'] == '-LABOR CHARGE') | (i['Item Description'] == 'TRANSPORTATION CHARGE') | (i['Item Description'] ==
                                                                                                                                                       'TRANSPORTATION CHARGES') | (i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS') | (i['Item Description'] == 'LUNCH BILL')]
    i = i.drop((i[i['Item Description'] == 'HOTEL BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGE'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGES'].index) | (
        i[i['Item Description'] == '-LABOR CHARGE'].index) | (i[i['Item Description'] == 'LUNCH BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS'].index))
    j = i
    return j


def consumption_preprocess(dataframe):
    dataframe = preprocess(dataframe)
    dataframe['Acp. Qty.'] = dataframe['Counted Qty.'] - dataframe['Rej. Qty.']
    dataframe = dataframe[["Item No.", "Item Description", "Acp. Qty.",'Acp. UD Qty.',
                           "GRN Date", "Rate", "Buyer", "month", "day", "PO UOM", "Amt.","Currency"]]
    # return dataframe
    Desc = dataframe[["Item No.", "Item Description", "Acp. Qty.",'Acp. UD Qty.',
                           "GRN Date", "Rate", "Buyer", "month", "day", "PO UOM","Amt.", "Currency"]]
    #Desc.drop(Desc[Desc['Item Description'] != '-'].index, inplace = True)
    for index_label, row_series in Desc.iterrows():
        if  Desc.at[index_label , 'Item Description'] == '-':
            Desc.at[index_label , 'Item Description'] = row_series['Item No.']
        else:
            continue
#     for index_label, row_series in Desc.iterrows():
#    # For each row update the 'Bonus' value to it's double
#         Desc.at[index_label , 'Item Description'] = row_series['Item No.']
    return Desc


def avgconsumption(dataframe, val):
    if val == '5':
        items_group = dataframe.groupby(by='Item Description')
        df_yearly_consumption = pd.DataFrame({
            'Total_consumption': items_group['Acp. Qty.'].sum(),
            'PO UOM': items_group['PO UOM'].first(),
            'Average Consumption': items_group['Acp. Qty.'].mean(),
            'Average Half Yearly Quantity': items_group['Acp. Qty.'].sum()/2,
            'Average Quaterly Quantity': items_group['Acp. Qty.'].sum()/4,
            # 'Total Rate': items_group['Rate'].sum(),
            'Average Rate':  (items_group['Amt.'].sum() / (items_group['Acp. Qty.'].sum()+items_group['Acp. UD Qty.'].sum())),
            'Currency': items_group['Currency'].first(),
        })
        df_yearly_consumption = df_yearly_consumption.sort_values(by = ['Total_consumption'],ascending=False)
        df_yearly_consumption = df_yearly_consumption.reset_index()
        return df_yearly_consumption

def avgconsumption1(dataframe, val, name):
    if val == '5':
        dataframe = dataframe[dataframe['Buyer'].str.contains(name)]
        items_group = dataframe.groupby(by='Item Description')
        df_yearly_consumption = pd.DataFrame({
            'Total_consumption': items_group['Acp. Qty.'].sum(),
            'PO UOM': items_group['PO UOM'].first(),
            'Average Consumption': items_group['Acp. Qty.'].mean(),
            'Average Half Yearly Quantity': items_group['Acp. Qty.'].sum()/2,
            'Average Quaterly Quantity': items_group['Acp. Qty.'].sum()/4,
            # 'Total Rate': items_group['Rate'].sum(),
            'Average Rate':  (items_group['Amt.'].sum() / (items_group['Acp. Qty.'].sum()+items_group['Acp. UD Qty.'].sum())),
            'Currency': items_group['Currency'].first(),
        })
        df_yearly_consumption = df_yearly_consumption.sort_values(by = ['Total_consumption'],ascending=False)
        df_yearly_consumption = df_yearly_consumption.reset_index()
        return df_yearly_consumption

def quarterlyframe(dataframe, val):
    if val == '1':
        dataframe = dataframe[((dataframe['month'] >= 4) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 6) & (dataframe["day"] <= 30))]
    if val == '2':
        dataframe = dataframe[((dataframe['month'] >= 7) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 9) & (dataframe["day"] <= 30))]
    if val == '3':
        dataframe = dataframe[((dataframe['month'] >= 10) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 12) & (dataframe["day"] <= 31))]
    if val == '4':
        dataframe = dataframe[((dataframe['month'] >= 1) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 3) & (dataframe["day"] <= 31))]
    items_group = dataframe.groupby(by="Item Description")
    df_yearly_consumption = pd.DataFrame({
        'Total Rate': items_group['Rate'].sum(),
        'Currency': items_group['Currency'].first(),
        'Total consumption': items_group['Acp. Qty.'].sum(),
        'PO UOM': items_group['PO UOM'].first(),
    })
    df_yearly_consumption = df_yearly_consumption.sort_values(by = ['Total consumption'],ascending=False)
    df_yearly_consumption = df_yearly_consumption.reset_index()
    return df_yearly_consumption

def quarterlyframe1(dataframe, val, name):
    dataframe = dataframe[dataframe['Buyer'].str.contains(name)]
    if val == '1':
        dataframe = dataframe[((dataframe['month'] >= 4) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 6) & (dataframe["day"] <= 30))]
    if val == '2':
        dataframe = dataframe[((dataframe['month'] >= 7) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 9) & (dataframe["day"] <= 30))]
    if val == '3':
        dataframe = dataframe[((dataframe['month'] >= 10) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 12) & (dataframe["day"] <= 31))]
    if val == '4':
        dataframe = dataframe[((dataframe['month'] >= 1) & (dataframe["day"] >= 1)) & (
            (dataframe['month'] <= 3) & (dataframe["day"] <= 31))]
    items_group = dataframe.groupby(by="Item Description")
    df_yearly_consumption = pd.DataFrame({
        'Total Rate': items_group['Rate'].sum(),
        'Currency': items_group['Currency'].first(),
        'Total consumption': items_group['Acp. Qty.'].sum(),
        'PO UOM': items_group['PO UOM'].first(),
    })
    df_yearly_consumption = df_yearly_consumption.sort_values(by = ['Total consumption'],ascending=False)
    df_yearly_consumption = df_yearly_consumption.reset_index()
    return df_yearly_consumption

def duplicate_preprocess(dataframe):
    dataframe = preprocess(dataframe)
    dataframe["Item Description"] = dataframe["Item Description"].str.lstrip(
        "-")
    dataframe = dataframe[["Supplier", "Rate", "Item No.", "Item Description"]]
    return dataframe


def duplicate_code(dataframe, name):
    # k = dataframe[dataframe.duplicated(subset='Item Description', keep=False)]
    dataframe.drop_duplicates(subset="Item No.", keep="first", inplace=True)
    df = dataframe[dataframe.duplicated(subset='Item Description', keep=False)]
    Desc = df.sort_values(by=['Item Description'])
    Desc = Desc.iloc[3:]
    return Desc


def highlight(s):
    if (s['Acp. rate (in %)'] <= 100.0) and (s['Acp. rate (in %)'] > 90.0):
        return ['background-color: green']*5
    elif (s['Acp. rate (in %)'] <= 90.0) and (s['Acp. rate (in %)'] > 70.0):
        return ['background-color: yellow']*5
    elif s['Acp. rate (in %)'] <= 70.0:
        return ['background-color: red']*5


def Supplier_Ranking(Buyer, Ranked_list, by, sortby):
    df1 = pd.read_excel(url)
    df2 = pd.read_excel(url_PO)
    if by == "Item Quality":
        if Buyer != "NULL":
            df1 = df1[df1['Buyer'].str.contains(Buyer)]
            df2 = df2[df2['Buyer'].str.contains(Buyer)]
        dff = pd.merge(df1, df2, on='Supplier', how='left')
        dff['Buyer'] = dff['Buyer_x']
        df3 = dff.groupby(['Buyer', 'Supplier']).sum()
        df3['Acp. rate (in %)'] = ((df3['Acp. Qty.']/df3['Counted Qty.'])*100).round(decimals=2)
        df3['Acp. UD. rate (in %)'] = ((df3['Acp. UD Qty.']/df3['Counted Qty.'])*100).round(decimals=2)
        df3['Rej. rate (in %)'] = ((df3['Rej. Qty.']/df3['Counted Qty.'])*100).round(decimals=2)
        df3 = df3.reset_index()
        if sortby == "Best to worst":
            df4 = df3.sort_values(
                by=['Acp. rate (in %)', 'Acp. UD. rate (in %)'], ascending=False)
        elif sortby == "Worst to best":
            df4 = df3.sort_values(
                by=['Acp. rate (in %)', 'Acp. UD. rate (in %)'])
        df4 = df4[['Buyer', 'Supplier',
                   'Acp. rate (in %)', 'Acp. UD. rate (in %)', 'Rej. rate (in %)']]

        if Ranked_list == "yes":
            grp = df4.groupby(by="Supplier")
            Ranked_supplier = []
            for supplier, value in grp:
                Ranked_supplier.append(supplier)
            return Ranked_supplier
        df4['Acp. rate (in %)'].round(decimals=2)
        return(df4)
    elif by == "In-time Delivery":
        if Buyer != "NULL":
            df1 = df1[df1['Buyer'].str.contains(Buyer)]
            df2 = df2[df2['Buyer'].str.contains(Buyer)]
        dff = pd.merge(df1, df2, on='Supplier', how='left')
        dff['Delayed delivery (in days)'] = dff['GRN Date'] - \
            dff['Sch. Date']
        dff['Delayed delivery (in days)'] = pd.to_numeric(
            dff['Delayed delivery (in days)'].dt.days, downcast='integer')
        dff['Frequency of order'] = 1
        dff['Buyer'] = dff['Buyer_x']
        dff1 = dff[['Buyer', 'Supplier',
                    'Delayed delivery (in days)', 'Frequency of order']]
        dfff = dff1.groupby(['Supplier', 'Buyer']).sum()
        dfff['Delivery rate (in %)'] = (dfff['Delayed delivery (in days)'] / \
            dfff['Frequency of order']).round(decimals=2)
        dfff = dfff.reset_index()
        if sortby == "Best to worst":
            dfff = dfff.sort_values(by='Delayed delivery (in days)')
        elif sortby == "Worst to best":
            dfff = dfff.sort_values(
                by='Delayed delivery (in days)', ascending=False)
        if Ranked_list == "yes":
            grp = dfff.groupby(by="Supplier")
            Ranked_supplier = []
            for supplier, value in grp:
                Ranked_supplier.append(supplier)
            return Ranked_supplier
        return dfff

def AllSupplier_Ranking(Ranked_list, by, sortby):
    df1 = pd.read_excel(url)
    df2 = pd.read_excel(url_PO)
    if by == "Item Quality":
        dff = pd.merge(df1, df2, on='Supplier', how='left')
        dff['Buyer'] = dff['Buyer_x']
        df3 = dff.groupby(['Supplier']).sum()
        df3['Acp. rate (in %)'] = ((df3['Acp. Qty.']/df3['Counted Qty.'])*100).round(decimals=2)
        df3['Acp. UD. rate (in %)'] = ((df3['Acp. UD Qty.']/df3['Counted Qty.'])*100).round(decimals=2)
        df3['Rej. rate (in %)'] = ((df3['Rej. Qty.']/df3['Counted Qty.'])*100).round(decimals=2)
        df3 = df3.reset_index()
        if sortby == "Best to worst":
            df4 = df3.sort_values(
                by=['Acp. rate (in %)', 'Acp. UD. rate (in %)'], ascending=False)
        elif sortby == "Worst to best":
            df4 = df3.sort_values(
                by=['Acp. rate (in %)', 'Acp. UD. rate (in %)'])
        df4 = df4[['Supplier',
                   'Acp. rate (in %)', 'Acp. UD. rate (in %)', 'Rej. rate (in %)']]

        if Ranked_list == "yes":
            grp = df4.groupby(by="Supplier")
            Ranked_supplier = []
            for supplier, value in grp:
                Ranked_supplier.append(supplier)
            return Ranked_supplier
        df4['Acp. rate (in %)'].round(decimals=2)
        return(df4)
    elif by == "In-time Delivery":
        dff = pd.merge(df1, df2, on=['Supplier','PO No.'], how='left')
        dff['Delayed delivery (in days)'] = dff['GRN Date'] - \
            dff['Sch. Date']
        dff['Delayed delivery (in days)'] = pd.to_numeric(
            dff['Delayed delivery (in days)'].dt.days, downcast='integer')
        dff['Frequency of order'] = 1
        dff['Buyer'] = dff['Buyer_x']
        rdff = dff[['Supplier','Delayed delivery (in days)']]
        new = dff[['Supplier','Buyer','Frequency of order','Item No. & Desc.','PO No.','GRN No.']]
        new = new.drop_duplicates(subset=['Supplier','Buyer','PO No.','GRN No.','Item No. & Desc.'])
        newdf = new.groupby(['Supplier']).sum()
        df2 = rdff.groupby(['Supplier']).mean().round(decimals=0)
        dff1 = pd.merge(df2,newdf,on=['Supplier'],how='left')
        dff1['Average Delay (In Days)'] = (dff1['Delayed delivery (in days)'] / \
            dff1['Frequency of order']).round(decimals=2)
        # dfff = dff1.groupby(['Supplier']).mean().round(decimals=0)
        dfff = dff1
        dfff['Delayed delivery (in days)'].round()
        dfff['Average Delay (In Days)'].round(decimals=2)
        dfff = dfff.reset_index()
        if sortby == "Best to worst":
            dfff = dfff.sort_values(by='Delayed delivery (in days)')
        elif sortby == "Worst to best":
            dfff = dfff.sort_values(
                by='Delayed delivery (in days)', ascending=False)
        dfff = dfff[['Supplier','Delayed delivery (in days)','Frequency of order','Average Delay (In Days)']]
        dfff = dfff.dropna()
        if Ranked_list == "yes":
            grp = dfff.groupby(by="Supplier")
            Ranked_supplier = []
            for supplier, value in grp:
                Ranked_supplier.append(supplier)
            return Ranked_supplier
        return dfff

        


def frequency_(p, Q, QC, R):
    # Create a group w.r.t. items
    items_group = p.groupby(by="Item Description")
    # Create a dataframe holding items, frequency count
    v = pd.DataFrame({
        Q: items_group['Item No.'].count(),
        QC: items_group['Challan Qty.'].sum(),
        R: items_group['Rate'].mean()
    })
    v = v.reset_index()
    v[R] = v[R].round(decimals=3)
    return v
    # View the dataframe
    # frequency


def merge(v):
    for i in v:
        if i is v[0]:
            frequency = v[0]
        else:
            frequency = pd.merge(
                frequency, i, on='Item Description', how='outer')
    return frequency


def final_result(v):
    frequency = v.fillna(0)
    frequency['max_freq_count'] = np.where(
        frequency['Q1_apr_jun'] > frequency['Q2_jul_sep'], frequency['Q1_apr_jun'], frequency['Q2_jul_sep'])
    frequency['max_freq_quater'] = np.where(
        frequency['Q1_apr_jun'] > frequency['Q2_jul_sep'], 'Q1', 'Q2')
    frequency['max_freq_count'] = np.where(
        frequency['Q3_oct_dec'] > frequency['max_freq_count'], frequency['Q3_oct_dec'], frequency['max_freq_count'])
    frequency['max_freq_quater'] = np.where(
        frequency['max_freq_count'] > frequency['Q3_oct_dec'], frequency['max_freq_quater'], 'Q3')
    frequency['max_freq_count'] = np.where(
        frequency['max_freq_count'] > frequency['Q4_jan_mar'], frequency['max_freq_count'], frequency['Q4_jan_mar'])
    frequency['max_freq_quater'] = np.where(
        frequency['max_freq_count'] > frequency['Q4_jan_mar'], frequency['max_freq_quater'], 'Q4')
    frequency['max_order_qty'] = np.where(
        frequency['Q1_ordered_qty'] > frequency['Q2_ordered_qty'], frequency['Q1_ordered_qty'], frequency['Q2_ordered_qty'])
    frequency['max_order_quarter'] = np.where(
        frequency['Q1_ordered_qty'] > frequency['Q2_ordered_qty'], 'Q1', 'Q2')
    frequency['max_order_qty'] = np.where(
        frequency['Q3_ordered_qty'] > frequency['max_order_qty'], frequency['Q3_ordered_qty'], frequency['max_order_qty'])
    frequency['max_order_quarter'] = np.where(
        frequency['max_order_qty'] > frequency['Q3_ordered_qty'], frequency['max_order_quarter'], 'Q3')
    frequency['max_order_qty'] = np.where(
        frequency['max_order_qty'] > frequency['Q4_ordered_qty'], frequency['max_order_qty'], frequency['Q4_ordered_qty'])
    frequency['max_order_quarter'] = np.where(
        frequency['max_order_qty'] > frequency['Q4_ordered_qty'], frequency['max_order_quarter'], 'Q4')
    return frequency


def obj_frequency(name, sort):
    df = pd.read_excel(url)
    df = preprocess(df)
    df = df[["GRN Date", "Supplier", "Buyer", "PO No.", "Item No.", "Item Description", "Challan Qty.",
             "Counted Qty.", "Acp. Qty.", "Acp. UD Qty.", "Rej. Qty.", "Rate", "Amt.", "Currency", "month", "day"]]
    if name != "NULL":
        df = df[df['Buyer'].str.contains(name)]
    df1 = df[((df['month'] >= 4) & (df["day"] >= 1)) &
             ((df['month'] <= 6) & (df["day"] <= 30))]
    df2 = df[((df['month'] >= 7) & (df["day"] >= 1)) &
             ((df['month'] <= 9) & (df["day"] <= 30))]
    df3 = df[((df['month'] >= 10) & (df["day"] >= 1)) &
             ((df['month'] <= 12) & (df["day"] <= 31))]
    df4 = df[((df['month'] >= 1) & (df["day"] >= 1)) &
             ((df['month'] <= 3) & (df["day"] <= 31))]
    frequency = frequency_(
        df, Q='full_yr', QC='fullyr_ordered_qty', R='avg_yr_rate')
    frequency1 = frequency_(df1, Q='Q1_apr_jun',
                            QC='Q1_ordered_qty', R='avg_Q1_rate')
    frequency2 = frequency_(df2, Q='Q2_jul_sep',
                            QC='Q2_ordered_qty', R='avg_Q2_rate')
    frequency3 = frequency_(df3, Q='Q3_oct_dec',
                            QC='Q3_ordered_qty', R='avg_Q3_rate')
    frequency4 = frequency_(df4, Q='Q4_jan_mar',
                            QC='Q4_ordered_qty', R='avg_Q4_rate')
    merged_freq = merge(
        [frequency, frequency1, frequency2, frequency3, frequency4])
    freq_final = final_result(merged_freq)
    freq_final = freq_final[["Item Description", "fullyr_ordered_qty", "Q1_ordered_qty", "Q2_ordered_qty", "Q3_ordered_qty",
                             "Q4_ordered_qty", "max_freq_count", "max_freq_quater", "max_order_qty", "max_order_quarter"]]

    freq_final = freq_final.sort_values(by=["fullyr_ordered_qty"],ascending=False )
    freq_final = freq_final.rename(columns = {
        'Item Description':'Item Description',
        'fullyr_ordered_qty' : 'Full Year Ordered Qty',
        'Q1_ordered_qty' : 'Quater 1 Qty',
        'Q2_ordered_qty' : 'Quater 2 Qty',
        'Q3_ordered_qty' : 'Quater 3 Qty',
        'Q4_ordered_qty' : 'Quater 4 Qty',
        'max_freq_count' : 'Maximum Frequency',
        'max_freq_quater' : 'Maximum Frequency Quater',
        'max_order_qty' : 'Maximum Order Qty',
        'max_order_quarter' : 'Maximum Order Quater'}, inplace = False)
    # if sort == "Ascending":
    #     freq_final = freq_final.sort_values(by=["max_order_qty"])
    # elif sort == "Descending":
    #     freq_final = freq_final.sort_values(
    #         by=["max_order_qty"], ascending=False)
    return freq_final

def obj_frequency1(sort):
    df = pd.read_excel(url)
    df = preprocess(df)
    df = df[["GRN Date", "Supplier", "Buyer", "PO No.", "Item No.", "Item Description", "Challan Qty.",
             "Counted Qty.", "Acp. Qty.", "Acp. UD Qty.", "Rej. Qty.", "Rate", "Amt.", "Currency", "month", "day"]]

    df1 = df[((df['month'] >= 4) & (df["day"] >= 1)) &
             ((df['month'] <= 6) & (df["day"] <= 30))]
    df2 = df[((df['month'] >= 7) & (df["day"] >= 1)) &
             ((df['month'] <= 9) & (df["day"] <= 30))]
    df3 = df[((df['month'] >= 10) & (df["day"] >= 1)) &
             ((df['month'] <= 12) & (df["day"] <= 31))]
    df4 = df[((df['month'] >= 1) & (df["day"] >= 1)) &
             ((df['month'] <= 3) & (df["day"] <= 31))]
    frequency = frequency_(
        df, Q='full_yr', QC='fullyr_ordered_qty', R='avg_yr_rate')
    frequency1 = frequency_(df1, Q='Q1_apr_jun',
                            QC='Q1_ordered_qty', R='avg_Q1_rate')
    frequency2 = frequency_(df2, Q='Q2_jul_sep',
                            QC='Q2_ordered_qty', R='avg_Q2_rate')
    frequency3 = frequency_(df3, Q='Q3_oct_dec',
                            QC='Q3_ordered_qty', R='avg_Q3_rate')
    frequency4 = frequency_(df4, Q='Q4_jan_mar',
                            QC='Q4_ordered_qty', R='avg_Q4_rate')
    merged_freq = merge(
        [frequency, frequency1, frequency2, frequency3, frequency4])
    freq_final = final_result(merged_freq)
    freq_final = freq_final[["Item Description", "fullyr_ordered_qty", "Q1_ordered_qty", "Q2_ordered_qty", "Q3_ordered_qty",
                             "Q4_ordered_qty", "max_freq_count", "max_freq_quater", "max_order_qty", "max_order_quarter"]]

    freq_final = freq_final.sort_values(by=["fullyr_ordered_qty"],ascending=False )
    freq_final = freq_final.rename(columns = {
        'Item Description':'Item Description',
        'fullyr_ordered_qty' : 'Full Year Ordered Qty',
        'Q1_ordered_qty' : 'Quater 1 Qty',
        'Q2_ordered_qty' : 'Quater 2 Qty',
        'Q3_ordered_qty' : 'Quater 3 Qty',
        'Q4_ordered_qty' : 'Quater 4 Qty',
        'max_freq_count' : 'Maximum Frequency',
        'max_freq_quater' : 'Maximum Frequency Quater',
        'max_order_qty' : 'Maximum Order Qty',
        'max_order_quarter' : 'Maximum Order Quater'}, inplace = False)
    # if sort == "Ascending":
    #     freq_final = freq_final.sort_values(by=["max_order_qty"])
    # elif sort == "Descending":
    #     freq_final = freq_final.sort_values(
    #         by=["max_order_qty"], ascending=False)
    return freq_final


def stocking_preprocess(i, j, name):
    i = preprocess(i)
    i['Acp. Qty.'] = i['Counted Qty.'] - i['Rej. Qty.']
    j['PO No.']=j['PO No.'].str.lstrip()
    dataframe = pd.merge(i, j, on='PO No.', how='left')
    dataframe = dataframe[["Item No.", "Item Description", "Acp. Qty.",
                           "GRN Date", "Rate", "Buyer_x", "PO Date"]]
    dataframe = dataframe[dataframe['Buyer_x'].str.contains(name)]
    dataframe["Lead time in days"] = (
        dataframe["GRN Date"]-dataframe["PO Date"]).dt.days
    dataframe = dataframe[dataframe['PO Date'].notna()]
    items_group = dataframe.groupby(by="Item Description")
    df_yearly_consumption = pd.DataFrame({
        'Total_consumption': items_group['Acp. Qty.'].sum(),
        'Maximum consumption': items_group['Acp. Qty.'].max(),
        'Avg_total_consumption': items_group['Acp. Qty.'].mean(),
        'Total Rate': items_group['Rate'].sum(),
        'Avg Rate': items_group['Rate'].mean(),
        'Avg Carrying cost': items_group['Rate'].mean(),
        'frequency': items_group['Item Description'].count(),
        'Average daily consumption': items_group['Acp. Qty.'].mean()/365,
        'Maximum daily consumption': items_group['Acp. Qty.'].max()/365,
        'Minimum daily consumption': items_group['Acp. Qty.'].min()/365,
        'Average Lead Time': items_group['Lead time in days'].mean(),
        'Maximum Lead Time': items_group['Lead time in days'].max(),
        'Minimum Lead Time': items_group['Lead time in days'].min(),
    })
    df_yearly_consumption = df_yearly_consumption.reset_index()
    df_yearly_consumption['demand_std'] = df_yearly_consumption['Average daily consumption'].std()
    df_yearly_consumption['lt_std'] = df_yearly_consumption['Average Lead Time'].std(
    )
    df_yearly_consumption['avg_lt'] = np.sqrt(
        (df_yearly_consumption['Average Lead Time']))
    return df_yearly_consumption

def Allstocking_preprocess(i, j):
    i = preprocess(i)
    i['Acp. Qty.'] = i['Counted Qty.'] - i['Rej. Qty.']
    j['PO No.']=j['PO No.'].str.lstrip()
    dataframe = pd.merge(i, j, on='PO No.', how='left')
    dataframe = dataframe[["Item No.", "Item Description", "Acp. Qty.",
                           "GRN Date", "Rate", "Buyer_x", "PO Date"]]
    dataframe["Lead time in days"] = (
        dataframe["GRN Date"]-dataframe["PO Date"]).dt.days
    dataframe = dataframe[dataframe['PO Date'].notna()]
    items_group = dataframe.groupby(by="Item Description")
    df_yearly_consumption = pd.DataFrame({
        'Total_consumption': items_group['Acp. Qty.'].sum(),
        'Maximum consumption': items_group['Acp. Qty.'].max(),
        'Avg_total_consumption': items_group['Acp. Qty.'].mean(),
        'Total Rate': items_group['Rate'].sum(),
        'Avg Rate': items_group['Rate'].mean(),
        'Avg Carrying cost': items_group['Rate'].mean(),
        'frequency': items_group['Item Description'].count(),
        'Average daily consumption': items_group['Acp. Qty.'].mean()/365,
        'Maximum daily consumption': items_group['Acp. Qty.'].max()/365,
        'Minimum daily consumption': items_group['Acp. Qty.'].min()/365,
        'Average Lead Time': items_group['Lead time in days'].mean(),
        'Maximum Lead Time': items_group['Lead time in days'].max(),
        'Minimum Lead Time': items_group['Lead time in days'].min(),
    })
    df_yearly_consumption = df_yearly_consumption.reset_index()
    df_yearly_consumption['demand_std'] = df_yearly_consumption['Average daily consumption'].std()
    df_yearly_consumption['lt_std'] = df_yearly_consumption['Average Lead Time'].std(
    )
    df_yearly_consumption['avg_lt'] = np.sqrt(
        (df_yearly_consumption['Average Lead Time']))
    return df_yearly_consumption


# function for categorizing items into class A, B, C


def ABC_segmentation(perc):
    '''
    Creates the 3 classes A, B, and C based
    on quantity percentages (A-70%, B-25%, C-15%)
    '''
    if perc > 0 and perc < 0.7:
        return 'A'
    elif perc >= 0.7 and perc < 0.85:
        return 'B'
    elif perc >= 0.85:
        return 'C'


def ABC_apply(dataframe):
    # create the column of the running CumCost of the cumulative cost
    dataframe['RunCumCost'] = dataframe['Total Rate'].cumsum()
    # create the column of the total sum
    dataframe['TotSum'] = dataframe['Total Rate'].sum()
    # create the column of the running percentage
    dataframe['RunPerc'] = dataframe['RunCumCost']/dataframe['TotSum']
    # create the column of the class
    dataframe['Class'] = dataframe['RunPerc'].apply(ABC_segmentation)
    return dataframe


def XYZ_segmentation(perc):
    "Creating three classes X, Y and Z based on frequency"
    # X = Runner, Y = Repeater, Z = Stranger
    if perc >= 0.75:
        return 'Runner'
    elif perc < 0.75 and perc > 0.4:
        return 'Repeater'
    elif perc <= 0.4:
        return 'Stranger'


def XYZ_apply(dataframe):
    dataframe['Total'] = dataframe['frequency'].sum()
    # create the column of the running percentage
    dataframe['Percent'] = dataframe['frequency']/dataframe['Total']*100
    # create the column of the class
    dataframe['Class'] = dataframe['Percent'].apply(XYZ_segmentation)
    return dataframe


def choose_analysis(dataframe, a):
    if a == "ABC":
        dataframe = ABC_apply(dataframe)
    elif a == "XYZ":
        dataframe = XYZ_apply(dataframe)
    return dataframe


def safetystocklevel(dataframe, c):
    # change with items(in analysis)
    # use widgets here
    # A: z=2.05, B: z=1.34, C: z=1.04
    for i, row in dataframe["Class"].iteritems():
        if str(row) == "A" or str(row) == "Runner":
            z = 2.05
        elif str(row) == "B" or str(row) == "Repeater":
            z = 1.34
        elif str(row) == "C" or str(row) == "Stranger":
            z = 1.04
        if c == 1:
            dataframe["safetystock"] = dataframe['Maximum daily consumption']*dataframe['Maximum Lead Time'] - \
                dataframe['Average daily consumption'] * \
                dataframe['Average Lead Time']
        # Normal Distribution with uncertainty about the demand
        elif c == 2:
            dataframe["safetystock"] = z * \
                dataframe["demand_std"]*dataframe["avg_lt"]
        # Normal distribution with uncertainty on the lead time
        elif c == 3:
            dataframe["safetystock"] = z * \
                dataframe['Average daily consumption']*dataframe["lt_std"]
        # Normal distribution with uncertainty on-demand and independent lead time
        elif c == 4:
            s1 = dataframe['Average daily consumption'] * \
                dataframe["lt_std"]*dataframe["lt_std"]
            s2 = dataframe['Average Lead Time'] * \
                dataframe["demand_std"]*dataframe["demand_std"]
            ss = np.sqrt((s1*s2))
            dataframe["safetystock"] = z*ss
        # Normal distribution with uncertainty on demand and dependent lead time
        elif c == 5:
            dataframe["safetystock"] = z*dataframe["demand_std"]*dataframe["avg_lt"] + \
                z*dataframe['Average daily consumption']*dataframe["lt_std"]
        # print("this is z value",z)
        # print("this is demand_std", dataframe["demand_std"])
        # print("this is avg_lt", dataframe["avg_lt"])
        # print("this is lt_std", dataframe["lt_std"])
    return dataframe


def price_fluctuation(v):

    total_acp_rate = ((v["avg_Q2_rate"]-v["avg_Q1_rate"])/v["avg_Q1_rate"])*100
    total_acp_rate = round(total_acp_rate, 3)
    v["Rate shift Q2"] = total_acp_rate
    v["Rate shift Q2"] = np.where((v["Q2_ordered_qty"] == 0.0) | (
        v["Q1_ordered_qty"] == 0.0), 100, v["Rate shift Q2"])
    total_acp_rate = ((v["avg_Q3_rate"]-v["avg_Q1_rate"])/v["avg_Q1_rate"])*100
    total_acp_rate = round(total_acp_rate, 3)
    v["Rate shift Q3"] = total_acp_rate
    v["Rate shift Q3"] = np.where((v["Q3_ordered_qty"] == 0.0) | (
        v["Q1_ordered_qty"] == 0.0), 100, v["Rate shift Q3"])
    total_acp_rate = ((v["avg_Q4_rate"]-v["avg_Q1_rate"])/v["avg_Q1_rate"])*100
    total_acp_rate = round(total_acp_rate, 3)
    v["Rate shift Q4"] = total_acp_rate
    v["Rate shift Q4"] = np.where((v["Q4_ordered_qty"] == 0.0) | (
        v["Q1_ordered_qty"] == 0.0), 100, v["Rate shift Q4"])
    return v


def stocking_quarter(v, w, c):
   # c= input('''Select the period in which you would like to perform Cost Optimization
    # 1. Cost optimization for max ordered quantity
    # 2.Q1
    # 3.Q2
    # 4.Q3
    # 5.Q4''')

    if c == 1:
        stocking_quarter = []
        money_saved = []
        for i in v.iterrows():
            EOQ = i[1][23]
            t = i[1][19]
            if EOQ == None:
                m = i[1][18]*w
                n = int(m)
                if m-n >= 0.5:
                    EOQ = n+1
                elif m-n < 0.5:
                    EOQ = n
            if (t == "Q2"):
                if(i[1][20] <= 0.000):
                    stocking_quarter.append("Q2")
                    money_saved.append(0)
                elif (i[1][20] > 0.000 and i[1][20] != 100.000):
                    stocking_quarter.append("Q1")
                    x = EOQ*i[1][6]
                    y = EOQ*i[1][9]
                    saved = y-x
                    saved = round(saved, 3)
                    money_saved.append(saved)
                elif (i[1][20] == 100.000):
                    money_saved.append(0)
                    if i[1][8] == 0:
                        stocking_quarter.append('Null')
                    elif i[1][5] == 0:
                        stocking_quarter.append('Q2')
            elif (t == "Q3"):
                if(i[1][21] <= i[1][20] and i[1][21] != 100):
                    stocking_quarter.append("Q3")
                    money_saved.append(0)
                elif (i[1][21] > i[1][20] and i[1][21] != 100.000):
                    stocking_quarter.append("Q2")
                    x = EOQ*i[1][9]
                    y = EOQ*i[1][12]
                    saved = y-x
                    saved = round(saved, 3)
                    money_saved.append(saved)
                elif(i[1][21] == 100 and i[1][20] == 100):
                    money_saved.append(0)
                    if i[1][11] == 0:
                        stocking_quarter.append('Null')
                    else:
                        stocking_quarter.append("Q3")
                    continue
                elif (i[1][21] == 100.000):
                    money_saved.append(0)
                    if i[1][11] == 0:
                        stocking_quarter.append('Null')
                    elif i[1][8] == 0:
                        stocking_quarter.append('Q3')
                        continue
            elif(t == "Q4"):
                if(i[1][22] <= i[1][21] and i[1][22] != 100):
                    stocking_quarter.append("Q4")
                    money_saved.append(0)
                elif (i[1][22] > i[1][21] and i[1][22] != 100.000):
                    stocking_quarter.append("Q3")
                    x = EOQ*i[1][12]
                    y = EOQ*i[1][15]
                    saved = y-x
                    saved = round(saved, 3)
                    money_saved.append(saved)
                elif (i[1][22] == 100.000 and i[1][21] == 100.000):
                    money_saved.append(0)
                    if i[1][14] == 0:
                        stocking_quarter.append('Null')
                    else:
                        stocking_quarter.append("Q4")
                    continue
                elif (i[1][22] == 100.000):
                    money_saved.append(0)
                    if i[1][14] == 0:
                        stocking_quarter.append('Null')
                    elif i[1][11] == 0:
                        stocking_quarter.append('Q4')
                        continue
            elif(t == "Q1"):
                stocking_quarter.append("Q1")
                money_saved.append(0)

        v["stocking_quarter"] = [x for x in stocking_quarter]
        v["money_saved"] = [x for x in money_saved]
        return v
    elif c == 2:
        stocking_quarter = []
        money_saved = []
        for i in v.iterrows():
            EOQ = i[1][23]
            if EOQ == None:
                m = i[1][8]*w
                n = int(m)
                if m-n >= 0.5:
                    EOQ = n+1
                elif m-n < 0.5:
                    EOQ = n
            if(i[1][20] <= 0.000):
                stocking_quarter.append("Q2")
                money_saved.append(0)
            elif (i[1][20] > 0.000 and i[1][20] != 100.000):
                stocking_quarter.append("Q1")
                x = EOQ*i[1][6]
                y = EOQ*i[1][9]
                saved = y-x
                saved = round(saved, 3)
                money_saved.append(saved)
            elif (i[1][20] == 100.000):
                money_saved.append(0)
                if i[1][8] == 0:
                    stocking_quarter.append('Null')
                elif i[1][5] == 0:
                    stocking_quarter.append('Q2')
        v["stocking_quarter"] = [x for x in stocking_quarter]
        v["money_saved"] = [x for x in money_saved]
        return v
    elif c == 3:
        stocking_quarter = []
        money_saved = []
        for i in v.iterrows():
            EOQ = i[1][23]
            if EOQ == None:
                m = i[1][11]*w
                n = int(m)
                if m-n >= 0.5:
                    EOQ = n+1
                elif m-n < 0.5:
                    EOQ = n
            if(i[1][21] <= i[1][20] and i[1][21] != 100):
                stocking_quarter.append("Q3")
                money_saved.append(0)
            elif (i[1][21] > i[1][20] and i[1][21] != 100.000):
                stocking_quarter.append("Q2")
                x = EOQ*i[1][9]
                y = EOQ*i[1][12]
                saved = y-x
                saved = round(saved, 3)
                money_saved.append(saved)
            elif(i[1][21] == 100 and i[1][20] == 100):
                money_saved.append(0)
                if i[1][11] == 0:
                    stocking_quarter.append('Null')
                else:
                    stocking_quarter.append("Q3")
                continue
            elif (i[1][21] == 100.000):
                money_saved.append(0)
                if i[1][11] == 0:
                    stocking_quarter.append('Null')
                elif i[1][8] == 0:
                    stocking_quarter.append('Q3')
                continue

        v["stocking_quarter"] = [x for x in stocking_quarter]
        v["money_saved"] = [x for x in money_saved]
        return v
    elif c == 4:
        stocking_quarter = []
        money_saved = []
        for i in v.iterrows():
            EOQ = i[1][23]
            if EOQ == None:
                m = i[1][14]*w
                n = int(m)
                if m-n >= 0.5:
                    EOQ = n+1
                elif m-n < 0.5:
                    EOQ = n
            if(i[1][22] <= i[1][21] and i[1][22] != 100):
                stocking_quarter.append("Q4")
                money_saved.append(0)
            elif (i[1][22] > i[1][21] and i[1][22] != 100.000):
                stocking_quarter.append("Q3")
                x = EOQ*i[1][12]
                y = EOQ*i[1][15]
                saved = y-x
                saved = round(saved, 3)
                money_saved.append(saved)
            elif(i[1][22] == 100 and i[1][21] == 100):
                money_saved.append(0)
                if i[1][14] == 0:
                    stocking_quarter.append('Null')
                else:
                    stocking_quarter.append("Q4")
                continue
            elif (i[1][22] == 100.000):
                money_saved.append(0)
                if i[1][14] == 0:
                    stocking_quarter.append('Null')
                elif i[1][11] == 0:
                    stocking_quarter.append('Q3')
                continue

        v["stocking_quarter"] = [x for x in stocking_quarter]
        v["money_saved"] = [x for x in money_saved]
        return v


def Supplier_selection(quarter, f0, Name, Sort, df):
    if quarter == '2':
        stocking_quarter(f0, 0.40, quarter)
        df1 = df[((df['month'] >= 4) & (df["day"] >= 1)) &
                 ((df['month'] <= 6) & (df["day"] <= 30))]
        df2 = df[((df['month'] >= 7) & (df["day"] >= 1)) &
                 ((df['month'] <= 9) & (df["day"] <= 30))]
        ss = df1[['Supplier', 'Item Description', 'Rate']]
        ss_ = df2[['Supplier', 'Item Description', 'Rate']]
        c = 0
        f1 = f0[['Item Description', 'avg_Q1_rate', 'avg_Q2_rate',
                 'stocking_quarter', 'money_saved']]
        f2 = f1[(f1['stocking_quarter'] != 'NULL')]
        item_supplier = {}
        item_supplier_ranked = {}
        for row in f2.iterrows():
            c += 1
            item = row[1][0]
            items_group = ss[ss['Item Description'] == item]
            items_group1 = ss_[ss_['Item Description'] == item]
            stocking_q = row[1][3]
            Q1 = row[1][1]
            Q2 = row[1][2]

            supplier = set()
            if item not in item_supplier.keys():
                item_supplier[item] = '#'
                if stocking_q == 'Q1':
                    for row1 in items_group.iterrows():
                        if row1[1][0] in supplier:
                            continue
                        else:
                            if row1[1][1] == item and row1[1][2] <= Q1:
                                supplier.add(row1[1][0])
                elif stocking_q == 'Q2':
                    for row1 in items_group1.iterrows():
                        if row[1][0] in supplier:
                            continue
                        else:
                            if row1[1][1] == item and row[1][2] <= Q2:
                                supplier.add(row1[1][0])
                else:
                    continue
            # print(c)
            item_supplier[item] = supplier
            ss = ss.drop((ss[ss['Item Description'] == item].index))
            ss_ = ss_.drop((ss_[ss_['Item Description'] == item].index))
    elif quarter == '3':
        stocking_quarter(f0, 0.40, quarter)
        df1 = df[((df['month'] >= 7) & (df["day"] >= 1)) &
                 ((df['month'] <= 9) & (df["day"] <= 30))]
        df2 = df[((df['month'] >= 10) & (df["day"] >= 1)) &
                 ((df['month'] <= 12) & (df["day"] <= 31))]
        ss = df1[['Supplier', 'Item Description', 'Rate']]
        ss_ = df2[['Supplier', 'Item Description', 'Rate']]
        c = 0
        f1 = f0[['Item Description', 'avg_Q2_rate', 'avg_Q3_rate',
                 'stocking_quarter', 'money_saved']]
        f2 = f1[(f1['stocking_quarter'] != 'NULL')]
        item_supplier = {}
        item_supplier_ranked = {}
        for row in f2.iterrows():
            c += 1
            item = row[1][0]
            items_group = ss[ss['Item Description'] == item]
            items_group1 = ss_[ss_['Item Description'] == item]
            stocking_q = row[1][3]
            Q1 = row[1][1]
            Q2 = row[1][2]

            supplier = set()
            if item not in item_supplier.keys():
                item_supplier[item] = '#'
                if stocking_q == 'Q2':
                    for row1 in items_group.iterrows():
                        if row1[1][0] in supplier:
                            continue
                        else:
                            if row1[1][1] == item and row1[1][2] <= Q1:
                                supplier.add(row1[1][0])
                elif stocking_q == 'Q3':
                    for row1 in items_group1.iterrows():
                        if row[1][0] in supplier:
                            continue
                        else:
                            if row1[1][1] == item and row[1][2] <= Q2:
                                supplier.add(row1[1][0])
                else:
                    continue
            # print(c)
            item_supplier[item] = supplier
            ss = ss.drop((ss[ss['Item Description'] == item].index))
            ss_ = ss_.drop((ss_[ss_['Item Description'] == item].index))
    elif quarter == '4':
        stocking_quarter(f0, 0.40, quarter)
        df1 = df[((df['month'] >= 10) & (df["day"] >= 1)) &
                 ((df['month'] <= 12) & (df["day"] <= 31))]
        df2 = df[((df['month'] >= 1) & (df["day"] >= 1)) &
                 ((df['month'] <= 3) & (df["day"] <= 31))]
        ss = df1[['Supplier', 'Item Description', 'Rate']]
        ss_ = df2[['Supplier', 'Item Description', 'Rate']]
        c = 0
        f1 = f0[['Item Description', 'avg_Q3_rate', 'avg_Q4_rate',
                 'stocking_quarter', 'money_saved']]
        f2 = f1[(f1['stocking_quarter'] != 'NULL')]
        item_supplier = {}
        item_supplier_ranked = {}
        for row in f2.iterrows():
            c += 1
            item = row[1][0]
            items_group = ss[ss['Item Description'] == item]
            items_group1 = ss_[ss_['Item Description'] == item]
            stocking_q = row[1][3]
            Q1 = row[1][1]
            Q2 = row[1][2]

            supplier = set()
            if item not in item_supplier.keys():
                item_supplier[item] = '#'
                if stocking_q == 'Q3':
                    for row1 in items_group.iterrows():
                        if row1[1][0] in supplier:
                            continue
                        else:
                            if row1[1][1] == item and row1[1][2] <= Q1:
                                supplier.add(row1[1][0])
                elif stocking_q == 'Q4':
                    for row1 in items_group1.iterrows():
                        if row[1][0] in supplier:
                            continue
                        else:
                            if row1[1][1] == item and row[1][2] <= Q2:
                                supplier.add(row1[1][0])
                else:
                    continue
            # print(c)
            item_supplier[item] = supplier
            ss = ss.drop((ss[ss['Item Description'] == item].index))
            ss_ = ss_.drop((ss_[ss_['Item Description'] == item].index))

    ranked_supplier_list = Supplier_Ranking(
        Buyer=Name, Ranked_list="yes", sortby=Sort, by="Item Quality")
    for key, value in item_supplier.items():
        supplier1 = []
        c = len(supplier1)
        v = len(value)
        for j in ranked_supplier_list:
            if j in value and c != v:
                j = str(c+1) + ". " + j
                supplier1.append(j)
                c += 1
            elif c == v:
                break
            else:
                continue
        item_supplier_ranked[key] = supplier1
    return item_supplier_ranked


def obj_cost_optimization(name, sort, Supplier_Ranking, Q, A):
    Df = pd.read_excel(url)
    df = pd.DataFrame()
    df = Df
    po = pd.read_excel(url_PO)
    df = preprocess(df)
    df = df[["GRN Date", "Supplier", "Buyer", "PO No.", "Item No.", "Item Description", "Challan Qty.",
             "Counted Qty.", "Acp. Qty.", "Acp. UD Qty.", "Rej. Qty.", "Rate", "Amt.", "Currency", "month", "day"]]
    Df = stocking_preprocess(Df, po, name=name)
    tf = choose_analysis(Df, A)
    tf1 = levels(tf)
    if name != "NULL":
        df = df[df['Buyer'].str.contains(name)]
    df1 = df[((df['month'] >= 4) & (df["day"] >= 1)) &
             ((df['month'] <= 6) & (df["day"] <= 30))]
    df2 = df[((df['month'] >= 7) & (df["day"] >= 1)) &
             ((df['month'] <= 9) & (df["day"] <= 30))]
    df3 = df[((df['month'] >= 10) & (df["day"] >= 1)) &
             ((df['month'] <= 12) & (df["day"] <= 31))]
    df4 = df[((df['month'] >= 1) & (df["day"] >= 1)) &
             ((df['month'] <= 3) & (df["day"] <= 31))]
    frequency = frequency_(
        df, Q='full_yr', QC='fullyr_ordered_qty', R='avg_yr_rate')
    frequency1 = frequency_(df1, Q='Q1_apr_jun',
                            QC='Q1_ordered_qty', R='avg_Q1_rate')
    frequency2 = frequency_(df2, Q='Q2_jul_sep',
                            QC='Q2_ordered_qty', R='avg_Q2_rate')
    frequency3 = frequency_(df3, Q='Q3_oct_dec',
                            QC='Q3_ordered_qty', R='avg_Q3_rate')
    frequency4 = frequency_(df4, Q='Q4_jan_mar',
                            QC='Q4_ordered_qty', R='avg_Q4_rate')
    merged_freq = merge(
        [frequency, frequency1, frequency2, frequency3, frequency4])
    freq_final = final_result(merged_freq)
    freq_final = price_fluctuation(freq_final)
    freq_final.sort_values(by="Item Description")
    tf1.sort_values(by="Item Description")
    freq_final["EOQ"] = tf1["EOQ"]
    # freq_final.info()
    null_columns = freq_final.columns[freq_final.isnull().any()]
    c = freq_final[null_columns].isnull().sum()
    # print(c)
    if Q == '1':
        stocking = stocking_quarter(freq_final, 0.30, 1)
        stocking = stocking[["Item Description", "avg_Q1_rate", "avg_Q2_rate", "avg_Q3_rate",
                             "avg_Q4_rate", "Rate shift Q2", "Rate shift Q3", "Rate shift Q4", "EOQ", "stocking_quarter", "max_order_quarter", "money_saved"]]
    elif Q == '2':
        stocking = stocking_quarter(freq_final, 0.30, 2)
        stocking = stocking[["Item Description", "avg_Q1_rate", "avg_Q2_rate",
                             "Rate shift Q2", "EOQ", "stocking_quarter", "money_saved"]]
    elif Q == '3':
        stocking = stocking_quarter(freq_final, 0.30, 3)
        stocking = stocking[["Item Description", "avg_Q2_rate", "avg_Q3_rate",
                             "Rate shift Q2", "Rate shift Q3", "EOQ", "stocking_quarter", "money_saved"]]
    elif Q == '4':
        stocking = stocking_quarter(freq_final, 0.30, 4)
        stocking = stocking[["Item Description", "avg_Q3_rate",
                             "avg_Q4_rate", "Rate shift Q3", "Rate shift Q4", "EOQ", "stocking_quarter", "money_saved"]]
    if Supplier_Ranking == True:
        item_supplier_ranked = Supplier_selection(
            quarter=Q, f0=stocking, Name=name, Sort=sort, df=df)
        # print(item_supplier_ranked)
        gg = pd.DataFrame({"Item Description": [x for x, y in item_supplier_ranked.items()],
                           "supplier": [y for x, y in item_supplier_ranked.items()]
                           })
        gg.sort_values(by="Item Description")
        stocking.sort_values(by="Item Description")
        stocking["supplier"] = gg["supplier"]
        return stocking
    elif Supplier_Ranking == False:
        return stocking

def Allobj_cost_optimization(sort, Supplier_Ranking, Q, A):
    Df = pd.read_excel(url)
    df = pd.DataFrame()
    df = Df
    po = pd.read_excel(url_PO)
    df = preprocess(df)
    df = df[["GRN Date", "Supplier", "Buyer", "PO No.", "Item No.", "Item Description", "Challan Qty.",
             "Counted Qty.", "Acp. Qty.", "Acp. UD Qty.", "Rej. Qty.", "Rate", "Amt.", "Currency", "month", "day"]]
    Df = Allstocking_preprocess(Df, po)
    tf = choose_analysis(Df, A)
    tf1 = levels(tf)
    
    
    df1 = df[((df['month'] >= 4) & (df["day"] >= 1)) &
             ((df['month'] <= 6) & (df["day"] <= 30))]
    df2 = df[((df['month'] >= 7) & (df["day"] >= 1)) &
             ((df['month'] <= 9) & (df["day"] <= 30))]
    df3 = df[((df['month'] >= 10) & (df["day"] >= 1)) &
             ((df['month'] <= 12) & (df["day"] <= 31))]
    df4 = df[((df['month'] >= 1) & (df["day"] >= 1)) &
             ((df['month'] <= 3) & (df["day"] <= 31))]
    frequency = frequency_(
        df, Q='full_yr', QC='fullyr_ordered_qty', R='avg_yr_rate')
    frequency1 = frequency_(df1, Q='Q1_apr_jun',
                            QC='Q1_ordered_qty', R='avg_Q1_rate')
    frequency2 = frequency_(df2, Q='Q2_jul_sep',
                            QC='Q2_ordered_qty', R='avg_Q2_rate')
    frequency3 = frequency_(df3, Q='Q3_oct_dec',
                            QC='Q3_ordered_qty', R='avg_Q3_rate')
    frequency4 = frequency_(df4, Q='Q4_jan_mar',
                            QC='Q4_ordered_qty', R='avg_Q4_rate')
    merged_freq = merge(
        [frequency, frequency1, frequency2, frequency3, frequency4])
    freq_final = final_result(merged_freq)
    freq_final = price_fluctuation(freq_final)
    freq_final.sort_values(by="Item Description")
    tf1.sort_values(by="Item Description")
    freq_final["EOQ"] = tf1["EOQ"]
    # freq_final.info()
    null_columns = freq_final.columns[freq_final.isnull().any()]
    c = freq_final[null_columns].isnull().sum()
    # print(c)
    if Q == '1':
        stocking = stocking_quarter(freq_final, 0.30, 1)
        stocking = stocking[["Item Description", "avg_Q1_rate", "avg_Q2_rate", "avg_Q3_rate",
                             "avg_Q4_rate", "Rate shift Q2", "Rate shift Q3", "Rate shift Q4", "EOQ", "stocking_quarter", "max_order_quarter", "money_saved"]]
    elif Q == '2':
        stocking = stocking_quarter(freq_final, 0.30, 2)
        stocking = stocking[["Item Description", "avg_Q1_rate", "avg_Q2_rate",
                             "Rate shift Q2", "EOQ", "stocking_quarter", "money_saved"]]
    elif Q == '3':
        stocking = stocking_quarter(freq_final, 0.30, 3)
        stocking = stocking[["Item Description", "avg_Q2_rate", "avg_Q3_rate",
                             "Rate shift Q2", "Rate shift Q3", "EOQ", "stocking_quarter", "money_saved"]]
    elif Q == '4':
        stocking = stocking_quarter(freq_final, 0.30, 4)
        stocking = stocking[["Item Description", "avg_Q3_rate",
                             "avg_Q4_rate", "Rate shift Q3", "Rate shift Q4", "EOQ", "stocking_quarter", "money_saved"]]
    if Supplier_Ranking == True:
        item_supplier_ranked = Supplier_selection(
            quarter=Q, f0=stocking, Sort=sort, df=df)
        # print(item_supplier_ranked)
        gg = pd.DataFrame({"Item Description": [x for x, y in item_supplier_ranked.items()],
                           "supplier": [y for x, y in item_supplier_ranked.items()]
                           })
        gg.sort_values(by="Item Description")
        stocking.sort_values(by="Item Description")
        stocking["supplier"] = gg["supplier"]
        return stocking
    elif Supplier_Ranking == False:
        return stocking

def levels(dataframe):
    # dataframe.info()
    dataframe = safetystocklevel(dataframe, 5)
    # print("this is safetystock",safetystock)
    # reorder level
    dataframe['Reorder level'] = dataframe['Average daily consumption'] * \
        dataframe['Average Lead Time'] + dataframe["safetystock"]

    # optimal reorder quantity
    dataframe["EOQ"] = np.sqrt(
        2*(dataframe['Avg_total_consumption']*dataframe['Avg Rate'])/dataframe['Avg Carrying cost'])

    # minimum stock level
    dataframe["Minimum Stock level"] = dataframe["Reorder level"] - \
        (dataframe["Average daily consumption"]*dataframe["Average Lead Time"])

    # maximum stock level
    dataframe["Maximum Stock level"] = dataframe["Reorder level"] + dataframe["EOQ"] - \
        (dataframe["Minimum daily consumption"]*dataframe["Minimum Lead Time"])

    dataframe = dataframe[["Item Description", "Maximum Stock level",
                           "Reorder level", "EOQ", "Minimum Stock level", "Class"]]
    return dataframe


def obj_stocking(A, name):
    Dataframe = pd.read_excel(url)
    PO = pd.read_excel(url_PO)
    Dataframe = stocking_preprocess(Dataframe, PO, name)
    df = choose_analysis(Dataframe, A)
    df = levels(df)
    return df

def Allobj_stocking(A):
    Dataframe = pd.read_excel(url)
    PO = pd.read_excel(url_PO)
    Dataframe = Allstocking_preprocess(Dataframe, PO)
    df = choose_analysis(Dataframe, A)
    df = levels(df)
    return df

def pareto_prepoc(i):
    i.rename(columns={'Item No. & Desc.': 'Items'}, inplace=True)
    item = i['Items'].str.split('--', n=1, expand=True)
    i['Item No.'] = item[0]
    i['Item Description'] = item[1]
    i.drop(['Items'], axis=1)
    i.loc[(i['Item Description'] == 'HOTEL BILL') | (i['Item Description'] == '-LABOR CHARGE') | (i['Item Description'] == 'TRANSPORTATION CHARGE') | (i['Item Description'] =='TRANSPORTATION CHARGES') | (i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS') | (i['Item Description'] == 'LUNCH BILL')]
    i = i.drop((i[i['Item Description'] == 'HOTEL BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGE'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGES'].index) | (i[i['Item Description'] == '-LABOR CHARGE'].index) | (i[i['Item Description'] == 'LUNCH BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS'].index))
    j = i
    return j

def pareto_pre():
    dataframe = pd.read_excel(url)
    dataframe1 = pareto_prepoc(dataframe)
    dataframe1["Amount"] = dataframe1["Amt."]
    dataframe1["Total Rate"] = dataframe1["Rate"]
    newdf = dataframe1[["Item Description","Supplier", "Acp. Qty.", "Total Rate", "Amount", "Currency"]]
    nw = newdf.sort_values(by=['Amount'],ascending=False)
    nw1= nw.groupby(['Item Description','Currency'], as_index=False)['Amount'].sum()
    nw1=nw1.sort_values(by=['Amount'],ascending=False)
    totalamt = nw1["Amount"].sum()
    amt70 = (totalamt * 70) / 100
    add=0
    count=0
    for i in nw1.index:
        add=add + nw1["Amount"][i]
        if add <=amt70:
            count=count+1
            continue
        else:
            break
    df1 = pd.DataFrame(columns = ['Item Description', 'Amount', 'Currency'])
    for i in nw1.index:
        if count>=0:
            per=(nw1['Amount'][i]*100)/amt70
            df1=df1.append({'Item Description' : nw1["Item Description"][i], 'Amount' :nw1["Amount"][i],'Currency' : nw1["Currency"][i],'Percentage' : per},ignore_index = True)
        count=count-1
    df3 = pd.DataFrame(columns = ['Item Description', 'Amount', 'Currency','Percentage','cumulative'])
    cadd=0
    c=0
    for i in range(0,len(df1.index)):
        cadd=cadd+df1['Percentage'][i]
        df3=df3.append({'Item Description' : df1["Item Description"][i], 'Amount' :df1["Amount"][i],'Currency' : df1["Currency"][i],'Percentage' : df1["Percentage"][i],'cumulative' : cadd},ignore_index = True)
    df3.insert(5,"Total Percentage",0.0000000001,True)
     # df3.round({"Total Percentage":3})

    for i in range(0,len(df1.index)):
        currentValue = (df3['Amount'][i]*100)/totalamt
        df3['Total Percentage'][i] = currentValue
        # df3 = df3.append({'Total Percentage IS':currentValue},ignore_index = True)
    print(totalamt)
    df3['Total Percentage'].round(decimals = 3)


    return(df3)

# def supplierdata_prepocessor(df):
#     df["Amount"] = df["Amt."]
#     df["Total Rate"] = df["Rate"]
#     newdf = df[["Supplier","Item No. & Desc.","Acp. Qty.", "Total Rate", "Amount", "Currency"]]
#     i=newdf
#     i.rename(columns={'Item No. & Desc.': 'Items'}, inplace=True)
#     item = i['Items'].str.split('--', n=1, expand=True)
#     i['Item No.'] = item[0]
#     i['Item Description'] = item[1]
#     i.drop(['Items'], axis=1)
#     i.loc[(i['Item Description'] == 'HOTEL BILL') | (i['Item Description'] == '-LABOR CHARGE') | (i['Item Description'] == 'TRANSPORTATION CHARGE') | (i['Item Description'] =='TRANSPORTATION CHARGES') | (i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS') | (i['Item Description'] == 'LUNCH BILL')]
#     i = i.drop((i[i['Item Description'] == 'HOTEL BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGE'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGES'].index) | (i[i['Item Description'] == '-LABOR CHARGE'].index) | (i[i['Item Description'] == 'LUNCH BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS'].index))
#     newdf = i[["Supplier","Item Description","Acp. Qty.", "Total Rate", "Amount", "Currency"]]
#     return newdf

def supplierdata_prepocessor(df):
    df["Amount"] = df["Amt."]
    df["Total Rate"] = df["Rate"]
    newdf = df[["Supplier","Item No. & Desc.","Acp. Qty.", "Total Rate", "Amount", "Currency"]]
    i=newdf
    i.rename(columns={'Item No. & Desc.': 'Items'}, inplace=True)
    item = i['Items'].str.split('--', n=1, expand=True)
    i['Item No.'] = item[0]
    i['Item Description'] = item[1]
    i.drop(['Items'], axis=1)
    i.loc[(i['Item Description'] == 'HOTEL BILL') | (i['Item Description'] == '-LABOR CHARGE') | (i['Item Description'] == 'TRANSPORTATION CHARGE') | (i['Item Description'] =='TRANSPORTATION CHARGES') | (i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS') | (i['Item Description'] == 'LUNCH BILL')]
    i = i.drop((i[i['Item Description'] == 'HOTEL BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGE'].index) | (i[i['Item Description'] == 'TRANSPORTATION CHARGES'].index) | (i[i['Item Description'] == '-LABOR CHARGE'].index) | (i[i['Item Description'] == 'LUNCH BILL'].index) | (i[i['Item Description'] == 'TRANSPORTATION OF MACHINES & GOODS & OTHER MISCELLANEOUS'].index))
    ## changes
    Desc = i[["Supplier","Item Description","Acp. Qty.", "Total Rate", "Amount", "Currency","Item No."]]
    Desc = Desc.sort_values(by=['Acp. Qty.'],ascending=True)
    #Desc.drop(Desc[Desc['Item Description'] != '-'].index, inplace = True)
#     for index_label, row_series in Desc.iterrows():
#    # For each row update the 'Bonus' value to it's double
#         Desc.at[index_label , 'Item Description'] = row_series['Item No.']
#     return Desc
    for index_label, row_series in Desc.iterrows():
   # For each row update the 'Bonus' value to it's double
        if  Desc.at[index_label , 'Item Description'] == '-':
            Desc.at[index_label , 'Item Description'] = row_series['Item No.']
        else:
            continue
    Desc = Desc.sort_values(by=['Amount'],ascending=True)
    return Desc



def allprocessingdata(df):
    nw1= df.groupby(['Item Description'], as_index=False)['Amount'].sum()
    nw2= df.groupby(['Item Description'], as_index=False)['Acp. Qty.'].sum()
    nw3= df.groupby(['Item Description'], as_index=False)['Total Rate'].mean()
    pd.options.display.float_format = "{:.2f}".format
    result=nw2.merge(nw3[['Item Description', 'Total Rate']])
    result=result.merge(nw1[['Item Description', 'Amount']])
    result = result.rename({'Total Rate': 'Avg.Rate'}, axis=1)
    # result = result.sort_values(by=['Amount'],ascending=False)
    
    return result

def processingdata(df,name):
    df1 = pd.DataFrame(columns = ['Item Description', 'Supplier','Acp. Qty.','Total Rate', 'Amount', 'Currency'])
    for i in df.index:
        if df['Supplier'][i] == name:
            df1=df1.append({'Item Description' : df["Item Description"][i], 'Supplier' :df["Supplier"][i],'Acp. Qty.' : df["Acp. Qty."][i],'Total Rate' : df["Total Rate"][i], 'Amount' :df["Amount"][i],'Currency' : df["Currency"][i]},ignore_index = True)
    nw1= df1.groupby(['Item Description', 'Supplier'], as_index=False)['Amount'].sum()
    nw2= df1.groupby(['Item Description', 'Supplier'], as_index=False)['Acp. Qty.'].sum()
    nw3= df1.groupby(['Item Description', 'Supplier'], as_index=False)['Total Rate'].mean()
    pd.options.display.float_format = "{:.2f}".format
    result=nw2.merge(nw3[['Item Description', 'Supplier', 'Total Rate']])
    result=result.merge(nw1[['Item Description', 'Supplier', 'Amount']])
    result = result.rename({'Total Rate': 'Avg.Rate'}, axis=1)
    result1=result.drop(['Supplier',], axis=1)
    # result1 = result1.sort_values(by=['Amount'],ascending=False)
    return result1

@app.route("/pareto/plot.png")
@login_required
def plot_png():

    df = CSV
    fig = Figure()
    axis = fig.add_subplot(1,1,1)
    # axis.set_title("Pareto Graph for first 20 items")
    axis.set_xlabel("Item Description",color="black",fontsize=14)
    axis.set_ylabel("Amount",color="black",fontsize=14)

    name = df["Item Description"].head(20)
    perc = df["Amount"].head(20)
    x_value = range(1, len(perc)+1)
    axis.set_xticklabels(x_value, rotation=0)
    colors = np.random.rand(len(perc),3)
    axis.set_facecolor('#ffe9f7')
    axis.grid(True)
    axis.bar(name, perc,color = colors,edgecolor = "black" )

    axis2=axis.twinx()
    name1 = df["Item Description"].head(20)
    price1 = df["cumulative"].head(20)
    x_value = range(1, len(price1)+1)
    axis2.set_xticklabels(x_value, rotation=0)
    axis2.plot(name1,price1,color='black',marker='*')
    axis2.set_ylabel("Cumulative Percentage",color="black",fontsize=14)

    output = io.BytesIO()
    FigureCanvas(fig).print_png(output)
    return Response(output.getvalue(), mimetype="image/png")

# @app.route("/pareto/plot1.png")
# @login_required
# def plot_png1():
#     df = CSV
#     fig = Figure(figsize=(12,5))
#     axis = fig.add_subplot(1,1,1)
#     axis.set_title("Pareto Graph for first 50 items")
#     axis.set_xlabel("Item Description",color="black",fontsize=14)
#     axis.set_ylabel("Amount( * 10^7)",color="black",fontsize=14)

#     name = df["Item Description"].head(50)
#     perc = df["Amount"].head(50)
#     x_value = range(1, len(perc)+1)
#     axis.set_xticklabels(x_value, rotation=0)
#     colors = np.random.rand(len(perc),3)
#     axis.set_facecolor('#ffe9f7')
#     axis.grid(True)
#     axis.bar(name, perc,color = colors,edgecolor = "black" )

#     axis2=axis.twinx()
#     name1 = df["Item Description"].head(50)
#     price1 = df["cumulative"].head(50)
#     x_value = range(1, len(price1)+1)
#     axis2.set_xticklabels(x_value, rotation=0)
#     axis2.plot(name1,price1,color='black',marker='*')
#     axis2.set_ylabel("Cumulative Percentage",color="black",fontsize=14)

#     output = io.BytesIO()
#     FigureCanvas(fig).print_png(output)
#     return Response(output.getvalue(), mimetype="image/png")

@app.route("/pareto/plot1.png")
@login_required
def plot_png1():
    df = CSV
    fig = Figure(figsize=(12,5))
    # changes
    fig.set_size_inches(22.5, 11.5, forward=True)
    axis = fig.add_subplot(1,1,1)
    # axis.set_title("Pareto Graph for first 50 items")
    axis.set_xlabel("Item Description",color="black",fontsize=45)
    axis.set_ylabel("Amount( * 10^7)",color="black",fontsize=45)

    name = df["Item Description"].head(50)
    perc = df["Amount"].head(50)
    x_value = range(1, len(perc)+1)
    axis.set_xticklabels(x_value, rotation=0)
    colors = np.random.rand(len(perc),3)
    axis.set_facecolor('#ffe9f7')
    axis.grid(True)
    axis.bar(name, perc,color = colors,edgecolor = "black" )

    axis2=axis.twinx()
    name1 = df["Item Description"].head(50)
    price1 = df["cumulative"].head(50)
    x_value = range(1, len(price1)+1)
    axis2.set_xticklabels(x_value, rotation=0)
    axis2.plot(name1,price1,color='black',marker='*')
    axis2.set_ylabel("Cumulative Percentage",color="black",fontsize=45)

    output = io.BytesIO()
    FigureCanvas(fig).print_png(output)
    return Response(output.getvalue(), mimetype="image/png")

@app.route("/pareto/plot2.png")
@login_required
def plot_png2():
    df = CSV
    fig = Figure(figsize=(15,8))
    fig.set_size_inches(50.5, 11.5, forward=True)
    axis = fig.add_subplot(1,1,1)
    # axis.set_title("Pareto Graph for first 80 items")
    axis.set_xlabel("Item Description",color="black",fontsize=45)
    axis.set_ylabel("Amount( * 10^7)",color="black",fontsize=45)
    axis.xaxis.label. set_size(60)
    name = df["Item Description"].head(80)
    perc = df["Amount"].head(80)
    x_value = range(1, len(perc)+1)
    axis.set_xticklabels(x_value, rotation=0)
    colors = np.random.rand(len(perc),3)
    axis.set_facecolor('#ffe9f7')
    axis.grid(True)
 
    axis.bar(name, perc,color = colors,edgecolor = "black" )

    axis2=axis.twinx()
    name1 = df["Item Description"].head(80)
    price1 = df["cumulative"].head(80)
    x_value = range(1, len(price1)+1)
    axis2.set_xticklabels(x_value, rotation=0)
    axis2.plot(name1,price1,color='black',marker='*')
    axis2.grid(True)
    axis2.set_ylabel("Cumulative Percentage",color="black",fontsize=45)

    output = io.BytesIO()
   
    FigureCanvas(fig).print_png(output)
    return Response(output.getvalue(), mimetype="image/png")

@app.route("/consumption", methods=['GET', 'POST'])
@login_required
def consumption():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        df = pd.read_excel(url)
        df = consumption_preprocess(df)
        form.buyer.choices = buyers_list
        if form.validate_on_submit() or request.method == 'POST':
            name = form.buyer.data
            if name == 'All Buyers':
                val = form.duration.data
                if form.duration.data == '5':
                    df = avgconsumption(df, val)
                    CSV = df
                    df.index += 1
                    return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Consumption', titles=df.columns.values, form=form)
                else:
                    df = quarterlyframe(df, val)
                    CSV = df
                    df.index += 1
                    return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Consumption', titles=df.columns.values, form=form)
            else:
                val = form.duration.data
                if form.duration.data == '5':
                    df = avgconsumption1(df, val, name)
                    CSV = df
                    df.index += 1
                    return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Consumption', titles=df.columns.values, form=form)
                else:
                    df = quarterlyframe1(df, val, name)
                    CSV = df
                    df.index += 1
                    return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Consumption', titles=df.columns.values, form=form)
        return render_template("consumption.html", form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')
    
@app.route("/duplicate", methods=['GET', 'POST'])
@login_required
def duplicate():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        df = pd.read_excel(url)
        df = duplicate_preprocess(df)
        form.buyer.choices = buyers_list
        if request.method == 'POST':
            name = form.buyer.data
            df = duplicate_code(df, name)
            CSV = df
            df.index = np.arange(0,len(df))
            df.index += 1
            print("CSV assigned")
            return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Duplicate item codes', titles=df.columns.values, form=form)
        return render_template("duplicates.html", form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')
    


@app.route("/frequency", methods=['GET', 'POST'])
@login_required
def frequency():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        form.buyer.choices = buyers_list
        if request.method == 'POST':
            name = form.buyer.data
            if name == 'All Buyers':
                sort = form.supplier_sort.data
                df = obj_frequency1(sort)
                CSV = df
                df.reset_index(inplace = True, drop = True)
                df.index += 1
                flash(f'"Item Description": Item Description given in the GRN Report <br> "Qx_ordered_qty": Qty. ordered in that quarter/year <br> "max_freq_count" : The maximum frequency of order <br> "max_freq_quarter" : Quarter with maximum order frequency <br> "max_order_qty" : Maximum ordered qty. in a quarter <br> "max_order_quarter" : The quarter in which maximum qty. was ordered', 'secondary')
                return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Order frequency', titles=df.columns.values, form=form)
            else:
                sort = form.supplier_sort.data
                df = obj_frequency(name, sort)
                CSV = df
                df.reset_index(inplace = True, drop = True)
                df.index += 1
                flash(f'"Item Description": Item Description given in the GRN Report <br> "Qx_ordered_qty": Qty. ordered in that quarter/year <br> "max_freq_count" : The maximum frequency of order <br> "max_freq_quarter" : Quarter with maximum order frequency <br> "max_order_qty" : Maximum ordered qty. in a quarter <br> "max_order_quarter" : The quarter in which maximum qty. was ordered', 'secondary')
                return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Order frequency', titles=df.columns.values, form=form)
        return render_template("frequency.html", form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')
    
@app.route("/supplier", methods=['GET', 'POST'])
@login_required
def supplier():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        form.buyer.choices = buyers_list
        if request.method == 'POST':
            Buyer = form.buyer.data
            by = form.supplier.data
            sortby = form.supplier_sort.data
            if Buyer == 'All Buyers':
                df = AllSupplier_Ranking("no", by, sortby)
                CSV = df
                df.reset_index(inplace = True, drop = True)
                df.index += 1
                if by == "Item Quality":
                    pd.options.display.float_format = "{:.2f}".format
                    # return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], titles=df.columns.values)
                    return render_template("colortable.html", the_title='Suplier Ranking by Item Quality', column_names=df.columns.values, row_data=list(df.values.tolist()), zip=zip)
                elif by == "In-time Delivery":
                    pd.options.display.float_format = "{:.2f}".format
                    return render_template("colortable-in.html", the_title='Suplier Ranking by In-time Delivery', column_names=df.columns.values, row_data=list(df.values.tolist()), zip=zip)
            else:
                df = Supplier_Ranking(Buyer, "no", by, sortby)
                CSV = df
                df.reset_index(inplace = True, drop = True)
                df.index += 1
                if by == "Item Quality":
                    pd.options.display.float_format = "{:.2f}".format
                    # return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], titles=df.columns.values)
                    return render_template("colortable.html", the_title='Suplier Ranking by Item Quality', column_names=df.columns.values, row_data=list(df.values.tolist()), zip=zip)
                elif by == "In-time Delivery":
                    pd.options.display.float_format = "{:.2f}".format
                    return render_template("colortable-in.html", the_title='Suplier Ranking by In-time Delivery', column_names=df.columns.values, row_data=list(df.values.tolist()), zip=zip)
        return render_template("supplier.html", form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')
    


@app.route("/cost", methods=['GET', 'POST'])
@login_required
def cost():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        form.buyer.choices = buyers_list
        if request.method == 'POST':
            name = form.buyer.data
            if name == 'All Buyers':
                sort = form.supplier_sort.data
                quarter = form.cost_duration.data
                A = form.analysis.data
                Supplier_Ranking = form.checkme.data
                df = Allobj_cost_optimization(sort, Supplier_Ranking, quarter, A)
                df=df.drop(['money_saved'], axis=1)
                CSV = df
                df.index += 1
                return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Cost Optimization', titles=df.columns.values, form=form)
            else:
                sort = form.supplier_sort.data
                quarter = form.cost_duration.data
                A = form.analysis.data
                Supplier_Ranking = form.checkme.data
                df = obj_cost_optimization(name, sort, Supplier_Ranking, quarter, A)
                df=df.drop(['money_saved'], axis=1)
                CSV = df
                df.index += 1
            return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Cost Optimization', titles=df.columns.values, form=form)
        return render_template("cost.html", form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')


@app.route("/inventory", methods=['GET', 'POST'])
@login_required
def inventory():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        form.buyer.choices = buyers_list
        if request.method == 'POST':
            name = form.buyer.data
            if name == 'All Buyers':
                val = form.analysis.data
                df = Allobj_stocking(val)
                CSV = df
                df.index += 1
                return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Stocking Policy', titles=df.columns.values, form=form)
            else:
                val = form.analysis.data
                df = obj_stocking(val, name)
                CSV = df
                df.index += 1
                return render_template('table.html',  tables=[df.to_html(classes='data', header="true")], the_title='Stocking Policy', titles=df.columns.values, form=form)
        return render_template("inventory.html", form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')
    
@app.route("/pareto",methods=['GET','POST'])
@login_required
def pareto():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        df = pareto_pre()
        if True:
            CSV = df
            df.index +=1
            pd.set_option('display.max_colwidth', 800)
            return render_template('tablepareto.html',  tables=[df.to_html(classes='data', header="true")], titles=df.columns.values)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')

@app.route("/supplierdata",methods=["GET", "POST"])
def supplierdata():
    if True:
        global CSV
        pd.options.display.float_format = "{:.2f}".format
        form = DataframeForm(request.form)
        df = pd.read_excel(url)
        df = supplierdata_prepocessor(df)

        form.supplierwisedata.choices = supplier_list
        if form.validate_on_submit() or request.method == 'POST':
            name = form.supplierwisedata.data
            if name == 'All Suppliers':
                df = allprocessingdata(df)
                df.index +=1
                return render_template('supplierdata.html',the_title='Supplier Data',  tables=[df.to_html(classes='data', header="true")], titles=df.columns.values, form=form)
            else:
                df = processingdata(df, name)
                df.index +=1
            return render_template('supplierdata.html',the_title='Supplier Data',  tables=[df.to_html(classes='data', header="true")], titles=df.columns.values, form=form)
    # df = allprocessingdata(df)
    # df.index +=1
        return render_template('supplierdata.html', form=form)
    else:
        flash(f'Please Upload files!', 'warning')
        return redirect('/home')
    

@app.route("/downloads", methods=["GET", "POST"])
@login_required
def downloads():
    global CSV
    if request.method == "GET":
        CSV = build_csv_data(CSV)

        try:
            return send_from_directory(app.config["get_result"], path="results.xlsx", as_attachment=True)
        except FileNotFoundError:
            abort(404)



if __name__ == '__main__':
    app.run(debug=True)