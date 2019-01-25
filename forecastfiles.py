#!/usr/bin/env python3

import os
import sys
import pyodbc
import datetime
import smtplib
import codecs

import numpy as np
import pandas as pd

from pandas import ExcelWriter
from openpyxl import load_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText

####################
# Global Variables #
####################

#Last Month's Date - Script gets called at the beginning of each month to generate the previous month's price files
#The files will get generated with data from 3 business days ago, to correspond to the delay for updating the ICE table
today = datetime.datetime.today()
#today=datetime.datetime(2018,8,13, datetime.datetime.now().time().hour,datetime.datetime.now().time().minute,datetime.datetime.now().time().second)
#today = datetime.datetime(int(sys.argv[1]), int(sys.argv[2]), int(sys.argv[3]), datetime.datetime.now().hour, datetime.datetime.now().minute, 
#lastmonth = today - datetime.timedelta(days=today.day)
curr_time = str(datetime.datetime.now().time().hour)+str(datetime.datetime.now().time().minute)+str(datetime.datetime.now().time().second)
lastmonth = pd.to_datetime(str(np.busday_offset(today, -3, roll = 'backward'))+curr_time, format='%Y-%m-%d%H%M%S')

#Open settings-windows.txt to define settings, directories, etc for updating the databases
settings={}
with open('/Users/developer/Desktop/settings-windows.txt') as f:
	for line in f:
		if '#' not in line and len(line)>1:
			(key,val)=line.strip().split(':')
			settings[key]=val

def main():
	
	#Generate file on First Friday of the Month
	#if today.day < 8 and today.weekday() == 4:
	#if today.weekday() < 4:
	
	try:
		#############
		#Gas Futures#
		#############
		query = "SELECT [Gas Futures].[Sale Date],[Gas Futures].[Delivery Date],[Gas Futures].[Settlement] FROM [Gas Futures] WHERE [Gas Futures].[Sale Date] LIKE '{}';".format(lastmonth.strftime('%m')+'%'+lastmonth.strftime('%Y'))
		cnxn = pyodbc.connect("DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={}".format(settings['database-file']))
		cursor = cnxn.cursor()

		#Save query to dataframe
		df = pd.read_sql_query(query, cnxn)
		
		#Convert to floating point, and sort by date
		df.loc[:,'Sale Date'] = pd.to_datetime(df.loc[:,'Sale Date'])
		df.loc[:,'Delivery Date'] = pd.to_datetime(df.loc[:,'Delivery Date'])
		df.loc[:,'Settlement']=df.loc[:,'Settlement'].astype(float)
		df.loc[:,'Settlement']=df.loc[:,'Settlement'].astype(float)
		df=df.sort_values(['Sale Date', 'Delivery Date']).reset_index(drop=True)
		df.loc[:,'Sale Date'] = [date.strftime('%m/%d/%Y') for date in df['Sale Date']]
		df.loc[:,'Delivery Date'] = [date.strftime('%m/%d/%Y') for date in df['Delivery Date']]

		#Save to Excel	
		writer = ExcelWriter(settings['forecast-directory']+'Gas Futures {}.xlsx'.format(lastmonth.strftime('%B %Y')))
		df.to_excel(writer,'Sheet1',startrow=0, startcol=0, index=False, float_format='%.3f')
		writer.save()
		writer.close()
			
		#Close database connection
		cnxn.commit()
		cnxn.close()
		
		#Log on success
		logs(settings, 'Database', 'Gas Futures')
		
	except Exception as e:
		errmsg = errmsg = 'Error on line {} of {}: {}'.format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[-1].tb_frame.f_code.co_filename,e)
		mail(settings, 'Gas Futures Forecast File Error', logs(settings, 'Error', 'Gas Futures Forecast File Error', e))
	
	try:
		####################
		#Basis Swap Futures#
		####################
		query = "SELECT [ICE].[Sale Date],[ICE].[Contract Month],[ICE].[Commodity Name], [ICE].[Settle Price] FROM [ICE] WHERE [ICE].[Sale Date] LIKE '{}' AND ([ICE].[Commodity Name] LIKE 'SCB%' OR [ICE].[Commodity Name] LIKE 'PGE%' OR [ICE].[Commodity Name] LIKE 'SCL%');".format(lastmonth.strftime('%m')+'%'+lastmonth.strftime('%Y'))
		#query = "SELECT [Basis Swaps].[Sales Date],[Basis Swaps].[Futures Date],[Basis Swaps].[PG&E Citygate], [Basis Swaps].[SoCal] FROM [Basis Swaps] WHERE [Basis Swaps].[Sales Date] LIKE '{}';".format(lastmonth.strftime('%m')+'%'+lastmonth.strftime('%Y'))
		cnxn = pyodbc.connect("DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={}".format(settings['database-file']))
		cursor = cnxn.cursor()

		#Save query to dataframe
		df = pd.read_sql_query(query, cnxn)
		
		#Rename columns
		df.rename(columns={'Sale Date':'Sales Date', 'Contract Month':'Futures Date'}, inplace=True)
		#Convert to datetime
		df.loc[:,'Sales Date'] = pd.to_datetime(df.loc[:,'Sales Date'])
		df.loc[:,'Futures Date'] = pd.to_datetime(df.loc[:,'Futures Date'])
		#Convert to floating point
		df.iloc[:,-1] = df.iloc[:,-1].astype(float)
		#df.loc[:,'PG&E Citygate']=df.loc[:,'PG&E Citygate'].replace('',np.nan).astype(float).fillna('')
		#df.loc[:,'SoCal']=df.loc[:,'SoCal'].replace('',np.nan).astype(float).fillna('')
		#Pivot the data so the commodities are each in their own row
		df = df.pivot_table(index=['Sales Date', 'Futures Date'], columns=['Commodity Name']).reset_index().fillna('')
		#Merge the MultiIndex columns
		df.columns = [' '.join(col).strip() for col in df.columns.values]
		#Sort the dataframe
		df = df.sort_values(['Sales Date', 'Futures Date'], ascending=[False, False])
		df.loc[:,'Sales Date'] = [date.strftime('%m/%d/%Y') for date in df['Sales Date']]
		df.loc[:, 'Futures Date'] = [date.strftime('%m/%d/%Y') for date in df['Futures Date']]
		#msk = (df['PG&E Citygate'] != '') | (df['SoCal'] != '')
		#df = df[msk]
		
		#Save to Excel	
		writer = ExcelWriter(settings['forecast-directory']+'Basis Swaps Futures {}.xlsx'.format(lastmonth.strftime('%B %Y')))
		df.to_excel(writer,'Sheet1',startrow=0, startcol=0, index=False, float_format='%.3f')
		writer.save()
		writer.close()
			
		#Close database connection
		cnxn.commit()
		cnxn.close()
		
		#Log on success
		logs(settings, 'Database', 'Basis Swaps Futures',spacers=True)
	
	except Exception as e:
		errmsg = errmsg = 'Error on line {} of {}: {}'.format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[-1].tb_frame.f_code.co_filename,e)
		mail(settings, 'Basis Swap Futures Forecast File Error', logs(settings, 'Error', 'Basis Swap Futures Forecast File Error', e))
		
#############
# Functions #
#############
		
def mail(inputs, subject, text):
	#Function to send email alerts
	gmail_sender = inputs['email-sender']
	gmail_pwd = inputs['email-pwd']
	msg = MIMEMultipart()
	msg['From'] = gmail_sender
	msg['To'] = inputs['email-recipient']
	msg['Subject'] = subject
	msg.attach(MIMEText(text))
	mailServer = smtplib.SMTP("smtp.gmail.com", 587)
	mailServer.ehlo()
	mailServer.starttls()
	mailServer.ehlo()
	mailServer.login(gmail_sender, gmail_pwd)
	mailServer.sendmail(gmail_sender, inputs['email-recipient'], msg.as_string())
	mailServer.close()

def logs(inputs, logname, logsource, error=None, spacers=False):
	#Check if the log files exist. If they don't, create them.
	direc = inputs['log-directory']
	err = inputs['log-directory']+inputs['error-log']
	db = inputs['log-directory']+inputs['database-log']
	if not os.path.exists(direc):
		os.makedirs(direc)
	if not os.path.isfile(db):
		open(log,'w').close()
	if not os.path.isfile(err):
		open(err,'w').close()
	#Error Log
	if logname == 'Error':
		f = inputs['log-directory']+inputs['error-log']
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		#Check if error message is a user-defined string, or Python error type
		if isinstance(error, str):
			#Save the error message
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} on {}'.format(error, today.strftime('%a %d %b %y at %I:%M %p')+'\n\n' + data))
		else:
			#Format and save the error message
			errmsg = 'Error on line {} of {}: {}'.format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[-1].tb_frame.f_code.co_filename,error)
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} Error on {}: '.format(logsource, today.strftime('%a %d %b %y at %I:%M %p'))+errmsg+'\n\n' + data)
			return errmsg
	#Database Log
	if logname == 'Database':
		f = inputs['log-directory']+inputs['database-log']
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		if error != None:
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} on {}\n'.format(error, today.strftime('%a %d %b %Y at %I:%M %p')) + data)
		else:	
			#Save database related messages to log
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} file was successfully generated on {}\n'.format(logsource, today.strftime('%a %d %b %Y at %I:%M %p')) + data)
		if spacers==True:
			with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{:*<50} {}'.format('','\n') + data)

	
if __name__ == "__main__": main()