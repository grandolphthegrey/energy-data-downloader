#!/usr/bin/env python3

import datetime
import pyodbc
import smtplib
import re
import sys
import os
import codecs
import glob

import numpy as np
import pandas as pd
import email.encoders as Encoders

from datetime import timedelta
from openpyxl import load_workbook
from pandas import ExcelWriter
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText


####################
# Global Variables #
####################

#Today's date
today = datetime.datetime.today()
#today=datetime.datetime(2019,1,3, datetime.datetime.now().time().hour,datetime.datetime.now().time().minute,datetime.datetime.now().time().second)
#today = datetime.datetime(int(sys.argv[1]), int(sys.argv[2]), int(sys.argv[3]), datetime.datetime.now().hour, datetime.datetime.now().minute, datetime.datetime.now().second)

#Next business day
next_day = pd.to_datetime(str(np.busday_offset(today, 1, roll='backward')))
#Yesterday's date
yesterday = pd.to_datetime(str(np.busday_offset(today, -1, roll = 'backward')))
file_date = yesterday.strftime('%m')+'-'+yesterday.strftime('%d')
#The day before yesterday
two_days = pd.to_datetime(str(np.busday_offset(today, -2, roll='backward')))


################
# Main Program #
################

def main():
	try:
		#Open settings-windows.txt to define settings, directories, etc for updating the databases
		settings={}
		with open('/Users/developer/Desktop/settings-windows.txt') as f:
			for line in f:
				if '#' not in line and len(line)>1:
					(key,val)=line.strip().split(':')
					settings[key]=val
		
		#Every table except CAISO gets updated only only weekdays -- Monday is 0 and Sunday is 6
		
		if today.weekday() < 5:
			
			#Update the Gas Futures table
			gas_futures(settings)
			
			#Update the Basis Swaps table
			basis_swaps(settings)

			#Update the Basis Swaps Nodal table
			nodal_prices(settings)

			#Update the Gas Daily table
			gas_prices(settings)
			
			#Update the Megawatt Daily table
			megawatt_daily(settings)
			
			#Update the ICE table with data from three business days ago
			ice(settings, icedate=pd.to_datetime(str(np.busday_offset(today, -3, roll = 'backward'))))
			
			#Update the nodal gas database
			nodal_exchange(settings)
		
		#Update the CAISO table
		caiso(settings)
		
		#If today is Friday, then update the ICE table with last week's data
		#if today.weekday() == 4:
		#	for daystep in np.arange(-9,-4,1):
		#		icedate = pd.to_datetime(str(np.busday_offset(today, daystep, roll = 'backward')))
		#		ice(settings,icedate)
		
	except Exception as e:
		#Note and log errors, if any
		mail(settings, 'Main Database Update Program Error', logs(settings, 'Error', 'Main Database Update Program', e))

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
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} on {}'.format(error, today.strftime('%a %d %b %Y at %I:%M %p')+'\n\n' + data))
		else:
			#Format and save the error message
			errmsg = 'Error on line {} of {}: {}'.format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[-1].tb_frame.f_code.co_filename,error)
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} Error on {}: '.format(logsource, today.strftime('%a %d %b %Y at %I:%M %p'))+errmsg+'\n\n' + data)
			return errmsg
	#Database Log
	if logname == 'Database':
		f = inputs['log-directory']+inputs['database-log']
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} table was successfully updated on {}\n'.format(logsource, today.strftime('%a %d %b %Y at %I:%M %p')) + data)
		if spacers==True:
			with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{:*<50} {}'.format('','\n') + data)
	#Holiday Log
	if logname == 'Holiday':
		f = inputs['log-directory']+inputs['database-log']
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('Holiday detected on {}. {} table was not updated. Message logged on {}\n'.format(yesterday.strftime('%d %b %Y'), logsource, today.strftime('%a %d %b %Y at %I:%M %p')) + data)
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{:*<50} {}'.format('','\n') + data)
	#Alert Log
	if logname == 'Alert':
		f = inputs['log-directory']+inputs['database-log']
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{}\n'.format(logsource) + data)
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{:*<50} {}'.format('','\n') + data)
			
def to_string(x):
	#Function to round floating point numbers to 6 places and then convert them to strings
	if type(x) == str:
		return x
	else:
		return str(round(x,6))

def access_update(inputs, df, cols, access_table):
	#Function to update Access databases
	#Access file
	dbfile = inputs['database-file']
	cnxn = pyodbc.connect('DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ={}'.format(dbfile))
	cursor = cnxn.cursor()
	#Loop through each dataframe row and insert the relevant data into the Access database
	for row in range(0,len(df)):
		cursor.execute("INSERT INTO [{}] ({}) VALUES ({})".format(access_table,', '.join(['['+item+']' for item in cols]),', '.join(["'"+to_string(value)+"'" for value in df.loc[row,(cols)].values])))
	cnxn.commit()
	cnxn.close()

def gas_futures(inputs):
	#Function to update the Gas Futures table
	try:
		#Dataframe columns
		cols = ['Report Date','Sale Date','Sale Year','Sale Month','Sale Day',
	        'Delivery Date','Delivery Year','Delivery Month',
	        'Volume','High','Low','Change','Settlement','Previous Settlement']

	    #Read in the henry hub data
		hh = pd.read_excel('{}{}/{}/hh-gas-futures-{}.xlsx'.format(inputs['henry-hub-directory'],yesterday.year, yesterday.strftime('%B %Y'),file_date),skiprows=1,skip_footer=1,na_values='-')    

		#Get the "Last Updated" date from the henry hub data -- it's in cell A1 of the Excel file
		hh_date = pd.read_excel('{}{}/{}/hh-gas-futures-{}.xlsx'.format(inputs['henry-hub-directory'],yesterday.year, yesterday.strftime('%B %Y'),file_date))
		sale_date = datetime.datetime.strptime(hh_date.iloc[:,0].name.split('Last Updated:')[-1], ' %A, %d %b %Y %I:%M %p') 

		#Read in the previous henry hub data -- the day before the "yesterday" date. If it isn't there, use a dummy dataframe of the same dimensions as the hh dataframe
		try:
			hh_previous = pd.read_excel('{}{}/{}/hh-gas-futures-{}.xlsx'.format(inputs['henry-hub-directory'],two_days.year, two_days.strftime('%B %Y'),two_days.strftime('%m-%d')),skiprows=1,skip_footer=1,na_values='-')
			#Replace the henry hub abbreviated month of July (JLY) with JUL
			hh_previous_delivery = [dates.replace('JLY','JUL') for dates in hh_previous.Month.values]
			hh_previous_delivery = [datetime.datetime.strptime(dates,'%b %y') for dates in hh_previous_delivery]
			#Name the Delivery columns
			hh_previous['Delivery'] = hh_previous_delivery
			#Set index of hh and hh_previous to Delivery/Futures date and remove NaN's
			hh_previous = hh_previous.set_index('Delivery').fillna(value='')
			#Rename the Settle column in the hh_previous dataframe, so as not to be confused with the Settle column in the hh dataframe
			hh_previous = hh_previous.rename(columns={'Settle':'Previous Settlement'})
		except:
			hh_previous=pd.DataFrame(index=hh.index,columns=hh.columns).fillna('').rename(columns={'Settle':'Previous Settlement'})
			
		#Replace the henry hub abbreviated month of July (JLY) with JUL
		hh_delivery = [dates.replace('JLY','JUL') for dates in hh.Month.values]
		hh_delivery = [datetime.datetime.strptime(dates,'%b %y') for dates in hh_delivery]
		
		#Name the Delivery columns
		hh['Delivery'] = hh_delivery
		
		#Set index of hh and hh_previous to Delivery/Futures date and remove NaN's
		hh = hh.set_index('Delivery').fillna(value='')

		#Concat the hh dataframe with the Previous Settlement data from the hh_previous dataframe
		toaccess = pd.concat([hh,hh_previous['Previous Settlement']],axis=1,join_axes=[hh.index])

		#Calculate the length of the dataframe
		length = len(toaccess)

		#Rename the other column names to match the names in the database
		toaccess = toaccess.rename(columns={'Estimated Volume':'Volume','Settle':'Settlement'})

		#Add additional columns not present in the original data, but are present in the database
		toaccess['Sale Year']=[sale_date.year]*length
		toaccess['Sale Month']=[sale_date.month]*length
		toaccess['Sale Day']=[sale_date.day]*length
		toaccess['Sale Date']=[sale_date.date().strftime('%m/%d/%Y')]*length
		toaccess['Delivery Year'] = toaccess.index.year.tolist()
		toaccess['Delivery Month'] = toaccess.index.month.tolist()
		#toaccess['Report Date'] = [pd.to_datetime(np.busday_offset(sale_date.date(),1)).strftime('%m/%d/%Y')]*length

		#Reset the index to 0...n based indexing
		toaccess = toaccess.reset_index()

		#Convert the Delivery Date column into mm/dd/yyyy format
		toaccess['Delivery Date'] = [pd.to_datetime(date).strftime('%m/%d/%Y') for date in toaccess['Delivery'].values]

		#Only pick out the relevant columns, per the col list specified above
		toaccess = toaccess.loc[:,cols].fillna(value='')

		#Update the Access Database
		#Check to make sure yesterday's data is valid. If the sale date doesn't equal yesterday's date, then the data is a repeat:
		if sale_date.strftime('%m %d') == yesterday.strftime('%m %d'):
			hh_today = pd.read_excel('{}{}/{}/hh-gas-futures-{}.xlsx'.format(inputs['henry-hub-directory'],today.year, today.strftime('%B %Y'),today.strftime('%m-%d')))
			sale_date_today = datetime.datetime.strptime(hh_today.iloc[:,0].name.split('Last Updated:')[-1], ' %A, %d %b %Y %I:%M %p')
			#Check to make sure today is not a holiday
			if sale_date_today.strftime('%m %d') == today.strftime('%m %d'):
				toaccess['Report Date'] = [pd.to_datetime(np.busday_offset(sale_date.date(),1)).strftime('%m/%d/%Y')]*length
			#If today is Thanksgiving, then the day after (Friday) is ALSO a holiday, so the report date should be the following Monday
			#elif today.weekday() == 3 and 22 <= today.day <= 28 and today.month == 11
			#	toaccess['Report Date'] = [pd.to_datetime(np.busday_offset(today.date(),2)).strftime('%m/%d/%Y')]*length
			#If today's date doesn't match the sale date in today's data, then today is a holiday. The report date should be one business day from TODAY
			else:
				toaccess['Report Date'] = [pd.to_datetime(np.busday_offset(today.date(),1)).strftime('%m/%d/%Y')]*length
			#Update the database
			access_update(inputs,toaccess,cols,'Gas Futures')		
			
			#Log on successful update
			logs(inputs, 'Database', 'Gas Futures')
		else:
			#Log that it's a holiday
			logs(inputs, 'Holiday', 'Gas Futures')

	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'Gas Futures Database Error', logs(inputs, 'Error', 'Gas Futures Database', e))

def basis_swaps(inputs):
	#Function to update the Basis Swaps database
	try:
		#List to store each dataframe of products in
		df_list = []

		#Read in the henry hub data
		hh = pd.read_excel('{}{}/{}/hh-gas-futures-{}.xlsx'.format(inputs['henry-hub-directory'],yesterday.year, yesterday.strftime('%B %Y'),file_date),skiprows=1,skip_footer=1)
		#Replace the henry hub abbreviated month of July ('JLY') with 'JUL'
		hh_futures = [dates.replace('JLY','JUL') for dates in hh.Month.values]
		hh_futures = [datetime.datetime.strptime(date,'%b %y').strftime('%m/%d/%Y') for date in hh_futures]
		hh['Futures'] = hh_futures
		hh = hh.set_index('Futures').loc[:,('Settle','Estimated Volume')]
		hh = hh.rename(columns={'Settle':'Henry Hub Swaps'})

		#Get raw Basis Swaps (aka NYMEX) data
		data_raw = pd.read_csv('{}/{}/{}/nymex_future{}.csv'.format(inputs['basis-swaps-directory'],yesterday.year, yesterday.strftime('%B %Y'),file_date))

		#Dataframe columns
		products = ['PG&E Citygate Fixed', 'SoCal Natural Gas Fixed', 
	            'Sumas', 'San Juan', 'Permian Basin', 
	            'Henry Hub','West Rockies','Waha']
		#Product codes
		product_keys = ['XQ', 'XN', 'NK', 'NJ', 'PM', 'HB', 'NR', 'NW']
		#Search fields
		fields = ['TRADEDATE','SETTLE','PRODUCT DESCRIPTION','CONTRACT YEAR','CONTRACT MONTH']
		#Access database columns
		cols = ['Sales Date','Futures Date', 'PG&E Citygate','SoCal','Sumas','San Juan','Permian Basin','Henry Hub','West Rockies','Waha']

		#Get the data for each product key
		for step, product in enumerate(product_keys):
		    data = data_raw[data_raw.loc[:,'PRODUCT SYMBOL'] == product].loc[:,fields].reset_index(drop=True)
		    futures_dates = data.loc[:,('CONTRACT YEAR', 'CONTRACT MONTH')].values
		    futures = [datetime.datetime.strptime(np.array_str(date), '[%Y %m]').strftime('%m/%d/%Y') for date in futures_dates]
		    data['Futures'] = futures
		    data = data.rename(columns={'SETTLE':products[step]})
		    data = data.set_index('Futures')
		    df_list.append(data)

		#Append the henry hub data to the dataframe list
		df_list.append(hh)

		#Concat all the dataframes in the list
		merged = pd.concat(df_list,axis=1,join_axes=[hh.index])

		#Extract only the products of interest
		merged = merged.loc[:,products]

		#Drop all the rows where EVERY field is 'NaN'
		merged = merged.dropna(axis=0,how='all')

		#Add a column with the date of the data
		merged['Sales Date']=[yesterday.strftime('%m/%d/%Y')] * len(merged)

		#Reset the index and rename the Futures column to 'Futures Date'
		merged = merged.reset_index().rename(columns={'Futures':'Futures Date'})

		#Calculate the PG&E Citygate and SoCal prices by subtracting the Henry Hub prices
		merged['PG&E Citygate'] = merged['PG&E Citygate Fixed'] - hh.reset_index()['Henry Hub Swaps']
		merged['SoCal'] = merged['SoCal Natural Gas Fixed'] - hh.reset_index()['Henry Hub Swaps']

		#Remove the NaN's from the dataframe
		toaccess = merged.fillna(value='')

		#Update the Access Database
		#Verify that the data is valid. If yesterday's date does NOT equals the Trade date in the nymex file, then yesterday was a holiday and the data should not be added to the database
		if data_raw['TRADEDATE'][0] == yesterday.strftime('%m/%d/%Y'):
			access_update(inputs,toaccess,cols,'Basis Swaps')
			#Log on successful update
			logs(inputs, 'Database', 'Basis Swaps')
		else:
			#Log that it's a holiday
			logs(inputs, 'Holiday', 'Basis Swaps')

	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'Basis Swaps Database Error', logs(inputs, 'Error', 'Basis Swaps Database', e))

def nodal_prices(inputs):
	#Function to update the Basis Swaps Nodal database
	try:
		#List to store each dataframe of products in
		df_list = []

		#Get raw Basis Swaps (aka NYMEX) data
		data_raw = pd.read_csv('{}/{}/{}/nymex_future{}.csv'.format(inputs['basis-swaps-directory'],yesterday.year, yesterday.strftime('%B %Y'),file_date))

		#Dataframe columns
		products = ['PG&E Citygate Fixed', 'SoCal Natural Gas Fixed']
		#Product codes
		product_keys = ['XQ', 'XN']
		#Search fields
		fields = ['SETTLE','TRADEDATE','PRODUCT DESCRIPTION','CONTRACT YEAR','CONTRACT MONTH']

		#Access database columns
		cols = ['Sales Date','Futures Date', 'PG&E Citygate Fixed','SoCal Natural Gas Fixed']

		#Get the data for each product key
		for step, product in enumerate(product_keys):
			data = data_raw[data_raw.loc[:,'PRODUCT SYMBOL'] == product].loc[:,fields].reset_index(drop=True)
			futures_dates = data.loc[:,('CONTRACT YEAR', 'CONTRACT MONTH')].values
			futures = [datetime.datetime.strptime(np.array_str(date), '[%Y %m]') for date in futures_dates]
			data['Futures'] = futures
			data = data.rename(columns={'SETTLE':products[step]})
			data = data.set_index('Futures').drop(fields[1:],1)
			df_list.append(data)

		#Concat all the dataframes in the list and fill in the NaNs with blanks
		merged = pd.concat(df_list,axis=1).fillna('')

		#Rename the Index
		merged.index.names = ['Futures Date']

		#Reset the Index for Access
		merged = merged.reset_index()

		#Add a column with the date of the data
		merged['Sales Date']=[yesterday.strftime('%m/%d/%Y')] * len(merged)

		#Convert the Futures Date column to a string
		merged['Futures Date'] = [date.strftime('%m/%d/%Y') for date in merged['Futures Date']]

		#Update the Access Database
		#Verify that the data is valid. If yesterday's date does NOT equals the Trade date in the nymex file, then yesterday was a holiday and the data should not be added to the database
		toaccess = merged
		if data_raw['TRADEDATE'][0] == yesterday.strftime('%m/%d/%Y'):
			access_update(inputs,toaccess,cols,'Nodal Gas Prices')
			#Log on successful update
			logs(inputs, 'Database', 'Nodal Gas Prices')
		else:
			#Log that it's a holiday
			logs(inputs, 'Holiday', 'Nodal Gas Prices')

	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'Nodal Gas Prices Database Error', logs(inputs, 'Error', 'Nodal Gas Prices Database', e))

def gas_prices(inputs):
	#Function to update the Gas Daily/Prices database
	try:
		#Access columns
		cols = ['Report Date','Transaction Date','Flow Start Date','Hub','Volume','Midpoint','Region','Hub ID']

		#Read in the gas daily data
		gd = pd.read_excel('{}{}/{}/gas-daily-{}.xlsx'.format(inputs['gas-daily-directory'],today.year,today.strftime('%B %Y'), today.strftime('%m-%d')), skiprows=5)

		#Remove all the '--' in the dataframe
		#gd = gd.replace('--.*$','', regex=True)
		gd = gd.replace('^\s*[-*]*\s*$', '', regex=True)
		
		#Remove any trailing or leading spaces in the dataframe
		gd = gd.replace(' *$', '', regex=True)

		#Get the header information
		gd_header = pd.read_excel('{}{}/{}/gas-daily-{}.xlsx'.format(inputs['gas-daily-directory'],today.year,today.strftime('%B %Y'), today.strftime('%m-%d')))
		gd_header_next_day = pd.read_excel('{}{}/{}/gas-daily-{}.xlsx'.format(inputs['gas-daily-directory'],next_day.year,next_day.strftime('%B %Y'), next_day.strftime('%m-%d')))
		
		#Reset the index
		gd_header = gd_header.reset_index(drop=True)
		#Drop everything except the actual header information
		gd_header = gd_header.drop(gd_header.index[3:])
		
		#Get the length of the dataframe
		length = len(gd)

		#Check if any new hubs have been added. If so, send a notification
		try:
			#Read in the previous gas daily data
			gd_prev = pd.read_excel('{}{}/{}/gas-daily-{}.xlsx'.format(inputs['gas-daily-directory'],yesterday.year,yesterday.strftime('%B %Y'), yesterday.strftime('%m-%d')), skiprows=5)
			#Compare yesterday's hubs with today's hubs
			if len(set(gd.Hub).symmetric_difference(gd_prev.Hub)) > 0:
				#If new hubs have been added, send a notification
				new_hubs = set(gd.Hub).symmetric_difference(gd_prev.Hub)
				msg = 'Gas Daily hub change detected on {} involving the following hubs: {}'.format(today.strftime('%d %B %Y at %I:%M %p'), new_hubs)
				logs(inputs, 'Alert', msg)
				mail(inputs,'Gas Daily Hub Alert', msg)
		except:
			pass
			
		#Verify Header information is correct
		#First check if "raw" header was downloaded
		if gd_header.iloc[1,0] == gd_header.iloc[1,3]:
			#Verify the National Prices are equal
			if float(gd_header.iloc[0,0].split(':')[-1].replace(' ','')) != float(gd_header.iloc[0,4]):
				#If they aren't equal, send an notification
				logs(inputs, 'Error', 'Gas Daily', 'Gas Daily National Price Mismatch')
				mail(inputs, 'Gas Daily National Price Mismatch', 'Gas Daily National Price mismatch detected on {}'.format(today.strftime('%d %B %Y at %I:%M %p')))
			#Verify the Transaction dates are equal
			if gd_header.iloc[1,1].replace('/','-').rjust(5,'0') != gd_header.iloc[1,4]:
				#If they aren't equal, send a notification
				logs(inputs, 'Error', 'Gas Daily', 'Gas Daily Transaction Date Mismatch')
				mail(inputs, 'Gas Daily Transaction Date Mismatch', 'Gas Daily Transaction Date mismatch detected on {}'.format(today.strftime('%d %B %Y at %I:%M %p')))
			#Get the transaction date from the raw header
			trans_date = datetime.datetime.strptime(gd_header.iloc[1,1]+'/'+str(today.year),'%m/%d/%Y').date()#.strftime('%m/%d/%Y')	
		else:
			#If "raw" header not present, then only modified header is. Get the transaction date
			trans_date = datetime.datetime.strptime(gd_header.iloc[1,1]+'-'+str(today.year),'%m-%d-%Y').date()#.strftime('%m/%d/%Y')

		#Verify the flow date is correct. If the file was downloaded on a holiday, then use today's date as the flow date-- NOT the flow date in the file.
		if gd_header.iloc[2,4] != today.strftime('%m-%d'):
			flow_date = today#.strftime('%m/%d/%Y')
		else:
			flow_date = datetime.datetime.strptime(gd_header.iloc[2,4]+'-'+str(today.year),'%m-%d-%Y').date()#.strftime('%m/%d/%Y')

		#Check to see if the transaction date and flow date should be in the same year
		if abs(flow_date.month - trans_date.month) > 1:
			trans_date = datetime.datetime(today.year-1, trans_date.month, trans_date.day).strftime('%m/%d/%Y')
		else:
			trans_date = trans_date.strftime('%m/%d/%Y')
			
		#Format the flow date
		flow_date = flow_date.strftime('%m/%d/%Y')
		
		#Get the national Avg Price
		nat = gd_header.iloc[0,4]
		
		#Determine if the current file is a holiday. Open the next day's file. If the flow date in that file does not equal the date of the next day's file, then today is a holiday. Example: Thanksgiving is on November 24. On November 24, open the gas_daily file for the next business day, November 25. Remember, the date in the gas daily files reflects +1 business day from the date it was saved. So the gas-daily-11-25.xlsx file is downloaded and saved on November 24. The flow date in the gas-daily-11-25.xlsx file is 11/24, which is NOT equal to 11-25. This means the 24th is a holiday.
		if gd_header_next_day.iloc[2,4] != next_day.strftime('%m-%d'):
			#Append the file with "holiday"
			path = '{}{}/{}/'.format(inputs['gas-daily-directory'],today.year,today.strftime('%B %Y'))
			os.rename(path+'gas-daily-{}.xlsx'.format(today.strftime('%m-%d')),path+'gas-daily-{}_holiday.xlsx'.format(today.strftime('%m-%d')))
			#Save a message in the log
			logs(inputs, 'Alert', 'Holiday detected on {0}. The file gas-daily-{1}.xlsx was renamed to gas-daily-{1}_holiday.xlsx'.format(today.strftime('%a %d %b %Y'),today.strftime('%m-%d')))
		
		'''#########################################################################
		   ###THIS LOGIC NO LONGER USED. SAVED ONLY FOR ARCHIVAL PURPOSES. IGNORE###
		   #########################################################################	 
		#Verify the Flow date is equal to today's date. If it isn't and today is NOT Monday, then the data is from a holiday. Use
		#today's date as the flow date when updating the database, instead of the date in the file. Append the file with "holiday"	
		if gd_header.iloc[2,1].replace('/','-').rjust(5,'0') != today.strftime('%m-%d') and today.strftime('%A') != 'Monday':
			#Use today's date as the flow date
			flow_date = today.strftime('%m/%d/%Y')
			#Get the National Avg Price in the header
			nat = gd_header.iloc[0,4]
			#Append the file with "holiday"
			path = '{}{}/{}/'.format(inputs['gas-daily-directory'],today.year,today.strftime('%B %Y'))
			os.rename(path+'gas-daily-{}.xlsx'.format(today.strftime('%m-%d')),path+'gas-daily-{}_holiday.xlsx'.format(today.strftime('%m-%d')))
		else:
			if gd_header.iloc[1,0] == gd_header.iloc[1,3]:
				flow_date = datetime.datetime.strptime(gd_header.iloc[2,4]+'-'+str(today.year),'%m-%d-%Y').date().strftime('%m/%d/%Y')
				#Get the National Avg Price in the header
				nat = gd_header.iloc[0,4]
			else:
				flow_date = datetime.datetime.strptime(gd_header.iloc[2,1]+'-'+str(today.year),'%m-%d-%Y').date().strftime('%m/%d/%Y')
				#Get the National Avg Price in the header
				nat = gd_header.iloc[0,1]

		'''		
		#Get the National Volume (sum of all the hub volumes)
		vol = gd.Volume.replace('','0').astype(float).sum()

		#Store the National data in a dictionary
		nat_data = {'Hub':'National','Volume':vol,'Midpoint':nat, 'Report Date':flow_date,'Transaction Date':trans_date,'Flow Start Date':flow_date}

		#Add the various date columns needed in the Access database
		gd['Report Date'] = [flow_date]*length
		gd['Transaction Date'] = [trans_date]*length
		gd['Flow Start Date'] = [flow_date]*length
		gd = gd.rename(columns={'ID':'Hub ID'})

		#Ready the dataframe to import to Access
		toaccess = gd.append(nat_data,ignore_index=True).fillna(value='').loc[:,cols]

		#Update the Access Database
		#NOTE: This gets called before the weekend logic because the FOR loop modifies the Flow Start Dates in the toaccess dataframe. If the weekend
		#logic was called before the initial access database update, then the Flow Start Date would have to be changed back to the original Flow Start 
		#Date. 
		access_update(inputs,toaccess,cols,'Gas Daily')	
		
		#If it's Monday, update the Database with the weekend flow dates
		if today.strftime('%A') == 'Monday':
			for weekend in range(2,0,-1): 
				toaccess['Flow Start Date'] = [(today-datetime.timedelta(weekend)).strftime('%m/%d/%Y')]*(length+1)
				access_update(inputs,toaccess,cols,'Gas Daily')

		#Log on successful update
		logs(inputs, 'Database', 'Gas Daily')

	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'Gas Daily Prices Database Error', logs(inputs, 'Error', 'Gas Daily Prices Database', e))
	
def megawatt_daily(inputs):
	#Function to update the Megawatt Daily database
	try:
		#Open the megawatt daily file
		md = pd.read_excel('{}{}/{}/megawatt-daily-{}.xlsx'.format(inputs['megawatt-daily-directory'], today.strftime('%Y'), today.strftime('%B %Y'), today.strftime('%m-%d')), skiprows=3,index_col=[0,1,2],na_values='N.A.').fillna('').reset_index()
		md_bidate = pd.read_excel('{}{}/{}/megawatt-daily-{}.xlsx'.format(inputs['megawatt-daily-directory'], today.strftime('%Y'), today.strftime('%B %Y'), today.strftime('%m-%d')), header=None)
		
		#Get the delivery date. If "raw" header is present, go down a row and get the date
		if md_bidate.iloc[1,2] == '($/MWh)':
			del_date = datetime.datetime.strptime(md_bidate.iloc[1,1],'%b %d %Y')
			raw_del_date = datetime.datetime.strptime(md_bidate.iloc[0,0].split('delivery ')[-1],'%b %d ($/MWh)')
			#If "raw" delivery date doesn't match the other date, send an alert and record it to the log
			if raw_del_date.strftime('%b %d') != del_date.strftime('%b %d'):
				logs(inputs, 'Error', 'Megawatt Daily', 'Bi-ahead Bilateral Delivery Date Mismatch')
				mail(inputs, 'Megawatt Daily Delivery Date Mismatch', 'Megawatt Daily Delivery Date mismatch detected on {}'.format(today.strftime('%d %B %Y at %I:%M %p')))
		#If no "raw" header present, use the only date in the file
		if md_bidate.iloc[0,2] == '($/MWh)':
			del_date = datetime.datetime.strptime(md_bidate.iloc[0,1],'%b %d %Y')
		
		#Add a column with the delivery date
		md['Delivery Date']=[del_date.strftime('%m/%d/%Y')]*len(md)
		
		#Rename columns to match those in Access
		md = md.rename(columns={'ID':'Hub ID', 'Product':'Hub'})
		
		#Make sure Period Column is title case
		md['Period']=[item.title() for item in md['Period']]
		
		#Update the Access database
		cols = ['Delivery Date', 'Region', 'Period', 'Hub', 'Hub ID', 'Index', 'Change', 'Range', 'Deals', 'Volume', 'Avg $/Mo']
		access_update(inputs,md,cols,'Megawatt Daily')
		
		#Log on successful update
		logs(inputs, 'Database', 'Megawatt Daily')
		
	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'Megawatt Daily Database Error', logs(inputs, 'Error', 'Megawatt Daily Database', e))
		
	#This code is for megawatt daily files before 8 March 2016
	# try:
		# #Open the file
		# md = pd.read_excel('{}{}/{}/megawatt-daily-{}.xlsx'.format(inputs['megawatt-daily-directory'], today.strftime('%Y'), today.strftime('%B %Y'), today.strftime('%m-%d')), skiprows=1,na_values='N.A.')
		# md_date = pd.read_excel('{}{}/{}/megawatt-daily-{}.xlsx'.format(inputs['megawatt-daily-directory'], today.strftime('%Y'), today.strftime('%B %Y'), today.strftime('%m-%d')),header=None)
		
		# #Loop through and generate a period vector
		# perix=[]
		# per=[]
		# drop=[]
		# md = md.replace('\*','',regex=True)
		# msk = (md.iloc[:,0] == 'On-peak') | (md.iloc[:,0] == 'Off-Peak')
		# idx = md[msk].index
		# for step in range(0,len(idx)-1):
			# perix.append(idx[step+1]-idx[step])
		# perix.append(len(md)-idx[-1])
		# for step,value in enumerate(perix):
			# tmp = [md.iloc[idx[step],0]]*value
			# per.extend(tmp)
		# #Ensure period vector is title case
		# md['Period'] = [item.title() for item in per]
		
		# #Append the indexes to a drop vector, so they can be dropped at the end
		# drop.extend(idx.values.tolist())
		
		# #Loop through and generate a region vector
		# reg=[]
		# msk = md.loc[:,'Southeast'] == 'West'
		# idx=md[msk].index
		# reg.extend(['Southeast']*idx[0])
		# reg.extend([md.loc[idx[0],'Southeast']]*(len(md)-idx[0]))
		# md['Region'] = reg

		# #Append the indexes to a drop vector, so they can be dropped at the end
		# drop.extend(idx.values.tolist())

		# #Drop the rows in the dataframe with jumbled headers
		# md = md.iloc[md.drop(drop)['Index'].dropna().index,:].fillna('')
		
		# #Add Delivery date
		# md['Delivery Date'] = [datetime.datetime.strptime(md_date.iloc[0,0].split('delivery ')[-1]+str(today.year),'%b %d ($/MWh)%Y').strftime('%m/%d/%Y')]*len(md)
		
		# #Rename columns to match Access
		# md = md.rename(columns={'Unnamed: 0':'Hub ID','Southeast':'Hub'})
		# md = md.reset_index(drop=True)
		
		# #Update Access
		# cols = ['Delivery Date', 'Region', 'Period', 'Hub', 'Hub ID', 'Index', 'Change', 'Range', 'Deals', 'Volume', 'Avg $/Mo']
		# access_update(inputs,md,cols,'Megawatt Daily')		
		
		# #Log on successful update
		# logs(inputs, 'Database', 'Megawatt Daily',spacers=False)
		
	# except Exception as e:
		# #Email and log errors, if any	
		# mail(inputs, 'Megawatt Daily Database Error', logs(inputs, 'Error', 'Megawatt Daily Database', e))
	
def caiso(inputs):
	try:
		#Open today's CAISO data file
		caiso = pd.read_csv(glob.glob('{}{}/{}/{}*.csv'.format(inputs['caiso-directory'], today.strftime('%Y'), today.strftime('%B %Y'), today.strftime('%Y%m%d')))[0])
		
		#Save the columns names to a list
		cols = caiso.columns
		
		#Update the CAISO table
		access_update(inputs, caiso, cols, 'CAISO')
		
		#Log on successful update
		logs(inputs, 'Database', 'CAISO', spacers=True)
		
	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'CAISO Database Error', logs(inputs, 'Error', 'CAISO Database', e))
	
def ice(inputs, icedate=yesterday):
	try:
		#Get a list of all of yesterday's ICE files
		#icedir = glob.glob('{}{}/{}/*_{}.xlsx'.format(inputs['ice-directory'], icedate.strftime('%Y'), icedate.strftime('%B %Y'), icedate.strftime('%m-%d-%Y')))
		icedir = glob.glob('{}{}/{}/*_{}.xlsx'.format(inputs['ice-directory'], icedate.strftime('%Y'), icedate.strftime('%B %Y'), icedate.strftime('%Y_%m_%d')))
		
		#Loop through each file and add to Access
		for file in icedir:
			#Open each file in a dataframe
			ice = pd.read_excel(file).fillna(value='')
		
			#Save the columns names to a list
			cols = ice.columns
		
			#Update the ICE table
			access_update(inputs, ice, cols, 'ICE')
		
		#Log on successful update
		logs(inputs, 'Database', 'ICE')
	
	except Exception as e:
		#Email and log errors, if any	
		mail(inputs, 'ICE Database Error', logs(inputs, 'Error', 'ICE Database', e))

def nodal_exchange(inputs):
	try:
		#Open today's Nodal Exchange data file
		nodal = pd.read_excel('{}{}/{}/EOD_FUTURES_REPORT_{}.xlsx'.format(inputs['nodal-exchange-directory'], today.strftime('%Y'), today.strftime('%B %Y'), today.strftime('%d-%b-%y')))
		
		#Convert the Expiry column from datetime format to string
		nodal['Expiry'] = [expdate.strftime('%m/%d/%y') for expdate in nodal['Expiry']]

		#Save the columns names to a list
		cols = nodal.columns
		
		#Update the Nodal Exchange Futures table
		access_update(inputs, nodal, cols, 'Nodal Exchange Futures')
		
		#Log on successful update
		logs(inputs, 'Database', 'Nodal Exchange Futures')
		
	except Exception as e:
		#Email and log errors, if any	
		#mail(inputs, 'Nodal Exchange Futures Database Error', logs(inputs, 'Error', 'Nodal Exchange Futures Database', e))
		print(e)
	
if __name__ == "__main__": main()