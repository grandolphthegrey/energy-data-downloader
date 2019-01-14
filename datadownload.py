#!/usr/bin/env python3

import urllib.request
import requests
import datetime
import os
import sys
import subprocess
import zipfile
import re
import smtplib
import codecs
import time
import xml

import numpy as np
import pandas as pd
import email.encoders as Encoders

from json2table import convert
from bs4 import BeautifulSoup
from lxml import html
from pandas import ExcelWriter
from openpyxl import load_workbook
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from collections import OrderedDict
from shutil import copyfile
from lxml import etree

####################
# Global Variables #
####################

#Today's date
today = datetime.datetime.today()
#today=datetime.datetime(2019,1,3, datetime.datetime.now().time().hour,datetime.datetime.now().time().minute,datetime.datetime.now().time().second)
#today = datetime.datetime(2018,11,13,17,3,22)

#Yesterday's date
yesterday = pd.to_datetime(str(np.busday_offset(today, -1, roll='backward')))
file_date = yesterday.strftime('%m')+'-'+yesterday.strftime('%d')
#The day before yesterday
two_days = pd.to_datetime(str(np.busday_offset(today, -2, roll='backward')))

################
# Main Program #
################

def main():
	
	#Open the "settings.txt" file to define settings, URLS, directories, etc for downloading all the data.
	settings={}
	with open('/Users/Dev/settings.txt') as f:
		for line in f:
			if '#' not in line and len(line)>1:
				(key,val)=line.strip().split(':',1)
				settings[key]=val
	
	#Determine if data should be save on the Server or locally. If saved on the Server, then mount it. The scripts runs at 5:03 every evening and that data is saved to the server. The script runs again at 9:03 and that data is saved locally. This second run acts as a backup.
	check_hour = 17
	if today.hour == check_hour:
		#Mount the network drive
		mount_drive(settings)
		save_dir = settings['server-save']
		log_dir = settings['server-log']
		
	else:
		save_dir = settings['local-save']
		log_dir = settings['local-log']			

	#Download the various data sources. The CAISO data gets downloaded every day. The other 4 data sources only get downloaded on weekdays.
	try:

		if today.weekday() < 5:
			
			#Download Henry Hub data
			henry_hub(settings, save_dir, log_dir)

			#Download Basis Swaps data
			basis_swaps(settings, save_dir, log_dir)

			#Download Gas Daily data
			gas_daily(settings, save_dir, log_dir)

			#Download Megawatt Daily data
			megawatt_daily(settings, save_dir, log_dir)

			#Download ICE data
			if today.hour == check_hour:
				ice(settings, save_dir, log_dir, icedate=pd.to_datetime(str(np.busday_offset(today, -3, roll = 'backward'))))	
			
			#Download Nodal Exchange Data
			nodal_exchange(settings, save_dir, log_dir)
	
		#If today is Friday, then convert last week's ICE files
		#if today.weekday() == 4 and today.hour == check_hour:
		#	for daystep in np.arange(-9, -4, 1):
		#		icedate = pd.to_datetime(str(np.busday_offset(today, daystep, roll = 'backward')))
		#		ice(settings, save_dir, log_dir, icedate)
		
		#Download CAISO data
		caiso(settings, save_dir, log_dir)

	except Exception as e:
		#Email and log errors, if any
		mail(settings, 'Main Data Download Program Error', logs(settings, log_dir, 'Error', 'Main Data Download Program', e))

	#Unmount the network drive -- this is outside the TRY loop in case an error is thrown. This ensures that the drive is unmounted.
	if today.hour == check_hour:
		unmount_drive(settings)

#############
# Functions #
#############

def mount_drive(inputs):
	#Function to mount the network drive
	#Check if sharepoint is mounted. If not, mount it.
	if not os.path.ismount(inputs['sharepoint-energy']) and not os.path.exists(inputs['sharepoint-energy']):
		os.makedirs(inputs['sharepoint-energy'])
		subprocess.call(['mount','-t', 'smbfs', '//'+inputs['server-username']+':'+inputs['server-password']+'@'+inputs['server-address-energy'], inputs['sharepoint-energy']])

	#Check if sharepoint already exists. If it does, then just mount the network drive
	if os.path.exists(inputs['sharepoint-energy']) and not os.path.ismount(inputs['sharepoint-energy']):
		subprocess.call(['mount','-t', 'smbfs', '//'+inputs['server-username']+':'+inputs['server-password']+'@'+inputs['server-address-energy'], inputs['sharepoint-energy']])

def unmount_drive(inputs):
	#Function to unmount network drive
	if os.path.ismount(inputs['sharepoint-energy']):
		subprocess.call(['umount', inputs['sharepoint-energy']])

def checkdir(inputs, save_dir, data_dir, thedate=today):
	#Note that thedate defaults to today if nothing is specified

	#If the user specified date is more than 1 month more than today's month, then the user 
	#specified date is in the next year: 
	#Example: today is 31 December 2016 and the user specified date is 3 January. The difference
	#in those months (1 - 12) is more than 1, so the user specified date is a few days in the future
	#and also in the next year.
	#if abs(thedate.month - today.month) > 1:
	if thedate.month - today.month == -11:
		yr = today.year + 1
		check_dir = save_dir+inputs[data_dir]+'{}/{} {}/'.format(yr,thedate.strftime('%B'),yr)
	#If the user specified date is in December and the today is in the month of January, then the year in the user specified date is the previous year
	elif thedate.month - today.month == 11:
		yr = today.year - 1
		check_dir = save_dir+inputs[data_dir]+'{}/{} {}/'.format(yr,thedate.strftime('%B'),yr)	
	#If the user specified date is in the past, then no year logic is needed.
	#elif thedate.month - today.month > 1:
	#	check_dir = save_dir+inputs[data_dir]+'{}/{} {}/'.format(thedate.year,thedate.strftime('%B'),thedate.year)
	#The (absolute) difference between all other months 
	#(Jan-Feb, Feb-Mar, Mar-Apr, etc) will always be 1 (1-2=1, 2-3=1, 3-4=1, etc)
	#elif abs(thedate.month - today.month) == 1:
	elif abs(thedate.month - today.month) == 1:
		check_dir = save_dir+inputs[data_dir]+'{}/{} {}/'.format(today.year,thedate.strftime('%B'),today.year)
	#In all other cases, the dates are in the same month, so no additional logic is needed
	else:
		check_dir = save_dir+inputs[data_dir]+'{}/{}/'.format(today.year,today.strftime('%B %Y'))
	#If the directories don't exist, then create them
	if not os.path.exists(check_dir):
		os.makedirs(check_dir)
	return check_dir

def logs(inputs, log_dir, logname, logsource, fn=None, error=None, spacers=False):
	#Check if the log files exist. If they don't, create them.
	direc = log_dir
	dl = log_dir+inputs['download-log']
	err = log_dir+inputs['error-log']
	if not os.path.exists(direc):
		os.makedirs(direc)
	if not os.path.isfile(dl):
		open(dl,'w').close()
	if not os.path.isfile(err):
		open(err,'w').close()
	#Download Log
	if logname=='Download':
		f = dl
		#Write the download status to download log
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} data download completed on {}: {}\n'.format(logsource, today.strftime('%a %d %b %Y at %I:%M %p'), fn) + data)
		if spacers==True:
			with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
			with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{:*<50} {}'.format('','\n') + data)
	#Error Log
	if logname=='Error':
		f = err
		#Format the error message to save to log
		errmsg = 'Error on line {} of {}: {}'.format(sys.exc_info()[-1].tb_lineno,sys.exc_info()[-1].tb_frame.f_code.co_filename,error)
		#Save error message to error log
		with codecs.open(f, 'r',encoding='utf-8', errors='ignore') as original: data = original.read()
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} Error on {}: '.format(logsource, today.strftime('%a %d %b %Y at %I:%M %p'))+errmsg+'\n\n' + data)
		return errmsg

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
	mailServer.sendmail(gmail_sender, inputs['email-recipient'].split(','), msg.as_string())
	mailServer.close()

def saveExcel(data_frame, file_name,sr, use_index=False):
	#Define name of excel file
	writefilestr = file_name
	writer = ExcelWriter(writefilestr, engine='openpyxl')
	#Write and save data to 'Sheet1' (default)
	data_frame.to_excel(writer,'Sheet1',index=use_index, startrow=sr,startcol=0)
	writer.save()
	writer.close()

def appendExcel(data_frame, file_name,sc=0,use_index=False):
	#Define and load excel file
	writefilestr = file_name
	book = load_workbook(writefilestr)
	writer = ExcelWriter(writefilestr, engine='openpyxl')
	writer.book = book
	#Get list of sheets in excel file
	writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
	#Write and save data to excel file
	data_frame.to_excel(writer,'Sheet1',index=use_index,startrow=0, startcol=sc,header=None)
	writer.save()
	writer.close()

def platts_login(inputs, source):
	#URLs
	homeurl = inputs['platts-home-url']
	tableurl = inputs['platts-table-url']
	gasurl = inputs['platts-gas-url']
	megurl = inputs['platts-meg-url']

	#Header
	header = {'Host':'pmc.platts.com',
			'Referer':'',
			'User-Agent':inputs['user-agent'],
			'DNT':'1'}

	#Passwords and login data
	credentials = {'ctl00$cphMain$txtEmail':inputs['platts-username'], 
			'ctl00$cphMain$txtPassword':inputs['platts-password'],
			'ctl00$cphMain$btnLogin':'Submit',
			'ctl00$cphMain$ftLogin$txtFTEmail': '',
			'ctl00$cphMain$ftLogin$ddlNewsletter':'Asian Petrochemicalscan'}

	#Query Dictionary
	payload = {'coverDate':'', 'tableCode':''}

	#Initialize a session
	session = requests.Session()

	#Go to the login page and save the cookies
	site = session.get(homeurl)
	cookies = dict(session.cookies)

	#Save the HTML from the login page and find the hidden form elements
	webpage = BeautifulSoup(site.text, 'lxml')
	hiddenInputs = webpage.find_all(name='input', type='hidden')

	#Loop through the hidden form elements and save them to the credentials dictionary
	for hidden in hiddenInputs:
		name = hidden['name']
		value = hidden['value']
		credentials[name] = value

	#Login to the site with the updated credentials dictionary
	header['Referer'] = homeurl
	login = session.post(url=homeurl,headers=header,data=credentials,cookies=cookies)

	#Save the cookies from the response
	cookies.update(dict(session.cookies))

	#Update the header with new JSON content type
	header['Content-Type']='application/json; charset=UTF-8'

	#Specify the gas daily referral URL and payload table code
	if source == 'gas':
		header['Referer'] = gasurl
		payload['tableCode'] = inputs['platts-gas-tablecode']
	#Specify the megawatt daily referral URL and payload table code
	elif source == 'meg':
		header['Referer'] = megurl
		payload['tableCode'] = inputs['platts-meg-tablecode']
	
	#Request the specified table
	platts_out = session.post(url=tableurl,headers=header,json=payload,cookies=cookies)

	#Save the JSON table output
	table = BeautifulSoup(convert(platts_out.json()),'lxml')

	#Output the formatted table
	return table

def henry_hub(inputs,save_dir,log_dir):
	try:
		#Go to the website and parse the page
		url = inputs['henry-hub-requrl']+today.strftime('%m/%d/%Y')
		'''
		webpage = requests.get(inputs['henry-hub-url'])
		soup = BeautifulSoup(webpage.text,'lxml')
		cols = [soup.thead('th')[colname].get_text() for colname in range(0, len(soup.thead('th')))]
		recs = []

		#Loop through the Henry Hub table and yank out the data
		for dd in range(0,len(soup.tbody('th'))):
			tmp = [soup.tbody('tr')[dd].find_all('td')[jj].get_text() for jj in range(0, len(soup.tbody('tr')[0].find_all('td')))]
			tmp.insert(0,soup.tbody('th')[dd].get_text())
			recs.append(pd.Series(tmp,index = cols))	
		'''
		site = requests.get(url)	

		#Save the data to a dataframe	
		#hh = pd.DataFrame(recs,columns = cols)
		pdout = dict(site.json())
		hh = pd.DataFrame.from_dict(pdout['settlements'], orient='columns')
		hh = hh.rename(columns={'month':'Month', 'open':'Open', 'high':'High', 'low':'Low', 'last':'Last', 'change':'Change',
					   'settle':'Settle', 'volume':'Estimated Volume','openInterest':'Prior Day Open Interest'})
		hh = hh.loc[:,pd.IndexSlice['Month','Open', 'High', 'Low', 'Last', 'Change','Settle', 'Estimated Volume', 'Prior Day Open Interest']]

		#Remove the "As" "Bs" and "Cs" from the Open, High, Low and Last Columns
		colfix=['Open', 'High', 'Low', 'Last']
		for col in colfix:
			for step,item in enumerate(hh[col]): hh[col][step] = hh[col][step][0:5]

		#Save the header to a separate dataframe
		#update = pd.DataFrame([soup.find_all('li',class_='cmeLegendItem cmeIconListItem')[0].get_text()])
		update = pd.DataFrame(['Last Updated: '+pdout['updateTime']])

		#Save the dataframes to Excel
		fn = 'hh-gas-futures-{}.xlsx'.format(today.strftime('%m-%d'))
		f = checkdir(inputs, save_dir, 'henry-hub-directory')+fn
		saveExcel(hh,f,sr=1, use_index=False)
		appendExcel(update, f)

		#Log on successful download
		logs(inputs, log_dir, 'Download', 'Henry Hub', fn)

	except Exception as e:
		#Email and log errors, if any
		mail(inputs, 'Henry Hub Data Download Error', logs(inputs, log_dir, 'Error', 'Henry Hub', e))	

def basis_swaps(inputs,save_dir,log_dir):
	try:
		#Specify where the basis swaps data should be saved to
		filename = 'nymex_future{}.csv'.format(today.strftime('%m-%d'))
		writefilestr = checkdir(inputs, save_dir, 'basis-swaps-directory')+filename
		
		#Download the data
		urllib.request.urlretrieve('ftp://ftp.cmegroup.com/pub/settle/nymex_future.csv',writefilestr)

		#Save to a dataframe for future parsing
		df = pd.read_csv(writefilestr)

		#Log on successful download
		logs(inputs, log_dir, 'Download', 'Basis Swaps',filename)

	except Exception as e:
		#Email and log errors, if any
		mail(inputs, 'Basis Swaps Data Download Error', logs(inputs, log_dir, 'Error', 'Basis Swaps', e))	

def caiso(inputs,save_dir,log_dir):
	try:
		#Specify the filename 
		filename = 'LMP_{}.zip'.format(today.strftime('%Y%m%d'))
		writefilestr = checkdir(inputs, save_dir, 'caiso-directory')+filename
		td = today + timedelta(days = 1)
		
		#Download the data
		URL = 'http://oasis.caiso.com/oasisapi/SingleZip?resultformat=6&queryname=PRC_LMP&version=1&startdatetime={}T08:00-0000&enddatetime={}T08:00-0000&market_run_id=DAM&node=TH_NP15_GEN-APND,TH_SP15_GEN-APND'.format(today.strftime('%Y%m%d'),td.strftime('%Y%m%d'))
		urllib.request.urlretrieve(URL,writefilestr)

		#Extract the zip file
		zf = zipfile.ZipFile(writefilestr,'r')
		zf.extractall(checkdir(inputs, save_dir, 'caiso-directory'))
		zf.close()

		#Log on successful download
		logs(inputs, log_dir, 'Download', 'CAISO', filename, spacers=True)

	except Exception as e:
		#Email and errors, if any
		mail(inputs, 'CAISO Data Download Error', logs(inputs, log_dir, 'Error', 'CAISO', e))

def gas_daily(inputs,save_dir,log_dir):
	try:
		#Get Gas Daily Table via HTTP POST in Platts Login Script
		gas_table = platts_login(inputs, source='gas')

		#Get Column Names
		cols = [gas_table.find('tr', class_="COLUMNHEAD").find_all('td')[colname].get_text() for colname in range(0,len(gas_table.find('tr', class_="COLUMNHEAD").find_all('td')))]
		#Append the two vectors that don't have names: ID and Hub (aka product)
		cols[0] = 'ID'
		cols[1] = 'Hub'
		cols.insert(0,'Region')
		#Some of the column names sometimes change.
		try:
			cols = [name.replace('Vol.','Volume') for name in cols]
		except:
			pass

		#Get the Regions
		regions = gas_table.find_all('tr',class_='UNDERLINE')

		#Create the header
		header_labels = []
		header_data = []
		header_raw = {}

		#Save "raw" Header -- verbatim information from webpage -- format can change, so include TRY statements for all previous formats
		#Title ("Platts Locations...")
		header_raw[gas_table.find('td', class_="TITLE_1").get_text()] = ''
		#National average price
		if 'NATIONAL AVERAGE PRICE' in gas_table.find('td', class_="NOTE_2").get_text():
			header_raw[gas_table.find('td', class_="NOTE_2").get_text()] = ''
		else:
			header_raw[gas_table.find('td', class_="NOTE_1").get_text()] = ''

		#Flow and Transaction Dates
		try:
			header_raw[gas_table.find_all('td', class_="NOTE_2")[1].get_text()] = gas_table.find_all('td', class_="NOTE_3")[0].get_text()
			header_raw[gas_table.find_all('td', class_="NOTE_2")[2].get_text()] = gas_table.find_all('td', class_="NOTE_3")[1].get_text()			
		except:
			pass
		try:
			header_raw[gas_table.find_all('td', class_="NOTE_2")[1].get_text()] = datetime.datetime.strptime(gas_table.find_all('td', class_="NOTE_3")[0].get_text(),'%d-%b').strftime('%m/%d')
			header_raw[gas_table.find_all('td', class_="NOTE_2")[2].get_text()] = datetime.datetime.strptime(gas_table.find_all('td', class_="NOTE_3")[1].get_text(),'%d-%b').strftime('%m/%d')
		except:
			pass
		try:
			header_raw[gas_table.find_all('td', class_="NOTE_2")[1].get_text()] = datetime.datetime.strptime(gas_table.find_all('td', class_="NOTE_3")[0].get_text()[0:6],'%d-%b').strftime('%m/%d')
			header_raw[gas_table.find_all('td', class_="NOTE_2")[2].get_text()] = datetime.datetime.strptime(gas_table.find_all('td', class_="NOTE_3")[1].get_text()[0:6],'%d-%b').strftime('%m/%d')
		except:
			pass
		try:
			header_raw[gas_table.find_all('td', class_="NOTE_1")[1].get_text()] = datetime.datetime.strptime(gas_table.find_all('td', class_="NOTE_2")[0].get_text(),'%d-%b').strftime('%m/%d')
			header_raw[gas_table.find_all('td', class_="NOTE_1")[2].get_text()] = datetime.datetime.strptime(gas_table.find_all('td', class_="NOTE_2")[1].get_text(),'%d-%b').strftime('%m/%d')	
		except:
			pass	
		
		#Convert raw header to DataFrame
		hdrawf = pd.DataFrame(pd.Series(header_raw).sort_values())

		#Save the Units
		units = gas_table.find('td', class_="TITLE_1").get_text().split('(')[1].rstrip(')')
		units_label = gas_table.find('td', class_="TITLE_1").get_text().split('(')[0]
		header_data.append(units)
		header_labels.append(units_label)

		#Save the Average Price
		try:
			if 'NATIONAL AVERAGE PRICE' in gas_table.find('td', class_="NOTE_2").get_text():
				ave_price_label = gas_table.find('td', class_="NOTE_2").get_text().split(':')[0]
				ave_price = float(gas_table.find('td', class_="NOTE_2").get_text().split(':')[-1].replace(' ',''))
		except:
			pass
		try:
			if 'NATIONAL AVERAGE PRICE' in gas_table.find_all('td', class_="NOTE_1")[0].get_text():
				ave_price_label = gas_table.find_all('td', class_="NOTE_1")[0].get_text().split(':')[0]
				ave_price = gas_table.find_all('td', class_="NOTE_1")[0].get_text().split(':')[-1].replace(' ','')
		except:
			pass
		header_data.append(ave_price)
		header_labels.append(ave_price_label)

		#Save the Transaction/Trade Date -- format has changed in the past, so include all previously seen formats 
		try:
			trans_date = gas_table.find_all('td', class_="NOTE_2")[0].get_text().replace('/','-')
			trans_date = datetime.datetime.strptime(trans_date,'%d-%b').strftime('%m-%d')
		except:
			pass
		try:
			trans_date = gas_table.find_all('td', class_="NOTE_3")[0].get_text().replace('/','-')
			trans_date = datetime.datetime.strptime(trans_date,'%m-%d').strftime('%m-%d')
		except:
			pass
		try:
			trans_date = gas_table.find_all('td', class_="NOTE_3")[0].get_text().replace('/','-')
			trans_date = datetime.datetime.strptime(trans_date,'%d-%b').strftime('%m-%d')
		except:
			pass

		header_data.append(trans_date)
		try:
			if 'Trade' in gas_table.find_all('td', class_="NOTE_2")[1].get_text():
				header_labels.append(gas_table.find_all('td', class_="NOTE_2")[1].get_text())
		except:
			pass
		try:
			if 'Trade' in gas_table.find_all('td', class_="NOTE_1")[1].get_text():
				header_labels.append(gas_table.find_all('td', class_="NOTE_1")[1].get_text())
		except:
			pass

		#Save the Flow Date
		flow_date_full = str(np.busday_offset((yesterday.strftime('%Y') + '-' + trans_date),1))
		flow_date = flow_date_full.split('-',1)[-1]
		header_data.append(flow_date)
		try:
			if 'Flow' in gas_table.find_all('td', class_="NOTE_2")[2].get_text():
				header_labels.append(gas_table.find_all('td', class_="NOTE_2")[2].get_text())
		except:
			pass
		try:
			if 'Flow' in gas_table.find_all('td', class_="NOTE_1")[2].get_text():
				header_labels.append(gas_table.find_all('td', class_="NOTE_1")[2].get_text())
		except:
			pass

		#Save Header into a DataFrame
		hdf = pd.DataFrame(pd.Series(header_data,header_labels))

		#Empty list to save the data
		recs = []

		#Loop through and extract the data from the table
		for step, region in enumerate(regions): 
			for line in region.next_siblings:
				if line == regions[-len(regions)+1+step]:
					break 
				else:
					if len(line) != 1:
						tmp = [line.find_all('td',class_=re.compile('ROW_*'))[ii].get_text() for ii in range(0,len(line.find_all('td',class_=re.compile('ROW_*'))))]
						if tmp != [] and not any([re.search('average', entity) for entity in tmp]):
							tmp.insert(0,region.get_text().strip())
							recs.append(pd.Series(tmp,index=cols))

		recs_multi = pd.DataFrame(recs,columns=cols).set_index('Region')
		recs_multi.Volume.replace('[a-zA-Z]',value='',regex=True, inplace=True)

		#Sum the Volumes
		sumvol = pd.to_numeric(recs_multi.Volume,errors='coerce').sum()

		#Check if directory exists
		d = datetime.datetime.strptime(flow_date, '%m-%d').date()
		f = checkdir(inputs, save_dir, 'gas-daily-directory', d)+'gas-daily-{}.xlsx'.format(flow_date)

		#Save to Excel
		if not os.path.isfile(f):
			saveExcel(recs_multi, f, sr=5, use_index=True)
			appendExcel(hdrawf, f, use_index=True)
			appendExcel(hdf, f, sc=3, use_index=True)
			fn='gas-daily-{}.xlsx'.format(flow_date)
		else:
			#If it's a holiday, then the flow date will be the same as the current day's date. Use the next day's date in the file name
			tomorrow = pd.to_datetime(str(np.busday_offset(today, 1)))
			fn='gas-daily-{}.xlsx'.format(tomorrow.strftime('%m-%d'))
			f_append = checkdir(inputs, save_dir, 'gas-daily-directory', d)+fn
			saveExcel(recs_multi, f_append, sr=5, use_index=True)
			appendExcel(hdrawf, f_append, use_index=True)
			appendExcel(hdf, f_append, sc=3, use_index=True)
			
		#Log on successful download
		logs(inputs, log_dir, 'Download', 'Gas Daily', fn)

	except Exception as e:
		#Email and log errors, if any
		mail(inputs, 'Gas Daily Data Download Error', logs(inputs, log_dir, 'Error', 'Gas Daily', e))

def megawatt_daily(inputs,save_dir,log_dir):
	try:
		#Get Megawatt Daily Table from Platts Login Script
		mw_table = platts_login(inputs, source='meg')

		#Get Column Names
		cols = [mw_table.find('tr', class_="COLUMNHEAD").find_all('td')[colname].get_text() for colname in range(2,len(mw_table.find('tr', class_="COLUMNHEAD").find_all('td')))]

		#Get the Date and Units
		title = mw_table.find('td',class_='TITLE_1').get_text()
		parenth = title.find('(')
		unit = title[parenth:]
		st = title.find('delivery')
		report_date = title[st+8:parenth].strip() + ' ' + today.strftime('%Y')
		report_date_decimal = datetime.datetime.strptime(title[st+8:parenth].strip(),'%b %d').strftime('%m-%d')

		#Get Region (Southwest, etc)
		regions = [mw_table.find_all('td', class_='COLUMNHEAD_2')[ii].get_text().replace('*','').strip() for ii in range(0,len(mw_table.find_all('td', class_='COLUMNHEAD_2')))]
		#Get time Period (On Peak, Off Peak)
		periods = [mw_table.find_all('tr',class_='UNDERLINE')[ii].get_text().replace('*','').strip() for ii in range(0,2)]
		#Get the products
		products = [mw_table.find_all('td',class_='ROW_2')[ii].get_text() for ii in range(0,len(mw_table.find_all('td',class_='ROW_2')))]
		#Get the product IDs
		ids = [mw_table.find_all('td',class_='ROW_1')[ii].get_text() for ii in range(0,len(mw_table.find_all('td',class_='ROW_1')))]

		#Make a Periods Vector
		dupes = [idx for idx, item in enumerate(products) if item in products[:idx]]
		Period = [False]*(max(dupes)+1)
		for step in dupes:
			Period[step] = periods[1]
		for step in range(0,len(Period)):
			if Period[step] == False:
				Period[step] = periods[0]

		#Make a Regions Vector
		firsts = [idx for idx, item in enumerate(products) if item not in products[:idx]]
		seconds = [idx for idx, item in enumerate(products) if item in products[:idx]]
		lower_gap = [idx for idx in firsts if idx in firsts[:idx]]
		Region = [False]*(max(seconds)+1)
		for step in range(firsts[0],seconds[0]):
			Region[step] = regions[0]   
		for step in range(firsts[-1],seconds[-1]+1):
			Region[step] = regions[1]
		for step in range(seconds[0],lower_gap[0]):
			Region[step] = regions[0]
		for step in lower_gap:
			Region[step] = regions[1]

		#Compile into a DataFrame
		vals = []

		#Only ROWS 3-8 Contain the table data
		for step in range(3,9):
			vals.append([mw_table.find_all('td',class_='ROW_{}'.format(step))[ii].get_text() \
						for ii in range(0,len(mw_table.find_all('td',class_='ROW_{}'.format(step))))])
		mwf = pd.DataFrame(dict(zip(cols, vals)))
		mwf['Period'] = Period
		mwf['Region'] = Region
		mwf['Product'] = products
		mwf['ID'] = ids

		#Make a Multiindex
		cols.insert(0,'ID')
		mwf_multi = pd.DataFrame(mwf.set_index(['Region','Period','Product']),columns=cols)

		#Save Header Into A DataFrame
		hdf = pd.DataFrame(pd.Series(title))
		hdf = hdf.append(pd.DataFrame([title[0:st+8],report_date,unit]).T).fillna('').reset_index(drop=True)

		#Check if directory exists
		d = datetime.datetime.strptime(report_date_decimal,'%m-%d').date()
		fn='megawatt-daily-{}.xlsx'.format(report_date_decimal)
		f = checkdir(inputs, save_dir, 'megawatt-daily-directory', d)+fn

		#Save to Excel
		saveExcel(mwf_multi, f, sr=3, use_index=True)
		appendExcel(hdf, f)

		#Log on successful download
		logs(inputs, log_dir, 'Download', 'Megawatt Daily', fn)

	except Exception as e:
		#Email and log errors, if any
		mail(inputs, 'Megawatt Daily Data Download Error', logs(inputs, log_dir, 'Error', 'Megawatt Daily', e))

def ice(inputs,save_dir,log_dir,icedate):
	#Formatted date for POST request
	#icedate = yesterday.strftime('%m/%d/%Y')
	yesterday = icedate
	
	#POST URL
	#url = inputs['ice-url']

	#Loop through all the products in the settings file
	for contract in inputs['ice-exchange-contract'].split(','):
		try: 
			#Header information
			#headers = {'User-Agent': inputs['user-agent'], 'Referer':inputs['ice-referer'], 'Host':inputs['ice-host'],
			#       'Origin':inputs['ice-origin']}
			#Form Data
			#payload = {'generateReport':'','exchangeCode':inputs['ice-exchange-code'],'exchangeCodeAndContract':contract.strip(),
			#           'optionRequest':'false','selectedDate':icedate,'submit':'Download'}

			#Start a session
			#session = requests.Session()
			#site = session.get(url)

			#Save the cookies
			#cookies = dict(site.cookies)

			#Add the cookies to the form data dictionary
			#payload.update(cookies)

			#Make the POST request
			#ice = session.post(url, headers=headers, data=payload, cookies=cookies)
			
			#Save the PDF
			#icedir = checkdir(inputs, save_dir, 'ice-directory', yesterday)
			icedir = save_dir+inputs['ice-directory']+'/{}/{}/'.format(yesterday.year,yesterday.strftime('%B %Y'))
			abbr = contract.split('|')[-1]
			fn = '{}_{}.pdf'.format(abbr, yesterday.strftime('%Y_%m_%d'))
			#print(icedir+fn)
			#fn = '{}_{}.pdf'.format(abbr, yesterday.strftime('%m-%d-%Y'))
			#fd = open(icedir+fn, 'wb')
			#fd.write(ice.content)
			#fd.close()
			
			#Specify a coordinates list to save XML row coordinates to
			coords = []

			#Create a dataframe to store the XML data in
			df = pd.DataFrame(columns=['Commodity Name','Contract Month','Settle Price','Settle Change','Total Volume','OI', 'Volume Change', 'EFP', 'EFS', 'Block Volume', 'Spread Volume'])

			#Convert the PDF to XML and parse it
			subprocess.run(['pdftohtml', '-xml', icedir+fn])
			soup = BeautifulSoup(open(icedir+fn.replace('.pdf','.xml')),"lxml")                         
																	  
			#Get the page numbers from the XML document
			pages = [number['number'] for number in soup.find_all('page')]

			#Get full Commodity name from XML document
			commodity = soup.find('text', string=re.compile('{}-.'.format(abbr))).get_text()

			#Loop through each page and save the table to the dataframe
			#Specify a failsafe variable, so if there is data missing, omit it from the Excel file
			failsafe_var = 0
			for page in pages:
				coords = [tags['top'] for tags in soup.find('page', attrs={'number':page}).find_all('text', string='{}'.format(abbr))]
				for idx,numb in enumerate(coords):
					row_len = len([item.get_text() for item in soup.find('page', attrs={'number':page}).find_all('text', attrs={"top":coords[idx]})])
					if row_len not in (11, 15):
						failsafe_var += 1
					else:
						row_beg = [item.get_text() for item in soup.find('page', attrs={'number':page}).find_all('text', attrs={"top":coords[idx]})[0:2]]
						row_end = [item.get_text() for item in soup.find('page', attrs={'number':page}).find_all('text', attrs={"top":coords[idx]})[-9:]]
						row = row_beg + row_end
						df = df.append(pd.DataFrame([row],columns=df.columns),ignore_index=True)
						
			#Change the date format in the Contract Month column
			try:
				df['Contract Month'] = [datetime.datetime.strptime(condate,'%b-%d-%y').strftime('%m/%d/%Y') for condate in df['Contract Month']]
			except:
				df['Contract Month'] = [datetime.datetime.strptime(condate,'%b%y').strftime('%m/%d/%Y') for condate in df['Contract Month']]

			#Rename the Commodity column to the more specific commodity variable saved above
			df['Commodity Name'] = [name.replace('{}'.format(abbr), '{}'.format(commodity)) for name in df['Commodity Name']]

			#Add a column with the sale date
			df.insert(1,'Sale Date',[icedate.strftime('%m/%d/%Y')]*len(df))

			#Save to Excel
			saveExcel(df, icedir+fn.replace('.pdf','.xlsx'), 0)

			#If the data is "correct" then save a log message
			if failsafe_var == 0:    
				logs(inputs, log_dir, 'Download', 'ICE', fn.replace('.pdf','.xlsx'))

			#The data has been flagged as possibly incorrectly formatted due to missing data. Rename the Excel file, so the data doesn't get added to Access. Then send an email alert.
			else:
				os.rename(icedir+fn.replace('.pdf','.xlsx'),icedir+fn.replace('.pdf','_FORMAT-ERROR.xlsx'))
				mail(inputs, 'Missing ICE Data Error on {}'.format(today.strftime('%d %B %Y')), 'Missing {} ICE data detected on {}'.format(abbr,today.strftime('%d %B %Y')))
			
			#Sleep to prevent R/W buffer errors
			#time.sleep(2)

		except Exception as e:
			print(e)
			#Email and log errors, if any
			#mail(inputs, 'ICE Data Download Error', logs(inputs, log_dir, 'Error', 'ICE', e))

def nodal_exchange(inputs,save_dir,log_dir):
	try:
		
		#Nodal exchange url
		url = inputs['nodal-market-data-url']
		#url = inputs['nodal-exchange-url']

		#Search the page for the URL to the end of day (EOD) futures PDF
		urlpdf = requests.compat.urljoin(url,BeautifulSoup(requests.get(url).text, 'lxml').find(href=re.compile('Futures',re.IGNORECASE))['href'])

		#Yank out filename from URL
		nodalfn = os.path.basename(urlpdf)

		#Save the extension as a case insensitive regular expression
		ext = re.compile(re.escape('pdf'), re.IGNORECASE)
		
		#Download and save PDF
		site = requests.get(urlpdf)
		with open(nodalfn, 'wb') as outfile:
			outfile.write(site.content)

		#Convert the PDF to XML for parsing
		subprocess.run(['pdftohtml', '-xml', nodalfn])

		#Parse the XML
		crap = xml.etree.ElementTree.parse(ext.sub('xml',nodalfn))
		root = crap.getroot()

		#List for storing data
		rows = []
		
		#Get a list of top coordinates
		coords = list(OrderedDict.fromkeys([item.attrib['top'] for item in root.findall('*/text[@font="2"]')]))
		
		#Loop through coordinates and save data
		for coord in coords:
			temp = [item.text for item in root.findall('*/text[@font="2"][@top="{}"]'.format(coord))]
			#If there are 9 entities in the row (or modulus thereof) then the row has data we are interested in
			if len(temp)%9 == 0:
				rows.append([temp[x:x+9] for x in range(0, len(temp), 9)])

		#Save the output to a dataframe (flatten i.e. denest the list of data so that every row is its own single list)
		df = pd.DataFrame([item for sublist in rows for item in sublist],columns=['Contract Code', 'Expiry', 'Settlement Price', 'Price Change', 'Open Interest', 'OI Change', 'Total Volume', 'EFRP Volume', 'Block Volume'])

		#Get the date from the PDF
		nodal_date = datetime.datetime.strptime(root.findall('*/text[@font="0"]')[-1].text.strip(),'%d-%b-%y')

		#Convert the Expiry row to datetime for sorting purposes
		df['Expiry'] = pd.to_datetime(df['Expiry'])

		#Add a column for the date of the data
		df.insert(1,'Sale Date',[nodal_date.strftime('%m/%d/%Y')]*len(df))

		#Save to Excel
		saveExcel(df.sort_values(by=['Contract Code','Expiry']), ext.sub('xlsx',nodalfn), 0)
		
		#Copy the PDF XML and XLS files and append the date to each
		nodaldir = checkdir(inputs, save_dir, 'nodal-exchange-directory', thedate=nodal_date)
		copyfile(nodalfn,nodaldir+os.path.splitext(nodalfn)[0]+'_'+nodal_date.strftime('%d-%b-%y')+os.path.splitext(nodalfn)[-1])
		copyfile(ext.sub('xml',nodalfn),nodaldir+os.path.splitext(nodalfn)[0]+'_'+nodal_date.strftime('%d-%b-%y')+os.path.splitext(ext.sub('xml',nodalfn))[-1])
		copyfile(ext.sub('xlsx',nodalfn),nodaldir+os.path.splitext(nodalfn)[0]+'_'+nodal_date.strftime('%d-%b-%y')+os.path.splitext(ext.sub('xlsx',nodalfn))[-1])

		#Delete the temp files on local machine
		os.remove(nodalfn)
		os.remove(ext.sub('xml',nodalfn))
		os.remove(ext.sub('xlsx',nodalfn))
		
		#Record log entry
		logs(inputs, log_dir, 'Download', 'Nodal Exchange', os.path.splitext(nodalfn)[0]+'_'+nodal_date.strftime('%d-%b-%y')+os.path.splitext(ext.sub('xlsx',nodalfn))[-1])
		
		#Download the Contract Specifications File
		conurl = inputs['nodal-contract-info-url']
		consoup = BeautifulSoup(requests.get(conurl).text ,'lxml')

		links = consoup.find_all('a')

		for link in links:
			if link.find(text=re.compile('Contract Specification',re.IGNORECASE)):
				thelink = link
				break
		conpdfurl = requests.compat.urljoin(conurl, thelink['href'])
		site = requests.get(conpdfurl)
		with open(save_dir+inputs['nodal-exchange-directory']+os.path.basename(conpdfurl), 'wb') as outfile:
			outfile.write(site.content)

		#Download Contract & Node List
		for link in links:
			if link.find(text=re.compile('Contract & Node List',re.IGNORECASE)):
				thelink = link
				break
		conpdfurl = requests.compat.urljoin(conurl, thelink['href'])
		site = requests.get(conpdfurl)
		with open(save_dir+inputs['nodal-exchange-directory']+os.path.basename(conpdfurl), 'wb') as outfile:
			outfile.write(site.content)

	except Exception as e:
		#Email and log errors, if any
		mail(inputs, 'Nodal Exchange Download Error', logs(inputs, log_dir, 'Error', 'Nodal Exchange', e))

if __name__ == "__main__": main()
