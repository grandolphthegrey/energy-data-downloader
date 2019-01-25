#!/usr/bin/env python3

import requests
import datetime
import os
import subprocess
import sys
import smtplib
import codecs

from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText

#Today's date
today = datetime.datetime.today()

def main():
	try:
		#Open the "settings.txt" file to define settings, URLS, directories, etc for downloading all the data.
		settings={}
		with open('/Users/Dev/settings.txt') as f:
			for line in f:
				if '#' not in line and len(line)>1:
					(key,val)=line.strip().split(':',1)
					settings[key]=val

		save_dir = settings['server-save']
		log_dir = settings['server-log']

		#Mount the Server
		mount_drive(settings['sharepoint-lib'], settings)
		mount_drive(settings['sharepoint-energy'], settings)

		#Get the filename and save location
		filename = today.strftime('%m').lstrip('0')+'-'+today.strftime('%d').lstrip('0')+'-'+today.strftime('%y')
		dirsave = settings['save-loc']+today.strftime('%Y')+' Daily Calendar/'

		#Start a web session
		session = requests.Session()

		#Save a credentials dictionary
		credentials = {'PubDateFrom':today.strftime('%m/%d/%y'),'PubDateTo':today.strftime('%m/%d/%y'),'SearchButton':'Search'}

		#Create a header dictionary
		header = {'Referer':settings['cpuc-url'],'User-Agent':settings['user-agent']}

		#HTTP GET request
		site = session.get(settings['cpuc-url'])
		cookies = dict(site.cookies)

		#Save the webpage and parse for input cookies
		webpage = BeautifulSoup(site.content,'html.parser')
		hiddenInputs = webpage.find_all(name = 'input', type ='hidden')
		for hidden in hiddenInputs:
			name = hidden['name']
			value = hidden['value']
			credentials[name] = value

		#HTTP POST request and parse response for link to PDF
		r = session.post(settings['cpuc-url'], data=credentials, headers=header, cookies=cookies)
		out = BeautifulSoup(r.content,'html.parser')
		pdf_link = settings['cpuc-root-url'] + out.find('a', text='PDF')['href']

		#Download PDF
		pdf = requests.get(pdf_link)

		#If the folder doesn't exist (i.e. first calendar of the year) then create the folder
		if not os.path.exists(dirsave):
			os.makedirs(dirsave)

		#Save file
		fn = dirsave+filename+'.pdf'
		fd = open(fn, 'wb')
		fd.write(pdf.content)
		fd.close()

		#Log on successful download
		logs(settings, log_dir, 'Download', 'CPUC Daily Calendar',spacers=True)
	
		#Unmount the drive
		unmount_drive(settings['sharepoint-lib'])
		unmount_drive(settings['sharepoint-energy'])
		

	except Exception as e:
		#Email and log errors, if any
		mail(settings, 'CPUC Calendar Download Error', logs(settings, log_dir, 'Error', 'CPUC Calendar Download Error', e))

		#Unmount the drive
		unmount_drive(settings['sharepoint-lib'])
		unmount_drive(settings['sharepoint-energy'])		

def mount_drive(sharepoint, inputs):
	#Function to mount the network drive
	#Check if sharepoint is mounted. If not, mount it.
	if not os.path.ismount(sharepoint) and not os.path.exists(sharepoint):
		os.makedirs(sharepoint)
		subprocess.call(['mount','-t', 'smbfs', '//'+inputs['server-username']+':'+inputs['server-password']+'@'+inputs['server-address-root']+sharepoint.split('/')[-1], sharepoint])

	#Check if sharepoint already exists. If it does, then just mount the network drive
	if os.path.exists(sharepoint) and not os.path.ismount(sharepoint):
		subprocess.call(['mount','-t', 'smbfs', '//'+inputs['server-username']+':'+inputs['server-password']+'@'+inputs['server-address-root']+sharepoint.split('/')[-1], sharepoint])

def unmount_drive(sharepoint):
	#Function to unmount network drive
	if os.path.ismount(sharepoint):
		subprocess.call(['umount', sharepoint])

def logs(inputs, log_dir, logname, logsource, error=None, spacers=False):
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
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} download completed on {}\n'.format(logsource, today.strftime('%a %d %b %Y at %I:%M %p')) + data)
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
		with codecs.open(f, 'w',encoding='utf-8', errors='ignore') as modified: modified.write('{} Error on {}: '.format(logsource, today.strftime('%a %d %b %y at %I:%M %p'))+errmsg+'\n\n' + data)
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
	mailServer.sendmail(gmail_sender, inputs['email-recipient'], msg.as_string())
	mailServer.close()

if __name__ == "__main__": main()
