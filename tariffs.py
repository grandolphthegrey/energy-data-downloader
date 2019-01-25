#!/usr/bin/env python3

import numpy as np
import pandas as pd
import requests
import urllib.request
import re
import os
import datetime
import subprocess
import smtplib
import time

from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText


def main():

	#Open the "settings.txt" file to define settings, URLS, directories, etc for downloading all the data.
	settings={}
	with open('/Users/Dev/settings.txt') as f:
		for line in f:
			if '#' not in line and len(line)>1:
				(key,val)=line.strip().split(':',1)
				settings[key]=val
			
	#Mount Drive
	mount_drive(settings)

	#Directories
	#root_dir = settings['sharepoint-lib']+'/Tariffs-gas_electric/'
	root_dir = settings['sharepoint-lib']+settings['tariff-dir']
	utilities = 'PG&E Electric {d}/@PG&E Gas {d}/@SDG&E Electric {d}/@SDG&E Gas {d}/@SCE {d}/@SoCal Gas {d}/'.format(d = datetime.date.today().strftime('%Y-%m')).split('@')

	tariffs = ['Preliminary Statements/', 'Rates-Schedules/', 'Rules/']	

	#Loop through and make the directories
	for utility in utilities:
		for tariff in tariffs:
			if not os.path.exists(root_dir+utility+tariff):
				os.makedirs(root_dir+utility+tariff)

	pge(settings, root_dir, utilities, tariffs)
	sdge(settings, root_dir, utilities, tariffs)
	sce(settings, root_dir, utilities, tariffs)
	socal(settings, root_dir, utilities, tariffs)


	#Unmount the network drive
	unmount_drive(settings)



def pge(settings, root_dir, utilities, tariffs):
	try:
		#########################			
		#########PG&E #########
		#########################
		webpage = settings['pge-webpage']
		get_content = settings['pge-get-content']
		index_command = settings['pge-index-command']

		soup =  BeautifulSoup(requests.get(webpage+index_command).text,'lxml')

		#########
		#Electric
		#########
		#Preliminary Statements
		elec_prelim_url = urllib.parse.urljoin(webpage,get_content+soup.find(text=search('Electric Preliminary Statements')).find_next('div',class_='tariff-contents contents')['data'])
		download_pdf(BeautifulSoup(requests.get(elec_prelim_url).text, 'lxml').tbody.find_all('tr'),webpage,root_dir+utilities[0]+tariffs[0])
		#Rate Schedules
		elec_rates_url = urllib.parse.urljoin(webpage,get_content+soup.find(text=search('Electric Rate Schedules')).find_next('div',class_='tariff-contents contents')['data'])
		download_pdf(BeautifulSoup(requests.get(elec_rates_url).text, 'lxml').tbody.find_all('tr'),webpage,root_dir+utilities[0]+tariffs[1])
	    #Rules
		elec_rules_url = urllib.parse.urljoin(webpage,get_content+soup.find(text=search('Electric Rules')).find_next('div',class_='tariff-contents contents')['data'])
		download_pdf(BeautifulSoup(requests.get(elec_rules_url).text, 'lxml').tbody.find_all('tr'),webpage,root_dir+utilities[0]+tariffs[2])

		####
		#Gas
		####
		#Preliminary Statements
		gas_prelim_url = urllib.parse.urljoin(webpage,get_content+soup.find(text=search('Gas Preliminary Statements')).find_next('div',class_='tariff-contents contents')['data'])
		download_pdf(BeautifulSoup(requests.get(gas_prelim_url).text, 'lxml').tbody.find_all('tr'),webpage,root_dir+utilities[1]+tariffs[0])
		#Rate Schedules
		gas_rates_url = urllib.parse.urljoin(webpage,get_content+soup.find(text=search('Gas Rate Schedules')).find_next('div',class_='tariff-contents contents')['data'])
		download_pdf(BeautifulSoup(requests.get(gas_rates_url).text, 'lxml').tbody.find_all('tr'),webpage,root_dir+utilities[1]+tariffs[1])
		#Rules
		gas_rules_url = urllib.parse.urljoin(webpage,get_content+soup.find(text=search('Gas Rules')).find_next('div',class_='tariff-contents contents')['data'])
		download_pdf(BeautifulSoup(requests.get(gas_rules_url).text, 'lxml').tbody.find_all('tr'),webpage,root_dir+utilities[1]+tariffs[2])

	except Exception as e:
		mail(settings, 'PG&E Tariff Download Error: {}'.format(datetime.date.today().strftime('%A %d %B %Y')), 'Error downloading tariffs on {}: {}'.format(datetime.date.today().strftime('%A %d %B %Y'), e))
	
def sdge(settings, root_dir, utilities, tariffs):
	try:	
		#########################
		########SDG&E #########
		#########################
		webpage = requests.get(settings['sdge-webpage'])
		soup = BeautifulSoup(webpage.text, 'lxml')
		
		#Electric
		elec_section = soup.find(string=search('Electric Tariff Book')).parent 
		#Preliminary Statements
		elec_prelim_url = elec_section.find_next('a',text=search('Preliminary Statement'))['href']
		download_pdf(BeautifulSoup(requests.get(elec_prelim_url).text, 'lxml').find_all('a'),elec_prelim_url,root_dir+utilities[2]+tariffs[0])
		#Rate Schedules
		elec_rates = elec_section.find_next(text=search('Schedule Of Rates')).parent.find_all('a')
		for rate in elec_rates:
		    download_pdf(BeautifulSoup(requests.get(rate['href']).text,'lxml').find_all('a'),rate['href'],root_dir+utilities[2]+tariffs[1])
		#Rules
		elec_rules_url = elec_section.find_next(text=search('Electric Rules')).parent['href']
		download_pdf(BeautifulSoup(requests.get(elec_rules_url).text, 'lxml').find_all('a'),elec_rules_url,root_dir+utilities[2]+tariffs[2])
		####
		#Gas
		####
		gas_section = soup.find(string=search('Gas Tariff Book')).parent
		#Preliminary Statements
		gas_prelim_url = gas_section.find_next('a', text=search('Preliminary Statement'))['href']
		download_pdf(BeautifulSoup(requests.get(gas_prelim_url).text, 'lxml').find_all('a'),gas_prelim_url,root_dir+utilities[3]+tariffs[0])
		#Rate Schedules
		gas_rates = gas_section.find_next(text=search('Schedule Of Rates')).parent.find_all('a')
		for rate in gas_rates:
		    download_pdf(BeautifulSoup(requests.get(rate['href']).text,'lxml').find_all('a'),rate['href'],root_dir+utilities[3]+tariffs[1])
		#Rules
		gas_rules_url = gas_section.find_next(text=search('Gas Rules')).parent['href']
		download_pdf(BeautifulSoup(requests.get(gas_rules_url).text, 'lxml').find_all('a'),gas_rules_url,root_dir+utilities[3]+tariffs[2])
	
	except Exception as e:
		mail(settings, 'SDG&E Tariff Download Error: {}'.format(datetime.date.today().strftime('%A %d %B %Y')), 'Error downloading tariffs on {}: {}'.format(datetime.date.today().strftime('%A %d %B %Y'), e))
	
def sce(settings, root_dir, utilities, tariffs):
	try:	
		#########################
		#########SCE ##########
		#########################
		webpage = settings['sce-webpage']
		session = requests.Session()
		limit = 5
		
		#Preliminary Statements
		prelim_url = settings['sce-prelim-webpage'] 
		for i in range (0,limit):
			#site = session.get(prelim_url)
			#prelim_resp = session.get(webpage+BeautifulSoup(site.content,'lxml').find('iframe')['src'])
			prelim_resp = session.get(prelim_url)
			if prelim_resp.status_code == 200:
				break
			else:
				pass
		#prelim_links = BeautifulSoup(prelim_resp.content, 'lxml').find_all('a')
		prelim_links = BeautifulSoup(prelim_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#prelim_links = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content,'lxml').find('iframe')['src']).content, 'lxml').find_all('a')
		download_pdf(prelim_links, webpage, root_dir+utilities[4]+tariffs[0])
		'''
			prelim_url = settings['sce-prelim-webpage']
			#prelim_url = 'https://www.sce.com/wps/portal/home/regulatory/tariff-books/preliminary-statements'
			prelim_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage,BeautifulSoup(requests.get(prelim_url).text,'lxml').iframe['src'])).text,'lxml').find_all('a')
			#download_pdf(prelim_links, webpage, root_dir+utilities[4]+tariffs[0])
		'''
		#Rate Schedules
		rate_url = settings['sce-rate-webpage']
		#Rate Schedules - Residential
		for i in range (0,limit):
			site = session.get(rate_url)
			res_url =  BeautifulSoup(site.content, 'lxml').find('a',text=search('Residential Rates'))['href']
			#res_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=search('Residential Rates'))['href']).content, 'lxml').find('iframe')['src']
			res_resp = session.get(webpage+res_url)
			if res_resp.status_code == 200:
				break
			else:
				pass
		res_links = BeautifulSoup(res_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#res_links = BeautifulSoup(res_resp.content, 'lxml').find_all('a')
		#res_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=re.compile('Residential Rates'))['href']).content, 'lxml').find('iframe')['src']
		#res_links = BeautifulSoup(session.get(webpage+res_url).content,'lxml').find_all('a')
		download_pdf(res_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
			res_page = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(rate_url).text, 'lxml').find('a', text=search('Residential Rates'))['href'])
			res_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(res_page).text, 'lxml').iframe['src'])).text, 'lxml').find_all('a')
			download_pdf(res_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
		#Rate Schedules - General Service
		for i in range (0,limit):
			site = session.get(rate_url)
			gen_url = BeautifulSoup(site.content, 'lxml').find('a',text=search('General Service/Industrial Rates'))['href']
			#gen_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=search('General Service/Industrial Rates'))['href']).content, 'lxml').find('iframe')['src']
			gen_resp = session.get(webpage+gen_url)
			if gen_resp.status_code == 200:
				break
			else:
				pass
		gen_links = BeautifulSoup(gen_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#gen_links = BeautifulSoup(gen_resp.content, 'lxml').find_all('a')
		#gen_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=re.compile('General Service/Industrial Rates'))['href']).content, 'lxml').find('iframe')['src']
		#gen_links = BeautifulSoup(session.get(webpage+gen_url).content,'lxml').find_all('a')
		download_pdf(gen_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
			gen_page = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(rate_url).text, 'lxml').find('a', text=search('General Service/Industrial Rates'))['href'])
			gen_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(gen_page).text, 'lxml').iframe['src'])).text, 'lxml').find_all('a')
			#download_pdf(gen_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
		#Rate Schedules - Agriculture
		for i in range (0,limit):
			site = session.get(rate_url)
			ag_url = BeautifulSoup(site.content, 'lxml').find('a',text=search('Agricultural and Pumping Rates'))['href']
			#ag_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=search('Agricultural and Pumping Rates'))['href']).content, 'lxml').find('iframe')['src']
			ag_resp = session.get(webpage+ag_url)
			if ag_resp.status_code == 200:
				break
			else:
				pass
		ag_links = BeautifulSoup(ag_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#ag_links = BeautifulSoup(ag_resp.content, 'lxml').find_all('a')
		#ag_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=re.compile('Agricultural and Pumping Rates'))['href']).content, 'lxml').find('iframe')['src']
		#ag_links = BeautifulSoup(session.get(webpage+res_url).content,'lxml').find_all('a')
		download_pdf(ag_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
			ag_page = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(rate_url, verify=False).text, 'lxml').find('a', text=search('Agricultural and Pumping Rates'))['href'])
			ag_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(ag_page, verify=False).text, 'lxml').iframe['src']),verify=False).text, 'lxml').find_all('a')
			#ag_links = BeautifulSoup(requests.get(BeautifulSoup(requests.get(ag_page), 'lxml').iframe['src']).text, 'lxml').find_all('a')
			#download_pdf(ag_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
		#Rate Schedules - Street Lighting
		for i in range (0,limit):
			site = session.get(rate_url)
			sl_url = BeautifulSoup(site.content, 'lxml').find('a',text=search('Street and Area Lighting/Traffic Control Rates'))['href']
			#sl_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=search('Street and Area Lighting/Traffic Control Rates'))['href']).content, 'lxml').find('iframe')['src']
			sl_resp = session.get(webpage+sl_url)
			if sl_resp.status_code == 200:
				break
			else:
				pass
		sl_links = BeautifulSoup(sl_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#sl_links = BeautifulSoup(sl_resp.content, 'lxml').find_all('a')
		#sl_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=re.compile('Street and Area Lighting/Traffic Control Rates'))['href']).content, 'lxml').find('iframe')['src']
		#sl_links = BeautifulSoup(session.get(webpage+res_url).content,'lxml').find_all('a')
		download_pdf(sl_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
			lite_page = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(rate_url, verify=False).text, 'lxml').find('a', text=search('Street and Area Lighting/Traffic Control Rates'))['href'])
			lite_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(lite_page, verify=False).text, 'lxml').iframe['src']), verify=False).text, 'lxml').find_all('a')
			#download_pdf(lite_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
		#Rate Schedules - Other Rates
		for i in range (0,limit):
			site = session.get(rate_url)
			oth_url = BeautifulSoup(site.content, 'lxml').find('a',text=search('Other Rates'))['href']
			#oth_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=re.compile('Other Rates'))['href']).content, 'lxml').find('iframe')['src']
			oth_resp = session.get(webpage+oth_url)
			if oth_resp.status_code == 200:
				break
			else:
				pass
		oth_links = BeautifulSoup(oth_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#oth_links = BeautifulSoup(oth_resp.content, 'lxml').find_all('a')
		#oth_url = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content, 'lxml').find('a',text=re.compile('Other Rates'))['href']).content, 'lxml').find('iframe')['src']
		#oth_links = BeautifulSoup(session.get(webpage+res_url).content,'lxml').find_all('a')
		download_pdf(oth_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
			oth_page = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(rate_url, verify=False).text, 'lxml').find('a', text=search('Other Rates'))['href'])
			oth_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(oth_page, verify=False).text, 'lxml').iframe['src']), verify=False).text, 'lxml').find_all('a')
			#download_pdf(oth_links, webpage, root_dir+utilities[4]+tariffs[1])
		'''
		#Rules
		rules_url = settings['sce-rules-webpage']
		site = session.get(rules_url)
		for i in range (0,limit):
			site = session.get(rules_url)
			#rules_resp = session.get(webpage+BeautifulSoup(site.content,'lxml').find('iframe')['src'])
			rules_resp = site
			if rules_resp.status_code == 200:
				break
			else:
				pass
		rules_links = BeautifulSoup(rules_resp.content, 'lxml').find(attrs={'class':'sce-hr-divider'}).find_all('a')
		#rules_links = BeautifulSoup(rules_resp.content, 'lxml').find_all('a')
		#rules_links = BeautifulSoup(session.get(webpage+BeautifulSoup(site.content,'lxml').find('iframe')['src']).content, 'lxml').find_all('a')
		download_pdf(rules_links, webpage, root_dir+utilities[4]+tariffs[2])
		'''
			rules_url = settings['sce-rules-webpage']
			#rules_url = 'https://www.sce.com/wps/portal/home/regulatory/tariff-books/rules'
			rules_links = BeautifulSoup(requests.get(urllib.parse.urljoin(webpage,BeautifulSoup(requests.get(rules_url, verify=False, headers=headers).text,'lxml').iframe['src']), verify=False, headers=headers).text,'lxml').find_all('a')
			print(rules_links)
			download_pdf(rules_links, webpage, root_dir+utilities[4]+tariffs[2])
		'''
	except Exception as e:
		#mail(settings, 'SCE Tariff Download Error: {}'.format(datetime.date.today().strftime('%A %d %B %Y')), 'Error downloading tariffs on {}: {}'.format(datetime.date.today().strftime('%A %d %B %Y'), e))
		print(e)

def socal(settings, root_dir, utilities, tariffs):
	try:
		#########################
		######SoCal Gas #######
		#########################
		webpage = settings['socal-webpage']
		tariff_page = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(webpage).text, 'lxml').find('a', text=search('Tariffs -- Gas Rate Schedules and Associated Rules'))['href'])

		#Preliminary Statements
		prelim_url = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(tariff_page).text, 'lxml').find('a', text=search('Preliminary Statement'))['href'])
		prelim_links = BeautifulSoup(requests.get(prelim_url).text, 'lxml').find('i').find_next('ul').find_all('a')
		download_pdf(prelim_links, webpage+'/tariffs/', root_dir+utilities[5]+tariffs[0])
		#Rate Schedules
		rate_url = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(tariff_page).text, 'lxml').find('a', text=search('Rate Schedules'))['href'])
		rate_links = BeautifulSoup(requests.get(rate_url).text, 'lxml').find('i').find_next('ul').find_all('a')
		download_pdf(rate_links, webpage+'/tariffs/', root_dir+utilities[5]+tariffs[1])
		#Rules
		rules_url = urllib.parse.urljoin(webpage, BeautifulSoup(requests.get(tariff_page).text, 'lxml').find('a', text=search('Rules'))['href'])
		rule_links = BeautifulSoup(requests.get(rules_url).text, 'lxml').find('i').find_next('ul').find_all('a')
		download_pdf(rule_links, webpage+'/tariffs/', root_dir+utilities[5]+tariffs[2])

	except Exception as e:
		mail(settings, 'SoCal Gas Tariff Download Error: {}'.format(datetime.date.today().strftime('%A %d %B %Y')), 'Error downloading tariffs on {}: {}'.format(datetime.date.today().strftime('%A %d %B %Y'), e))	
		
#Search Function
def search(name):
    return re.compile('^'+name,re.IGNORECASE)

def download_pdf_test(link_list, root_url, save_dir='{}/Desktop/'.format(os.path.expanduser('~'))):
	for item in link_list:
		print(item)    

#Download Function
def download_pdf(link_list, root_url, save_dir='{}/Desktop/'.format(os.path.expanduser('~'))):
	limit = 5
	for item in link_list:
		attempts = 0
		#Check if input is a list of links:
		if item.td == None:
			name = ''.join(item.get_text().strip().splitlines()).replace('/',' ').replace(':','-') + '.pdf'
			#link = urllib.parse.urljoin(root_url, item['href'])
			link = urllib.parse.urljoin(root_url, item['href'].replace(' ','%20'))
			#print(link)
			while attempts < limit:
				try:
					#urllib.request.urlretrieve(link,save_dir+name)
					ping = requests.get(link)
					with open(save_dir+name,'wb') as outfile:
						outfile.write(ping.content)
					#print(name)
					attempts = limit+1
					time.sleep(1)
					break
				except Exception as e:
					attempts += 1
					print(e)
					pass
					
		else:
			name = item.find_all('td')[0].get_text() + ' - ' + ''.join(item.find_all('td')[-1].get_text().strip().splitlines()).replace('/',' ').replace(':','-') + '.pdf'
			#link = urllib.parse.urljoin(root_url, item.a['href'])
			link = urllib.parse.urljoin(root_url, item.a['href'].replace(' ','%20'))
			#print(link)
			while attempts < limit:
				try:
					#urllib.request.urlretrieve(link,save_dir+name)	
					ping = requests.get(link)
					with open(save_dir+name,'wb') as outfile:
						outfile.write(ping.content)
					#print(name)
					attempts = limit+1
					time.sleep(1)
					break
				except:
					attempts += 1
					pass
					
def mount_drive(inputs):
	#Function to mount the network drive
	#Check if sharepoint is mounted. If not, mount it.
	if not os.path.ismount(inputs['sharepoint-lib']) and not os.path.exists(inputs['sharepoint-lib']):
		os.makedirs(inputs['sharepoint-lib'])
		subprocess.call(['mount','-t', 'smbfs', '//'+inputs['server-username']+':'+inputs['server-password']+'@'+inputs['server-address-lib'], inputs['sharepoint-lib']])

	#Check if sharepoint already exists. If it does, then just mount the network drive
	if os.path.exists(inputs['sharepoint-lib']) and not os.path.ismount(inputs['sharepoint-lib']):
		subprocess.call(['mount','-t', 'smbfs', '//'+inputs['server-username']+':'+inputs['server-password']+'@'+inputs['server-address-lib'], inputs['sharepoint-lib']])

def unmount_drive(inputs):
	#Function to unmount network drive
	subprocess.call(['umount', inputs['sharepoint-lib']])

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

if __name__ == "__main__": main()
