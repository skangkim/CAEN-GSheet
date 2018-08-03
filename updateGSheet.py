# API
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# Registry Key
import winreg
# Runtime calculation
from datetime import datetime
import os, time
# Slack bot message
import requests, json
#log
import logging
# command arg
import sys
#gets API and returns the worksheet object
def getAPI():
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	credentials = ServiceAccountCredentials.from_json_keyfile_name('caen-systems-0992fda1910e.json', scope)
	gc = gspread.authorize(credentials)
	wks = gc.open_by_key('13mRgaFcLumZ_Zpyx2YLQf7qo5rSGgHq-61JNBpey9-w').worksheet('RAWDATA')

	return wks

# reads txt file and returns (MAC address, computer model, start, end, runtime, task_sequence_name)
def getInfo():

	CP_name = os.getenv('COMPUTERNAME')
	textfile = CP_name + ".txt" # get the text file name 
	file = open("c:\windows\logs\\"  + textfile, "r")
	line = file.readline()

	Task_sequence = ""
	Computer_Model = ""
	MACAddress = ""
	start = "" #format ex. Tue 07/03/2018 11:02:29.31
	end = ""
	Dual_status = -1

	while line:
		if line.find("Task Sequence: ") != -1:
			Task_sequence = line[15:].split('\n')[0]
			logger.info("retrieved Task_sequence from txt file")
		
		if line.find("Computer Model:") != -1:
			line = file.readline()
			Computer_Model = line.split('\n')[0]
			logger.info("retrieved Computer_Model from txt file")
		
		if line.find("MACAddress") != -1:
			line = file.readline()
			while len(line.strip()) == 0: # read empty lines
				line = file.readline()
			MACAddress = line.split('\n')[0]
			logger.info("retrieved MACAddress from txt file")
		
		if line.find("START") != -1:
			start = line[:(line.find("START")-1)]
			logger.info("found START from txt file")
		
		if line.find("END") != -1:
			end = line[:(line.find("END")-1)]
			logger.info("found END from txt file")
		
		if line.find("Workstation is a Windows") != -1:
			Dual_status = "Windows Only"
			logger.info("Dual boot status = " + Dual_status)
		
		if line.find("Linux deployed") != -1:
			Dual_status = "Windows and Linux"
			logger.info("Dual boot status = " + Dual_status)
		
		if line.find("dual boot Mac") != -1:
			Dual_status = "Windows and OS"
			logger.info("Dual boot status = " + Dual_status)

		line = file.readline()
	file.close()	 

	#Runtime calculation
	start_time = ""
	end_time = ""
	time_diff = ""
	try:
		start_time = datetime.strptime(start, "%a %m/%d/%Y %H:%M:%S.%f")
	except Exception as e:
		logger.info('Start time not in this format: %a %m/%d/%Y %H:%M:%S.%f')

		try: 
			logger.info('Trying to retrieve start time in this format: %m/%d/%Y %H:%M:%S')
			start_time = datetime.strptime(start, "%m/%d/%Y %H:%M:%S")
			logger.info('Successfully retrieved start time')
		except Exception as e:
			logger.error('Failed to get start time: ' + str(e))


	try:
		end_time = datetime.strptime(end, "%a %m/%d/%Y %H:%M:%S.%f")
	except Exception as e:
		logger.info('Start time not in this format: %a %m/%d/%Y %H:%M:%S.%f')
		try: 
			logger.info('Trying to retrieve end time in this format: %m/%d/%Y %H:%M:%S')
			start_time = datetime.strptime(start, "%m/%d/%Y %H:%M:%S")
			logger.info('Successfully retrieved end time')
		except Exception as e:
			logger.error('Failed to get start time: ' + str(e))

	try:
		time_diff = end_time - start_time
		logger.info("Successfully calculated the runtime")
	except Exception as e:
		logger.error('Failed to get the runtime: ' + str(e))
		time_diff = "N/A"

	return (MACAddress, Computer_Model, start, end, str(time_diff), Task_sequence, Dual_status)

	
def getRegistryKey():
    # Open the key and return the handle object.
    hKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\CAEN")

    # Read the value.                      
    product = winreg.QueryValueEx(hKey, "PRODUCT")
    version = winreg.QueryValueEx(hKey, "VERSION")
    logger.info("retrieved product & version Registry Key")

    # Close the handle object.
    winreg.CloseKey(hKey)

    return (product[0], version[0])

def makePOST():
	logger.error("5 minutes have passed. Sending slack message")
	URL = "https://hooks.slack.com/services/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"
	message = "Deployment log to Google Sheet failed after 5 minutes for " + os.getenv('COMPUTERNAME') + "\n"
	message += "Error returned: gspread.exceptions.APIError"
	slack_data = {
		'attachments' : [ 
			{
				'color' : "#ff0000",
				'text' : message
			}
		]
		
	}
	r = requests.post(URL, data = json.dumps(slack_data), headers = {'Content-Type' : 'application/json'})
	logger.error("Sent slack message")


# start logging 
logging.basicConfig(filename = "C:\caen\gsheetupdater.log", level= logging.INFO)
logger = logging.getLogger("my-logger")
logger.info("START TIME: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

wks = getAPI()
logger.info("set Google Sheets API")
Info = getInfo()
RegistryKey = getRegistryKey()

# column format
#	computer name, MAC address, computer model, start, end, runtime, Product, version, task sequence name
CP_name = os.getenv('COMPUTERNAME')

row = [CP_name, Info[0], Info[1], Info[2], Info[3], Info[4], RegistryKey[0], RegistryKey[1], Info[5], Info[6], sys.argv[1]]

# maximum time 
# 5 min 

i = 0
while True and time.perf_counter() <= 300:
	try:
		wks.append_row(row)
	# When it exceeds Google Sheets API Limit
	except gspread.exceptions.APIError:
		logger.info("tried to append row " + i + " times and failed due to gspread.exceptions.APIError.")
		# Wait for 10 seconds 
		time.sleep(10)
		logger.info("waited for 10 seconds")
		continue
	break
 
# Send slack bot msg if time limit is exceeded 
if time.perf_counter() > 300:
	makePOST()
else:
	logger.info("Successfully updated the sheet")

logger.info("END TIME: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
