import pandas as pd
import os
import inspect
from selenium import webdriver
from auto import *
import json,psutil

import datetime 
import dateutil.relativedelta

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import win32com.client as win32
import re
import sys
from langdetect import detect
import os, os.path
import win32com.client
import openpyxl as xl
import socket

def TIME(s2,s1):
	from datetime import datetime
	time = datetime.strptime(s2, '%H:%M') - datetime.strptime(s1, '%H:%M')
	return time
def LogTXT(content): #dir_path + '\\LogTimeRun.txt'
	import logging
	import os, inspect
	from datetime import datetime as dat

	CurDir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
	dir_path = CurDir+'\\report\\LogTimeRun.txt'

	logger = logging.getLogger(__name__)
	logger.setLevel(logging.DEBUG)
	handler = logging.FileHandler(dir_path,encoding = "utf-8")
	handler.setLevel(logging.DEBUG)
	formatter = logging.Formatter('------------------ %(asctime)s >>>'+content+'<<<----------------------', datefmt='%d/%m/%Y %H:%M:%S')
	handler.setFormatter(formatter)
	logger.addHandler(handler)

	timeStart = dat.now() 
	# Running
	timeEnd = dat.now()

	logger.warning(str(timeEnd-timeStart))

	logging.basicConfig(format='------------------ %(asctime)s >>>  %(message)s  <<<----------------------', datefmt='%d/%m/%Y %H:%M:%S')
	LOG_INFO = logging.warning # logging.info ...
	LOG_INFO(content)
def load_web(path_DOWNLOAD):
	chrome_options = webdriver.ChromeOptions()


	appState = {
            "recentDestinations": [
                {
                    "id": "Save as PDF",
                    "origin": "local"
                }
            ],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
        }


	prefs = {
        'printing.print_preview_sticky_settings.appState': json.dumps(appState),
        'savefile.default_directory': path_DOWNLOAD,
        "download.default_directory": path_DOWNLOAD,
        "safebrowsing.enabled": "false", 
        "download.prompt_for_download": "false",
        "directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"],
        'profile.default_content_setting_values.automatic_downloads': 1
    }
	chrome_options.add_experimental_option("prefs", prefs)
	chrome_options.add_argument("---printing")
	chrome_options.add_argument("--disable-extensions")
	chrome_options.add_argument("--disable-javascript")
	chrome_options.add_argument("--no-sandbox")
	chrome_options.add_argument("--disable-gpu")
	chrome_options.add_argument("--incognito")
	chrome_options.add_argument("--start-maximized")
	driver = webdriver.Chrome(executable_path=os.path.abspath(CurDir + "\\tmpl\\chromedriver.exe"), chrome_options=chrome_options)

	return driver

LogTXT('-----START THE PROCESS-----')
LogTXT('IP: '+str(socket.gethostbyname(socket.gethostname()))+'________NAME PC: '+str(platform.node()))
CurDir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))) # Lấy đường dẫn của của file
os.system("taskkill /f /im EXCEL.EXE")
Wait(2)

LogTXT('Get the value five months') 
strYear_now  = str(datetime.date.today() + dateutil.relativedelta.relativedelta(months=-1)).split('-')[0]
strMonth_now = str(datetime.date.today() + dateutil.relativedelta.relativedelta(months=-1)).split('-')[1]

LogTXT('Read the related files in the CONF directory') 
df_list_email = pd.read_excel(open(CurDir +'\\conf\\list_employee.xlsx','rb'), sheet_name = 0, dtype = object, header = 0) #file thông tin email của nhân viên
df_setting =  pd.read_excel(open(CurDir +'\\conf\\setting.xlsx','rb'), sheet_name = 0, dtype = object, header = 0)
df_Tick =  pd.read_excel(open(CurDir +'\\conf\\holiday.xlsx','rb'), sheet_name = 0, dtype = object, header = 0)


# LogTXT('Create folder five months in the output, input') 
RemoveFolder(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now))
RemoveFolder(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now))
Wait(2)
CreateFolder(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now))
CreateFolder(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now))
LogTXT('Copy file JIBANNET tmpl to folder output JIBANNET') 
CopyFile(CurDir +'\\tmpl\\JIBANNET.xlsx',CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)+'.xlsx')

LogTXT('Copy file output tmpl to folder output') 
while True:
	try:
		CopyFile(CurDir +'\\tmpl\\output.xlsm',CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\output.xlsm')
		break
	except:
		pass



LogTXT('Login Jobcan with user admin') 
path_DOWNLOAD = CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)

driver = load_web(path_DOWNLOAD)
driver.get('https://ssl.jobcan.jp/client')


element_ID_Company = driver.find_element_by_id('client_login_id')
driver.execute_script("arguments[0].setAttribute('value', '" + str(df_setting['User_jobcan_ID'][0]) +"')", element_ID_Company)

element_ID_Group_Manager = driver.find_element_by_id('client_manager_login_id')
driver.execute_script("arguments[0].setAttribute('value', '" + str(df_setting['Joncan_Manager'][0]) +"')", element_ID_Group_Manager)

element_PW = driver.find_element_by_id('client_login_password')
driver.execute_script("arguments[0].setAttribute('value', '" + str(df_setting['User_jobcan_PW'][0]) +"')", element_PW)

element_Login = driver.find_element_by_css_selector('body > div > div:nth-child(1) > form > div:nth-child(5) > button') 
driver.execute_script("arguments[0].click()", element_Login)

# LogTXT('Click the holiday list')
# while True:
# 	try:
# 		element_attendance_management = driver.find_element_by_css_selector('#holiday-manage-menu > ul > li:nth-child(2) > dl > dd > ul > li:nth-child(1) > a') 
# 		driver.execute_script("arguments[0].click()", element_attendance_management)
# 		break
# 	except:
# 		pass
# LogTXT('Click the Year now')
# while True:
# 	try:
# 		driver.find_elements_by_name('year')[0].send_keys(str(strYear_now))
# 		break
# 	except:
# 		pass
# LogTXT('Click the Month now')
# while True:
# 	try:
# 		driver.find_elements_by_name('month')[0].send_keys(str(strMonth_now))
# 		break
# 	except:
# 		pass

# LogTXT('Click parameter settings dowload')

# while True:
# 	try:
# 		Wait(1)
# 		element_pick_Year = driver.find_element_by_id('group_id') 
# 		driver.execute_script("arguments[0].click()", element_pick_Year)
# 		driver.find_element_by_css_selector('#group_id > option:nth-child(75)').click()
# 		break
# 	except:
# 		pass
# LogTXT('Click parameter settings dowload')
# while True:
# 	try:
# 		Wait(1)
# 		element_pick_Year = driver.find_element_by_id('my_status_0') 
# 		driver.execute_script("arguments[0].click()", element_pick_Year)
# 		break
# 	except:
# 		pass
# while True:
# 	try:
# 		Wait(1)
# 		element_pick_Year = driver.find_element_by_id('other_status_0') 
# 		driver.execute_script("arguments[0].click()", element_pick_Year)
# 		break
# 	except:
# 		pass
# while True:
# 	try:
# 		Wait(1)
# 		element_pick_Year = driver.find_element_by_id('application_status_0') 
# 		driver.execute_script("arguments[0].click()", element_pick_Year)
# 		break
# 	except:
# 		pass
# LogTXT('Click show')
# while True:
# 	try:
# 		Wait(1)
# 		element_pick_Year = driver.find_element_by_css_selector('#search > table > tbody > tr:nth-child(11) > td > input') 
# 		driver.execute_script("arguments[0].click()", element_pick_Year)
# 		break
# 	except:
# 		pass


# LogTXT('Click button dowload')
# while True:
# 	try:
# 		Wait(1)
# 		element_pick_Year = driver.find_element_by_css_selector('#search-result > div.execute-area > div.button > input:nth-child(3)') 
# 		driver.execute_script("arguments[0].click()", element_pick_Year)
# 		break
# 	except:
# 		pass

LogTXT('Click on the date box') 
# Click on the date box
while True:
	try:
		element_attendance_management = driver.find_element_by_xpath('//*[@id="adit-manage-step"]/a') 
		driver.execute_script("arguments[0].click()", element_attendance_management)
		break
	except:
		pass
while True:
	try:
		element_attendance_management = driver.find_element_by_css_selector('#adit-manage-menu > ul > li:nth-child(1) > dl > dd > ul > li:nth-child(2) > a:nth-child(1)') 
		driver.execute_script("arguments[0].click()", element_attendance_management)
		break
	except:
		pass


LogTXT('Click parameter settings dowload')
while True:
	try:
		Wait(1)
		element_pick_Year = driver.find_element_by_id('setting_id_form') 
		driver.execute_script("arguments[0].click()", element_pick_Year)
		driver.find_element_by_css_selector('#setting_id_form > option:nth-child(3)').click()
		break
	except:
		pass
while True:
	try:
		Wait(1)
		element_attendance_management = driver.find_element_by_id('number_of_afile_several') 
		driver.execute_script("arguments[0].click()", element_attendance_management)
		break
	except:
		pass

LogTXT('Click OK javascrip') 
WebDriverWait(driver, 3).until(EC.alert_is_present(),
                    'Timed out waiting for PA creation ' +
                    'confirmation popup to appear.')

alert = driver.switch_to.alert
alert.accept()

LogTXT('Click parameter settings dowload') 
while True:
	try:
		Wait(1)
		element_pick_Year = driver.find_element_by_id('number_of_afile_select') 
		driver.execute_script("arguments[0].click()", element_pick_Year)
		driver.find_element_by_id('number_of_afile_select').send_keys(str('50'))
		break
	except:
		pass
	
while True:
	try:
		Wait(1)
		element_attendance_management = driver.find_element_by_css_selector('#search > table > tbody > tr:nth-child(6) > th > label > input[type=radio]') 
		driver.execute_script("arguments[0].click()", element_attendance_management)
		break
	except:
		pass


LogTXT('Choose a year')
while True:
	try:
		Wait(1)
		element_pick_Year = driver.find_element_by_id('mfrom[y]') 
		driver.execute_script("arguments[0].click()", element_pick_Year)
		driver.find_element_by_id('mfrom[y]').send_keys(str(strYear_now))
		break
	except:
		pass
LogTXT('Choose a Month')
while True:
	try:
		Wait(1)
		element_pick_Month = driver.find_element_by_id('mfrom[m]') 
		driver.execute_script("arguments[0].click()", element_pick_Month)
		driver.find_element_by_id('mfrom[m]').send_keys(str(strMonth_now))
		break
	except:
		pass
LogTXT('Click parameter settings dowload')
while True:
	try:
		Wait(1)
		element_pick_Year = driver.find_element_by_id('group_id') 
		driver.execute_script("arguments[0].click()", element_pick_Year)
		driver.find_element_by_css_selector('#group_id > option:nth-child(70)').click()
		break
	except:
		pass
LogTXT('Click dowload')
# click dowload
while True:
	try:
		Wait(1)
		element_dowload = driver.find_element_by_css_selector('#download-link > a > span') 
		driver.execute_script("arguments[0].click()", element_dowload)
		break
	except:
		pass
LogTXT('Wait for the input zip file')
intcout = 0
while True:
	Wait(1)
	for filename in os.listdir(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)):
		if '.zip' in filename:
			print(filename)
			intcout = 1
			break
	if intcout == 1:
		break
driver.quit()

LogTXT('Extract the input file after downloading') 
for filename in os.listdir(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)):
	if '.zip' in filename:
		UnzipFolder(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\'+filename)
		Wait(2)
		RemoveFile(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\'+filename)

# LogTXT('Move the input file sheets into the output file') 
# intcout = 0
# while intcout <=2:
# 	for file_input in os.listdir(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)):
# 		Ls_sheet_name = pd.ExcelFile(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\'+file_input).sheet_names
# 		# print(len(Ls_sheet_name))
# 		# print(file_input)
# 		for name_sheet in Ls_sheet_name :
# 			# print(name_sheet)
# 			# print(detect(str(name_sheet).lower()))
# 			# if 'ko' in detect(str(name_sheet).lower()) or 'zh' in detect(str(name_sheet).lower()) or 'ja' in detect(str(name_sheet).lower()):
# 			if 'vi' in detect(str(name_sheet).lower()) or 'tl'  in detect(str(name_sheet).lower()) or 'sk'  in detect(str(name_sheet).lower()) or 'so' in detect(str(name_sheet).lower())  or 'af' in detect(str(name_sheet).lower()) or 'fr' in detect(str(name_sheet).lower()) or 'fi' in detect(str(name_sheet).lower()) or 'es' in detect(str(name_sheet).lower()):
# 				print(detect(str(name_sheet).lower())+"-------------------------")
# 				Path_input = CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\'+ file_input
# 				Path_output = CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\output.xlsm'
# 				MoveSheet(Path_input,Path_output,str(name_sheet),'Sheet1')
# 	intcout += 1

LogTXT('Run macro file collection output.xlsm') 
# Run macro
xlApp = win32com.client.DispatchEx('Excel.Application')
xlApp.DisplayAlerts = False
xlsPath = os.path.expanduser(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\output.xlsm')
wb = xlApp.Workbooks.Open(Filename=xlsPath)
xlApp.Run('GetSheets') # Tên macro
wb.Save()
xlApp.Quit()


LogTXT('Run macro file sort ID output.xlsm') 
# Run macro
xlApp = win32com.client.DispatchEx('Excel.Application')
xlApp.DisplayAlerts = False
xlsPath = os.path.expanduser(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\output.xlsm')
wb = xlApp.Workbooks.Open(Filename=xlsPath)
xlApp.Run('SortWksByCell') # Tên macro
wb.Save()
xlApp.Quit()




LogTXT('Copy File input to output') 
while True:
	try:
		CopyFile(CurDir +'\\input\\'+ str(strYear_now)+str(strMonth_now)+'\\output.xlsm',CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsm')
		break
	except:
		pass

LogTXT('Remove Sheet1 in file output.xlsm') 
try:
	Remove_sheet_EX(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsm','Sheet1')
except:
	pass
LogTXT('convert file output xlsm to output xlsx')
wb = xl.load_workbook(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsm')
wb.save(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsx')
RemoveFile(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsm')


dfJibannet = pd.read_excel(open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)+'.xlsx','rb'), sheet_name = 0, dtype = object, header =1)



	


Wait(2)
Ls_sheet_name = pd.ExcelFile(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsx').sheet_names
for name_sheet in Ls_sheet_name:
	dfOutput = pd.read_excel(open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsx','rb'), sheet_name = str(name_sheet), dtype = object, header =9)
	dfOutput_checkMNV = pd.read_excel(open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\'+str(int(strMonth_now))+'月出勤情報.xlsx','rb'), sheet_name = str(name_sheet), dtype = object, header =0)
	print(name_sheet)
	# print(dfOutput_checkMNV.loc[2][2])
	intID = 0
	for index1, row1 in dfJibannet.iterrows():
		if convert_unsigned(str(row1['ID']).lower().replace(' ','')) == convert_unsigned(str(dfOutput_checkMNV.loc[2][2]).lower().replace(' ','')):
			intID = index1
	dfGetID = pd.read_excel(open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)+'.xlsx','rb'), sheet_name = 0, dtype = object, header =2)
	for indexJB, rowJB in dfGetID.iterrows():
		if str(rowJB[2]).lower() == 'nan':
			intIDJB = indexJB
			break
	if intID == 0:
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")
		wb1.Cells(intIDJB+4,2).Value = str(dfOutput_checkMNV.loc[2][2])
		wb1.Cells(intIDJB+4,3).Value = str(name_sheet)
		wlwb.Close(True)
		xlApp.Quit()

		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\tmpl\\JIBANNET.xlsx') 
		wb1= wlwb.Sheets("社員の勤怠")
		wb1.Cells(intIDJB+4,2).Value = str(dfOutput_checkMNV.loc[2][2])
		wb1.Cells(intIDJB+4,3).Value = str(name_sheet)
		wlwb.Close(True)
		xlApp.Quit()

	dfJibannet = pd.read_excel(open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)+'.xlsx','rb'), sheet_name = 0, dtype = object, header =1)
	for index1, row1 in dfJibannet.iterrows():
		if convert_unsigned(str(row1['ID']).lower().replace(' ','')) == convert_unsigned(str(dfOutput_checkMNV.loc[2][2]).lower().replace(' ','')):
			intID = index1

	intdayoff = 0
	

	#OFF
	intDay = 0
	lsDay = []

	#LATE:
	intDaylate = 0
	lsDaylate = []

	intTick_hours_late = 0
	intTick_minutes_late = 0

	#COME BACK SOON
	intComebacksoon = 0
	lsComebacksoon = []

	intTick_hours_combacksoon = 0
	intTick_minutes_combacksoon = 0

	#OVER TIME
	intovertime = 0
	lsOvertime = []

	intTick_hours = 0 
	intTick_minutes = 0
	
	#OVERTIME HOLIDAYS
	intoverholiday = 0
	lsOverhoidays = []

	intTick_hours_holiday = 0 
	intTick_minutes_holiday = 0

	from datetime import datetime
	for index, row in dfOutput.iterrows():

		
		#OFF
		if '欠勤' in str(row['勤怠状況']):
			lsDay.append(str(row['日付']).split('(')[0])
			intDay += 1
		if '有休' in str(row['勤怠状況']):
			if '04:00' in str(row[14]):
				intdayoff += 0.5
			else:
				intdayoff += 1
		#LATE:
		if '遅刻' in str(row['勤怠状況']):
			if '有休' in str(row['勤怠状況']):
				if  datetime.strptime(str(row['実労働\n時間']), '%H:%M') < datetime.strptime('03:00', '%H:%M'):
					intDaylate = 1
					strTimelate = TIME(str(row['出勤\n時刻']),'13:00')
					intTick_hours_late = intTick_hours_late + int(str(int(str(strTimelate).split(':')[0])))
					intTick_minutes_late = intTick_minutes_late + int(str(int(str(strTimelate).split(':')[1])))
					lsDaylate.append(str(int(str(strTimelate).split(':')[0]))+'h'+str(int(str(strTimelate).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
			else:
				intDaylate = 1
				if str(row['休日区分']).strip() == '祝日':
					strTimelate = TIME(str(row['出勤\n時刻']),'08:00')
					intTick_hours_late = intTick_hours_late + int(str(int(str(strTimelate).split(':')[0])))
					intTick_minutes_late = intTick_minutes_late + int(str(int(str(strTimelate).split(':')[1])))
					lsDaylate.append(str(int(str(strTimelate).split(':')[0]))+'h'+str(int(str(strTimelate).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
				else:
					strTimelate = TIME(str(row['出勤\n時刻']),str(row['シフト\n開始時刻']))
					intTick_hours_late = intTick_hours_late + int(str(int(str(strTimelate).split(':')[0])))
					intTick_minutes_late = intTick_minutes_late + int(str(int(str(strTimelate).split(':')[1])))
					lsDaylate.append(str(int(str(strTimelate).split(':')[0]))+'h'+str(int(str(strTimelate).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
		#COME BACK SOON
		if '早退' in str(row['勤怠状況']):
			
			intComebacksoon = 1
			if  datetime.strptime(str(row['退勤\n時刻']), '%H:%M') < datetime.strptime('12:00', '%H:%M'):
				strTimesoon = TIME(str(row["シフト\n終了時刻"]),str(row["退勤\n時刻"]))
				intTick_hours_combacksoon = intTick_hours_combacksoon + int(str(int(str(strTimesoon).split(':')[0])-1))
				intTick_minutes_combacksoon = intTick_minutes_combacksoon + int(str(int(str(strTimesoon).split(':')[1])))
				lsComebacksoon.append(str(int(str(strTimesoon).split(':')[0])-1)+'h'+str(int(str(strTimesoon).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
			else:
				if datetime.strptime(str(row['退勤\n時刻']), '%H:%M') < datetime.strptime('13:00', '%H:%M'):
					strTimesoon = '04:00:00'
					intTick_hours_combacksoon = intTick_hours_combacksoon + int(str(int(str(strTimesoon).split(':')[0])))
					intTick_minutes_combacksoon = intTick_minutes_combacksoon + int(str(int(str(strTimesoon).split(':')[1])))
					lsComebacksoon.append(str(int(str(strTimesoon).split(':')[0]))+'h'+str(int(str(strTimesoon).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
				else:
					if str(row['休日区分']).strip() == '祝日':
						strTimesoon = TIME('17:00',str(row["退勤\n時刻"]))
						intTick_hours_combacksoon = intTick_hours_combacksoon + int(str(int(str(strTimesoon).split(':')[0])))
						intTick_minutes_combacksoon = intTick_minutes_combacksoon + int(str(int(str(strTimesoon).split(':')[1])))
						lsComebacksoon.append(str(int(str(strTimesoon).split(':')[0]))+'h'+str(int(str(strTimesoon).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
					else:
						strTimesoon = TIME(str(row["シフト\n終了時刻"]),str(row["退勤\n時刻"]))
						intTick_hours_combacksoon = intTick_hours_combacksoon + int(str(int(str(strTimesoon).split(':')[0])))
						intTick_minutes_combacksoon = intTick_minutes_combacksoon + int(str(int(str(strTimesoon).split(':')[1])))
						lsComebacksoon.append(str(int(str(strTimesoon).split(':')[0]))+'h'+str(int(str(strTimesoon).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
		#OVER TIME
		
		if str(row["退勤\n時刻"]).lower() != 'nan' and str(row['休日区分']).strip() != '祝日':
			if str(row["実残業\n時間"]) != '00:00' and str(row["シフト\n終了時刻"]) != '00:00':
				intovertime = 1
				strOvertime = str(row["実残業\n時間"]) + str(':00')
				intTick_hours = intTick_hours + int(str(int(str(strOvertime).split(':')[0])))
				intTick_minutes = intTick_minutes + int(str(int(str(strOvertime).split(':')[1])))
				lsOvertime.append(str(int(str(strOvertime).split(':')[0]))+'h'+str(int(str(strOvertime).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")

		#OVERTIME HOLIDAYS
		if '土' in str(row['日付']).strip() or '日' in str(row['日付']).strip() :
			if str(row['シフト外\n労働時間']) != '00:00':
				intoverholiday = 1
				if  datetime.strptime(str(row["出勤\n時刻"]), '%H:%M') < datetime.strptime('08:00', '%H:%M'):
					a = str(TIME('08:00',str(row["出勤\n時刻"]))).replace(':00','')
				else:
					a = '0:00'
				strOvertimeholiday = TIME(str(row['シフト外\n労働時間']),a)

				intTick_hours_holiday = intTick_hours_holiday + int(str(int(str(strOvertimeholiday).split(':')[0])))
				intTick_minutes_holiday = intTick_minutes_holiday + int(str(int(str(strOvertimeholiday).split(':')[1])))

				lsOverhoidays.append(str(int(str(strOvertimeholiday).split(':')[0]))+'h'+str(int(str(strOvertimeholiday).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")
		for indexholiday, rowholiday in df_Tick.iterrows():
			if str(rowholiday['TICK']) != 'nan':
				if '/' + str(rowholiday['日付']) in str(row['日付']).strip():
					if str(row["実労働\n時間"]) != '00:00':
						intoverholiday = 1
						strOvertimeholiday = TIME(str(row["実労働\n時間"]),'0:00')

						intTick_hours_holiday = intTick_hours_holiday + int(str(int(str(strOvertimeholiday).split(':')[0])))
						intTick_minutes_holiday = intTick_minutes_holiday + int(str(int(str(strOvertimeholiday).split(':')[1])))

						lsOverhoidays.append(str(int(str(strOvertimeholiday).split(':')[0]))+'h'+str(int(str(strOvertimeholiday).split(':')[1]))+"'(" + str(row['日付']).split('(')[0]+")")


	if intdayoff > 0:
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")
		wb1.Cells(intID+ 3,4).Value = str(intdayoff)
		wlwb.Close(True)
		xlApp.Quit()
	if intDay > 0:
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")
		wb1.Cells(intID+ 3,6).Value = str(intDay) +"日 "+ str(lsDay)
		wlwb.Close(True)
		xlApp.Quit()
	if intDaylate == 1:
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")

		if intTick_minutes_late >= 60:
			hours_late = intTick_minutes_late // 60
			minutes_late = intTick_minutes_late % 60
		else:
			hours_late = 0
			minutes_late = intTick_minutes_late

		wb1.Cells(intID+ 3,7).Value =str(int(intTick_hours_late)+hours_late)+'h'+str(int(minutes_late))+"'["+ str(lsDaylate).replace("[",'').replace("]",'').replace('"','')
		wlwb.Close(True)
		xlApp.Quit()
	if intComebacksoon == 1:
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")

		if intTick_minutes_combacksoon >= 60:
			hours_comebacksoon = intTick_minutes_combacksoon // 60
			minutes_comebacksoon = intTick_minutes_combacksoon % 60
		else:
			hours_comebacksoon = 0
			minutes_comebacksoon = intTick_minutes_combacksoon

		wb1.Cells(intID+ 3,8).Value = str(int(intTick_hours_combacksoon )+hours_comebacksoon)+'h'+str(int(minutes_comebacksoon))+"'["+ str(lsComebacksoon).replace("[",'').replace("]",'').replace('"','') +']'
		wlwb.Close(True)
		xlApp.Quit()
	if intovertime == 1  :
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")

		if intTick_minutes >= 60:
			hours = intTick_minutes // 60
			minutes = intTick_minutes % 60
		else:
			hours = 0
			minutes = intTick_minutes
		wb1.Cells(intID+ 3,9).Value = str(int(intTick_hours)+hours)+'h'+str(int(minutes))+"'["+ str(lsOvertime).replace("[",'').replace("]",'').replace('"','') +']'
		wlwb.Close(True)
		xlApp.Quit()
	if intoverholiday == 1:
		xlApp = win32.Dispatch("Excel.Application")
		wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
		wb1= wlwb.Sheets("社員の勤怠")

		if intTick_minutes_holiday >= 60:
			hours_holiday = intTick_minutes_holiday // 60
			minutes_holiday = intTick_minutes_holiday % 60
		else:
			hours_holiday = 0
			minutes_holiday = intTick_minutes_holiday

		wb1.Cells(intID+ 3,10).Value = str(int(intTick_hours_holiday)+hours_holiday)+'h'+str(int(minutes_holiday))+"'["+ str(lsOverhoidays).replace("[",'').replace("]",'').replace('"','')+']'
		wlwb.Close(True)
		xlApp.Quit()
	# print(lsDaylate)
	# print(intdayoff)






LogTXT('Check diligence') 
dfTick_diligence = pd.read_excel(open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)+'.xlsx','rb'), sheet_name = 0, dtype = object, header =1)
Lsholiday = []
for indexholiday, rowholiday in df_Tick.iterrows():
	if str(rowholiday['TICK']) != 'nan':
		Lsholiday.append(str(rowholiday['日付']))
for index_TD, row in dfTick_diligence.iterrows():
	if str(row[1]) != 'nan':
		if str(row[3]) == '0' or str(row[3]) == 'nan':
			for x in Lsholiday:
				if '/'+ x in str(row[5]):
					intCheck = 1
				else:
					if '/'+ x in str(row[9]):
						intCheck = 1
					else:
						intCheck = 0
			print(str(row[2]) + str(row[5]).split('日')[0] +' '+ str(intCheck))
			if str(row[5]) == 'nan' or intCheck == 1:
				if str(row[5]) != 'nan':
					if float(str(row[5]).split('日')[0]) <= len(Lsholiday):
						if str(row[6]) == 'nan':
							if str(row[7]) == 'nan':
								print(str(row[7]))
								xlApp = win32.Dispatch("Excel.Application")
								wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
								wb1= wlwb.Sheets("社員の勤怠")
								wb1.Cells(index_TD+ 3,12).Value = str('ü')
								wlwb.Close(True)
								xlApp.Quit()
				else:
					if str(row[6]) == 'nan':
						if str(row[7]) == 'nan':
							print(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now))
							xlApp = win32.Dispatch("Excel.Application")
							wlwb = xlApp.Workbooks.Open(CurDir +'\\output\\'+ str(strYear_now)+str(strMonth_now)+'\\JIBANNET'+str(int(strMonth_now))+'_'+str(strYear_now)) 
							wb1= wlwb.Sheets("社員の勤怠")
							wb1.Cells(index_TD+ 3,12).Value = str('ü')
							wlwb.Close(True)
							xlApp.Quit()

LogTXT('-------PROCESS HAS COMPLETED--------')
DisplayMessageBox('-------PROCESS HAS COMPLETED--------')

	