import tkinter as tk 
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog

import pandas as pd 
import xlrd

import os
import requests

HEIGHT = 700
WIDTH = 1250

global answer

## File types specification
def chooseFile():
	my_filetypes = [('all files', '.*'), ('csv files', '.csv')]
	global answer 
	answer = filedialog.askopenfilenames(parent=root, 
		initialdir=os.getcwd(), title='Please select a file:',
		filetypes=my_filetypes)
	def convertTuple(tup):
		str = ''.join(tup)
		return str
	answer = convertTuple(answer)
	#####
	global choice
	choice = selection.get()
	if choice == 1:
		dataAnalyzerAR(answer)
	elif choice == 2:
		dataAnalyzerMM(answer)
	elif choice == 3:
		dataAnalyzerCSI(answer)

## Display Results
def format_response(analysis_kVA):
	try:
		kVA_A = analysis_kVA.at['max', 'Avg. (X)']
		kVA_B = analysis_kVA.at['max', 'Avg. (Y)']
		kVA_C = analysis_kVA.at['max', 'Avg. (Z)']
		kVA_tot = analysis_kVA.at['max', 'Total Avg.']

		#kW_A = analysis_kW.at['max', 'Avg. kW (X)']

		final_str = 'kVA (A): %s \t kVA (B): %s \t kVA (C): %s \t kVA (Total): %s' % (kVA_A, kVA_B, kVA_C, kVA_tot)
	except:
		final_str = 'There was a problem. Please try again.'
	return final_str
	
## Data Analysis (CSI) Code Here
def dataAnalyzerCSI(answer):
	df_all = pd.read_csv(answer)
	#print(df_all.head(5))

	df_kVA = df_all[['Unnamed: 22','Unnamed: 52','Unnamed: 82','Unnamed: 113']]
	df_kVA.columns = ['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']
	df_kVA = df_kVA.iloc[10:]
	df_kVA = df_kVA.dropna()

	df_kVA[['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']] = df_kVA[['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']].astype(float)
	analysis_kVA = df_kVA.describe()

	label['text'] = format_response(analysis_kVA)

## Data Analysis (Morrow Meadows) Code Here
def dataAnalyzerMM(answer):
	# kVA Analysis
	df_all = pd.read_excel(answer)

	df_kVA = df_all[['Apparent Power L1N Avg','Apparent Power L2N Avg','Apparent Power L3N Avg','Apparent Power Total Avg']]
	df_kVA.columns = ['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']
	df_kVA = df_kVA.iloc[0:]
	df_kVA[['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']] = df_kVA[['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']].astype(float)

	analysis_kVA = df_kVA.describe()
	analysis_kVA = analysis_kVA.div(1000)

	label['text'] = format_response(analysis_kVA)

## Data Analysis (A&R Electric) Code Here
def dataAnalyzerAR(answer):
	# kVA Analysis
	df_all = pd.read_csv(answer)
	#print(df_all.head(5))
	
	df_kVA = df_all[['Channel 1.11','Channel 2.11','Channel 3.11','Channel 5.11']]
	df_kVA.columns = ['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']
	df_kVA = df_kVA.iloc[1:]
	df_kVA.head()

	df_kVA[['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']] = df_kVA[['Avg. (X)','Avg. (Y)','Avg. (Z)','Total Avg.']].astype(float)
	analysis_kVA = df_kVA.describe()

	# kW Analysis
	df_kW = df_all[['Channel 1.8','Channel 2.8','Channel 3.8','Channel 5.8']]
	df_kW.columns = ['Avg. kW (X)','Avg. kW (Y)','Avg. kW (Z)','Total Avg. kW']
	df_kW = df_kW.iloc[1:]
	
	df_kW[['Avg. kW (X)','Avg. kW (Y)','Avg. kW (Z)','Total Avg. kW']] = df_kW[['Avg. kW (X)','Avg. kW (Y)','Avg. kW (Z)','Total Avg. kW']].astype(float)

	analysis_kW = df_kW.describe()

	label['text'] = format_response(analysis_kVA)

## UI Code
def openApp():
	global root 
	global selection
	global label

	root = tk.Tk()
	selection = tk.IntVar()
	selection.set(1)

	canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
	canvas.pack()

	#background_image = tk.PhotoImage(file='iFactorLogo.png')
	#background_label = tk.Label(root, image=background_image)
	#background_label.place(relwidth=1, relheight=1)

	frame = tk.Frame(root, bg='#FF6C2E', bd=5)
	frame.place(relx=0.5, rely=0.155, relwidth=.75, relheight=0.1, anchor='n')

	button = tk.Button(frame, text='Select File', font=('Roboto', 18), command=lambda: chooseFile())
	button.place(relx=0.1, relwidth=0.8, relheight=1)

	lower_frame = tk.Frame(root, bg='#FF6C2E', bd=10)
	lower_frame.place(relx=0.5, rely=0.25, relwidth=0.75, relheight=0.65, anchor='n')

	s1 = tk.Radiobutton(lower_frame, text='A&R Electric', variable=selection, value=1)
	s1.place(relx=0.1, rely=0)

	s2 = tk.Radiobutton(lower_frame, text='Morrow Meadows', variable=selection, value=2)
	s2.place(relx=0.42, rely=0)

	s3 = tk.Radiobutton(lower_frame, text='CSI Electrical', variable=selection, value=3)
	s3.place(relx=0.73, rely=0)

	label = tk.Label(lower_frame, font=('Courier', 11), anchor='nw', justify='left', bd=4)
	label.place(relwidth=1, relheight=.8, rely=0.2)

	root.mainloop()

openApp()