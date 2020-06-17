import os
import errno

from copy import deepcopy

# importing tkinter and tkinter.ttk 
# and all their functions and classes 
from tkinter import * 
from tkinter import messagebox
from tkinter.ttk import *
  
# importing askopenfile function 
# from class filedialog 
from tkinter.filedialog import askopenfilename, askdirectory 

import pandas as pd
#import numpy as np
#import matplotlib as plt

from datetime import date

print("Libraries imported\n")

# Define value class to locate and store value from test results dataframe
class TestValue:
   def __init__(self, name, coordinate, val):
      self.name = name
      self.coordinate = coordinate
      self.val = val

# Set entries and labels
root = Tk() 
path_group = LabelFrame(root, text="Input Files and Report Destination")
path_group.grid(row=1, column=1, padx=15, pady=15, sticky=W)
ent1 = Entry(path_group, font=40)
ent1.grid(row=2, column=2)
label1 = Label(path_group, text="Test data: ")
label1.grid(row=2, column=1, padx=10, pady=10, sticky=W)
ent2 = Entry(path_group, font=40)
ent2.grid(row=3, column=2)
label2 = Label(path_group, text="Logsheet: ")
label2.grid(row=3, column=1, padx=10, pady=10, sticky=W)
ent3 = Entry(path_group, font=40)
ent3.grid(row=4, column=2)
label3 = Label(path_group, text="Report destination: ")
label3.grid(row=4, column=1, padx=10, pady=10, sticky=W)
# Rig selection
rig_group = LabelFrame(root, text="Rig Selection")
rig_group.grid(row=2, column=1, padx=15, pady=15, sticky=W)
rig = StringVar()
R1 = Radiobutton(rig_group, text="F5 Rig", variable=rig, value="F5")
R1.grid(row=5, column=1, padx=10, pady=10, sticky=W)
R2 = Radiobutton(rig_group, text="Yellow Rig", variable=rig, value="yellow")
R2.grid(row=5, column=2, padx=10, pady=10, sticky=W)
# Report data selection
rep_group = LabelFrame(root, text="Report format")
rep_group.grid(row=3, column=1, padx=15, pady=15, sticky=W)
repVals = StringVar()
R3 = Radiobutton(rep_group, text="Standard", variable=repVals, value='standard')
R3.grid(row=6, column=1, padx=10, pady=10, sticky=W)
R4 = Radiobutton(rep_group, text="Custom", variable=repVals, value='custom', command = lambda:custom_report())
R4.grid(row=6, column=2, padx=10, pady=10, sticky=W)
R5 = Radiobutton(rep_group, text="From Template", variable=repVals, value='template', command = lambda:template_report())
R5.grid(row=6, column=3, padx=10, pady=10, sticky=W)

# Define button styling
style = Style()
style.map("C.TButton",
    foreground=[('pressed', 'red'), ('active', 'blue')],
    background=[('pressed', '!disabled', 'black'), ('active', 'white')]
    ) 

# Set relative path to test log and logsheet
#xls_path = "LOGall-fullycond.xls"
#logsheet = "Logsheet_PR_curve.txt"

# Set headers and intialize an empty data frame
#headers = ["Conditions", "Logs", "f [Hz]", "PR", "VR", f"SG [{chr(176)}C]", f"DG [{chr(176)}C]", f"SSH [{chr(176)}C]", "Duty [kW]", "Flow rate [m3/h]", "IE [%]", "VE [%]", "COP [-]"]
#output_df = pd.DataFrame(columns = headers)

def select_report_destination():
   # First delete exsiting entry (if there is existing entry)
   ent3.delete(0, END)
   # Compile test report name based on date
   today = date.today()
   day = today.strftime("%d%m%y")
   output_file_name = f"test_report_{day}"
   print(output_file_name)
   report_dest = askdirectory()
   # Set the output file directory
   test_report = f"{report_dest}\{output_file_name}.xlsx"
   print(f"Destination fodler: {test_report}")
   if not os.path.exists(os.path.dirname(test_report)):
      try:
         os.makedirs(os.path.dirname(test_report))
      except OSError as exc: # Guard against race condition
         if exc.errno != errno.EEXIST:
               raise
   ent3.insert(END, test_report)

# Read the log file
#results = pd.read_excel(xls_path)
global results

def open_results(): 
   ent1.delete(0, END)
   global results_sheet
   results_sheet = askopenfilename(filetypes =[("all files","*.*")]) 
   if results_sheet is not None:       
      ent1.insert(END, results_sheet)
 
def open_template():
   if ent4:
      ent4.delete(0, END)
   template = askopenfilename(filetypes =[("all files","*.*")])
   with open(template,'r') as testlog:
      ent4.insert(END, template)

def open_logsheet():
   ent2.delete(0, END)
   logfile = askopenfilename(filetypes =[("all files","*.*")])
   with open(logfile,'r') as testlog:
      ent2.insert(END, logfile)

# Define custom report function:
def custom_report():
   if rig.get() =='F5':
      # Load params based on F5 test data format
      test_results = ent1.get()
      test_data_df = pd.read_excel(test_results)
      df_shape = test_data_df.shape
      possible_headers = {}
      report_format = {'headers': [], 'averaged_results': {}}
      for i in range(1, df_shape[0]):
         # Define unit:
         unit = ''
         if type(test_data_df.iloc[i, 1]) == float:
            unit = "[-]"
         else:
            unit = f"[{test_data_df.iloc[i, 1]}]"
         possible_headers[i] = {'param': f"[{i}] {test_data_df.iloc[i, 0]} {unit}", 'unit': unit, 'coordinate': i}
      #param_group = LabelFrame(rep_group, text="Select the Report Parameters:")
      #param_group.grid(row=7, column=1, padx=10, pady=10, columnspan=3, sticky=W)
   
   elif rig.get() == 'yellow':
      # Load params based on Yellow rig test data format
      test_results = ent1.get()
      test_data_df = pd.read_csv(test_results, encoding='latin1') 
      df_shape = test_data_df.shape
      possible_headers = {}
      # Create possible headers dict where each key represents a row in df
      for i in range(1, df_shape[0]):
         # Define name and unit:
         name = f"{test_data_df.iloc[i, 0]}"
         unit = f"[{test_data_df.iloc[i, 1]}]"
         # Append to possible_headers
         possible_headers[i] = {'param': f"[{i}] {name} {unit}", 'unit': unit, 'coordinate': i}

   # for scrolling vertically 
   yscrollbar = Scrollbar(rep_group) 
   yscrollbar.grid(row=7, column=4, rowspan=5, sticky=E+N+S)

   global param_list
   param_list = Listbox(rep_group, selectmode = "multiple", yscrollcommand = yscrollbar.set)
      
   # Widget expands horizontally and  
   # vertically by assigning both to 
   # fill option 
   param_list.grid(row=7, column=0, padx = 10, pady = 10, rowspan=5, columnspan=4, sticky=W+E+N+S)   
   for h in possible_headers:
      param_list.insert(END, possible_headers[h]["param"])
   # Attach listbox to vertical scrollbar 
   yscrollbar.config(command = param_list.yview) 

# Function to get UI for stan
def template_report():
   if rig.get() =='F5':
      # Load params based on F5 test data format
      possible_headers = {}
      report_format = {'headers': [], 'averaged_results': {}}
   elif rig.get() =='yellow':
      # Load params based on F5 test data format
      possible_headers = {}
      report_format = {'headers': [], 'averaged_results': {}}
   global ent4
   ent4 = Entry(rep_group, font=40)
   ent4.grid(row=7, column=0, padx = 10, pady = 10, rowspan=1, columnspan=4, sticky=W+E+N+S)   

   # Define button to select log sheet
   template_btn = Button(rep_group, text ='Browse', style="C.TButton", command = lambda:open_template())
   template_btn.grid(row=7,column=4, padx=10, pady=10)

   
   

# Define function to get parameters from the test data and creat report dataframe
def make_F5_report_df(report_type):   
   if report_type == 'standard':
      rep = {'headers': ["Conditions", "Logs", "f [Hz]", "PR", "VR", f"SG [{chr(176)}C]", f"DG [{chr(176)}C]", f"SSH [{chr(176)}C]", "Duty [kW]", "Flow rate [m3/h]", "IE [%]", "VE [%]", "COP [-]"],
             'averaged_results': {
                  "Logs": "",
                  "Conditions": TestValue(name="Conditions", coordinate=3, val=""),
                  "f [Hz]": TestValue(name="f [Hz]", coordinate=446, val=0),                          # averaged_results["freq"] += results.loc[446, log]
                  "PR": TestValue(name="PR", coordinate=11, val=0),                                   # averaged_results["PR"] = results.loc[11, log]
                  "VR": TestValue(name="VR", coordinate=19, val=0),                                   # averaged_results["VR"] = results.loc[19, log]
                  "Duty [kW]": TestValue(name="Duty [kW]", coordinate=249, val=0),                    # averaged_results["Duty"] += results.loc[249, log]
                  "Flow rate [m3/h]": TestValue(name="Flow rate [m3/h]", coordinate=240, val=0),      # averaged_results["Flow rate"] += results.loc[240, log]
                  f"SG [{chr(176)}C]": TestValue(name=f"SG [{chr(176)}C]", coordinate=30, val=0),     # averaged_results["SG"] += results.loc[30, log]
                  f"DG [{chr(176)}C]": TestValue(name=f"DG [{chr(176)}C]", coordinate=34, val=0),     # averaged_results["DG"] += results.loc[34, log]
                  f"SSH [{chr(176)}C]": TestValue(name=f"SSH [{chr(176)}C]", coordinate=29, val=0),   # averaged_results["SSH"] += results.loc[29, log]  
                  "IE [%]": TestValue(name="IE [%]", coordinate=449, val=0),                          # averaged_results["IE"] += results.loc[449, log]
                  "VE [%]": TestValue(name="VE [%]", coordinate=448, val=0),                          # averaged_results["VE"] += results.loc[448, log]
                  "COP [-]": TestValue(name="COP [-]", coordinate=452, val=0)}}                       # averaged_results["COP"] += results.loc[452, log]
      return rep
   elif report_type == 'custom':
      rep = {'headers': ["Logs"], 'averaged_results': {"Logs": ""}}
      selection = param_list.curselection()
      for e in selection:
         selected = param_list.get(e)
         header = selected[selected.find(']')+2:]
         rep["headers"].append(header)
         parameter = TestValue(name=f"{header}", coordinate=int(selected[1:selected.find(']')]), val=0)
         rep["averaged_results"][parameter.name] = parameter
      return rep
   elif report_type == 'template':
      rep = {'status': 'Comming Soon'}
      return rep

def make_yellow_report_df(report_type):
   if report_type == 'standard':
      # 'Conditions' not available in yellow rig test data
      rep = {'headers': ["Logs", "f [Hz]", "PR", "VR", f"SG [{chr(176)}C]", f"DG [{chr(176)}C]", f"SSH [{chr(176)}C]", "Duty [kW]", "Flow rate [m3/h]", "IE [%]", "VE [%]", "COP [-]"],
             'averaged_results': {
                "Logs": "",
                "f [Hz]": TestValue(name="f [Hz]", coordinate=68, val=0) ,                         #averaged_results["freq"] += float(results.loc[68, log])
                "PR": TestValue(name="PR", coordinate=23, val=0),                                  #averaged_results["PR"] = float(results.loc[23, log])
                "VR": TestValue(name="VR", coordinate=7, val=0),                                   #averaged_results["VR"] = float(results.loc[7, log])
                f"SG [{chr(176)}C]": TestValue(name=f"SG [{chr(176)}C]", coordinate=13, val=0),    #averaged_results["SG"] += float(results.loc[13, log]) 
                f"DG [{chr(176)}C]": TestValue(name=f"DG [{chr(176)}C]", coordinate=17, val=0),    #averaged_results["DG"] += float(results.loc[17, log])
                f"SSH [{chr(176)}C]": TestValue(name=f"SSH [{chr(176)}C]", coordinate=15, val=0),  #averaged_results["SSH"] += float(results.loc[15, log])
                "Duty [kW]": TestValue(name="Duty [kW]", coordinate=59, val=0),                    #averaged_results["Duty"] += float(results.loc[59, log])
                "Flow rate [m3/h]": TestValue(name="Flow rate [m3/h]", coordinate=31, val=0),      #averaged_results["Flow rate"] += float(results.loc[31, log])                            
                "IE [%]": TestValue(name="IE [%]", coordinate=46, val=0),                          #averaged_results["IE"] += float(results.loc[46, log])
                "VE [%]": TestValue(name="VE [%]", coordinate=45, val=0),                          #averaged_results["VE"] += float(results.loc[45, log])
                "COP [-]": TestValue(name="COP [-]", coordinate=62, val=0)}}                        #averaged_results["COP"] += float(results.loc[62, log])
      return rep
   elif report_type == 'custom':
      rep = {'headers': ["Logs"], 'averaged_results': {"Logs": ""}}
      selection = param_list.curselection()
      for e in selection:
         selected = param_list.get(e)
         header = selected[selected.find(']')+2:]
         parameter = TestValue(name=f"{header}", coordinate=int(selected[1:selected.find(']')]), val=0)
         rep["headers"].append(header)
         rep["averaged_results"][parameter.name] = parameter
      return rep
   elif report_type == 'template':
      test_results = ent1.get()
      rep = {'headers': ["Logs"], 'averaged_results': {"Logs": ""}}
      test_data_df = pd.read_csv(test_results, encoding='latin1') 
      df_shape = test_data_df.shape
      template_file = ent4.get()
      with open(template_file, 'r') as template:
         for line in template:
            line = line.strip("\n")
            if ' AS ' in line :
               names = line.split(' AS ')
               new_name = names[1]
               query = names[0]
               for i in range(1, df_shape[0]):
                  if query == test_data_df.iloc[i, 0]:
                     unit = f"[{test_data_df.iloc[i, 1]}]"
                     parameter = TestValue(name=f"{new_name} {unit}", coordinate=i, val=0)
                     rep["headers"].append(f"{new_name} {unit}")
                     rep["averaged_results"][parameter.name] = parameter
            else:
               for i in range(1, df_shape[0]):
                  if line == test_data_df.iloc[i, 0]:
                     name = f"{test_data_df.iloc[i, 0]}"
                     unit = f"[{test_data_df.iloc[i, 1]}]"
                     parameter = TestValue(name=f"{name} {unit}", coordinate=i, val=0)
                     rep["headers"].append(f"{name} {unit}")
                     rep["averaged_results"][parameter.name] = parameter
      return rep

def make_report_df():
   report_type = repVals.get()
   if rig.get() == "F5":
      return make_F5_report_df(report_type)
   elif rig.get() == "yellow":
      return make_yellow_report_df(report_type)

def make_F5_report():
   # Get entries
   test_results = ent1.get()
   logsheet = ent2.get()
   test_report = ent3.get()
   # Set headers and intialize an empty data frame
   target_df = make_report_df()
   headers = target_df['headers']
   output_df = pd.DataFrame(columns = headers)
   print(f"Report made from following files:\nF5 test log: {test_results}\nLogsheet: {logsheet}")
   print(target_df)
   print(output_df.head())
   # Read the log file
   results = pd.read_excel(test_results)
   with open(logsheet,'r') as testlog:
      for line in testlog:
         single_logs = line.split("-")
         # Initialize dictionary with average values
         averaged_results = deepcopy(target_df['averaged_results'])
         averaged_results["Logs"] =  TestValue(name="Logs", coordinate=-1, val=line.strip("\n"))
         #print(averaged_results)
         #print(f"Processing logs: {single_logs}")
         for entry in single_logs:
            log = int(entry)
            for key in averaged_results:
               if key != "Logs":
                  #print(key)
                  param = averaged_results[key]
                  #print(f"Parameter is {param}")
                  #print(f"Name: {param.name}\nCoordinate: {param.coordinate}\nValue: {param.val}")
                  coordinate = param.coordinate
                  if type(results.loc[coordinate, log]) is str or type(averaged_results[key].val) is str:
                     #new_value = averaged_results[key]  
                     #new_value.val =  results.loc[coordinate, log]               
                     #averaged_results[key] = new_value
                     averaged_results[key].val = results.loc[coordinate, log]
                  else:
                     #print(f"Exctracting {param.name} fromg Log: {log}")
                     #new_value = averaged_results[key]  
                     #new_value.val +=  results.loc[coordinate, log]               
                     #averaged_results[key] = new_value
                     averaged_results[key].val += results.loc[coordinate, log]
               
            # Get test data and add to dict
            #Conditions = results.loc[3, log]
         
         # Average data and replace testValue object in averaged_results with values ()
         for key in averaged_results:
            if type(averaged_results[key].val) is not str:
               averaged_results[key] = averaged_results[key].val / len(single_logs)
            else:
               averaged_results[key] = averaged_results[key].val            
            
         # Compile string for conditons
         #test_conditions = f"{round(averaged_results['SG'], 1)} / {round(averaged_results['DG'], 1)} @ {int(averaged_results['freq'] + 0.5)} Hz"

         # Replace testValue object in averaged_results with values ()
         # Add results to output dataframe
         output_df = output_df.append(averaged_results, ignore_index=True)

         # Print results
         #print("Conditions: " + test_conditions)
         #print(f" SG = {round(averaged_results['SG'], 2)}\n DG = {round(averaged_results['DG'], 2)}\n SSH = {round(averaged_results['SSH'], 2)}\n Duty = {round(averaged_results['Duty'], 2)} kW\n Flow rate = {round(averaged_results['Flow rate'], 2)} m3/h\n IE = {round(averaged_results['IE'], 2)} %\n VE = {round(averaged_results['VE'], 2)} %\n COP = {round(averaged_results['COP'], 3)}\n")

   print("Writing report")
   with pd.ExcelWriter(test_report) as writer:    
      # Write the dataframe into test report excel file
      output_df.to_excel(writer, sheet_name='Summary', startrow=1, startcol=1, index=False)
      #summary_sheet = writer.sheets['Summary']
      #summary_sheet.set_column('B:O', 20)
      # Add test log to the report
      results.to_excel(writer, sheet_name='Test_log', startrow=0, startcol=0, index=False)
      #data_sheet = writer.sheets['Test_log']
      #data_sheet.set_column('A:', 40)
   messagebox.showinfo("Success", f"Report successfuly generated at: {test_report}")
   return

# Function to process yellow rig test data
def make_yellow_report():
   # Get entries
   test_results = ent1.get()
   logsheet = ent2.get()
   test_report = ent3.get()
   print("Yellow rig report preparing")
   # Set headers and intialize an empty data frame
   target_df = make_report_df()
   headers = target_df['headers']
   print(headers)
   output_df = pd.DataFrame(columns = headers)
   print(f"Report made from following files:\nF5 test log: {test_results}\nLogsheet: {logsheet}")
   # Read the log file
   results = pd.read_csv(test_results, encoding='latin1')
   # Define list of data which have string value, based on yellow rig report
   string_keys = ['Date of log [-]', 'Model Number [-]', 'Serial Number [-]', 'Tester [-]', 'Compressor Size [-]', 'Motor [-]', 'Economised or Non-Economised [-]']
   with open(logsheet,'r') as testlog:
      for line in testlog:
         single_logs = line.split("-")
         averaged_results = deepcopy(target_df['averaged_results'])
         averaged_results["Logs"] =  TestValue(name="Logs", coordinate=-1, val=line.strip("\n"))
         for entry in single_logs:
            log = entry.strip('\n')
            # Get test data and add to dict
            for key in averaged_results:
               if key != "Logs":
                  param = averaged_results[key]
                  coordinate = int(param.coordinate)
                  if key in string_keys:
                     averaged_results[key].val = results.loc[coordinate, log]
                  else:
                     averaged_results[key].val += float(results.loc[coordinate, log])
            
         # Average data and replace testValue object in averaged_results with values ()
         for key in averaged_results:
            if type(averaged_results[key].val) is not str:
               averaged_results[key] = averaged_results[key].val / len(single_logs)
            else:
               averaged_results[key] = averaged_results[key].val
            
         # Compile string for conditons
         #test_conditions = f"{round(averaged_results['SG'], 1)} / {round(averaged_results['DG'], 1)} @ {int(averaged_results['freq'] + 0.5)} Hz"

         # Add results to output dataframe
         output_df = output_df.append(averaged_results, ignore_index=True)
         
   with pd.ExcelWriter(test_report) as writer:    
      # Write the dataframe into test report excel file
      output_df.to_excel(writer, sheet_name='Summary', startrow=1, startcol=1, index=False)
      #summary_sheet = writer.sheets['Summary']
      #summary_sheet.set_column('B:O', 20)
      # Add test log to the report
      results.to_excel(writer, sheet_name='Test_log', startrow=0, startcol=0, index=False)
      #data_sheet = writer.sheets['Test_log']
      #data_sheet.set_column('A:', 40)
   messagebox.showinfo("Success", f"Report successfuly generated at: {test_report}")
   return

# Close function
def close_window(): 
    root.destroy()

# Detect rig
def make_report():
   detected_rig = rig.get()
   if detected_rig =="F5":
      make_F5_report()
   elif detected_rig == "yellow":
      make_yellow_report()
   return

# Define button to select test data
test_data_btn = Button(path_group, text ='Select test data', style="C.TButton", command = lambda:open_results()) 
test_data_btn.grid(row=2,column=3, padx=10, pady=10)

# Define button to select log sheet
logsheet_btn = Button(path_group, text ='Select Logsheet', style="C.TButton", command = lambda:open_logsheet())
logsheet_btn.grid(row=3,column=3, padx=10, pady=10)

# Define button to select report destination folder
report_dest_btn = Button(path_group, text='Browse', style="C.TButton", command = lambda:select_report_destination())
report_dest_btn.grid(row=4, column=3, padx=10, pady=10)

# Define button to generate the report
make_btn = Button(root, text ='Make report', style="C.TButton", command = lambda:make_report())
make_btn.grid(row=8,column=1, padx=10, pady=10, sticky=W)

# Temporary rig button
#rig_btn = Button(root, text ='Detect Rig', style="C.TButton", command = lambda:detect_rig())
#rig_btn.grid(row=5,column=2, padx=10, pady=10)

# Define button to close/terminate the program
cancel_btn = Button(root, text="Close", style="C.TButton", command = lambda:close_window())
cancel_btn.grid(row=8, column=3, padx=10, pady=10, sticky=W)

mainloop() 