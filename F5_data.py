import os
import errno

import pandas as pd
#import numpy as np
#import matplotlib as plt

from datetime import date

print("Libraries imported\n")

# Set relative path to test log and logsheet
xls_path = "LOGall-fullycond.xls"
logsheet = "Log.txt"

# Compile test report name based on date
today = date.today()
day = today.strftime("%d%m%y")
output_file_name = f"test_report_{day}"
print(output_file_name)

# Set the output file directory
test_report = f"reports\{output_file_name}.xlsx"
if not os.path.exists(os.path.dirname(test_report)):
    try:
        os.makedirs(os.path.dirname(test_report))
    except OSError as exc: # Guard against race condition
        if exc.errno != errno.EEXIST:
            raise

# Read the log file
results = pd.read_excel(xls_path)

# Set headers and intialize an empty data frame
headers = ["Conditions", "Logs", "f [Hz]", "VR", f"SG [{chr(176)}C]", f"DG [{chr(176)}C]", f"SSH [{chr(176)}C]", "Duty [kW]", "Flow rate [m3/h]", "IE [%]", "VE [%]", "COP [-]"]
output_df = pd.DataFrame(columns = headers)
#print(output_df.head())

# Read the log numbers
with open(logsheet,'r') as testlog:
    for line in testlog:
        single_logs = line.split("-")
        # Initialize dictionary with average values
        averaged_results = {"freq": 0, "VR": 1.6, "Duty": 0, "Flow rate": 0, "SG": 0, "DG": 0, "SSH": 0, "IE": 0, "VE": 0, "COP":0}
        print(f"Processing logs: {single_logs}")
        for entry in single_logs:
            log = int(entry)
            # Get test data and add to dict
            Conditions = results.loc[3, log]
            averaged_results["freq"] += results.loc[446, log]
            averaged_results["VR"] = results.loc[19, log]
            averaged_results["Duty"] += results.loc[249, log]
            averaged_results["Flow rate"] += results.loc[240, log]
            averaged_results["SG"] += results.loc[30, log]
            averaged_results["DG"] += results.loc[34, log]
            averaged_results["SSH"] += results.loc[29, log]            
            averaged_results["IE"] += results.loc[449, log]
            averaged_results["VE"] += results.loc[448, log]
            averaged_results["COP"] += results.loc[452, log]
           
        # Average data
        for val in averaged_results:
            if val != "Conditions" and val!= "VR":
                averaged_results[val] = averaged_results[val] / len(single_logs)
            
        # Compile string for conditons
        test_conditions = f"{round(averaged_results['SG'], 1)} / {round(averaged_results['DG'], 1)} @ {int(averaged_results['freq'] + 0.5)} Hz"

        # Add results to output dataframe
        output_df = output_df.append({"Conditions": test_conditions,
                                        "Logs": line.strip('\n'),
                                        "f [Hz]": int(averaged_results["freq"] + 0.5),
                                        "VR": round(averaged_results['VR'], 1),
                                        f"SG [{chr(176)}C]": averaged_results['SG'],
                                        f"DG [{chr(176)}C]": averaged_results['DG'],
                                        f"SSH [{chr(176)}C]": averaged_results["SSH"],
                                        "Duty [kW]": averaged_results["Duty"],
                                        "Flow rate [m3/h]": averaged_results["Flow rate"],
                                        "IE [%]": averaged_results["IE"],
                                        "VE [%]": averaged_results["VE"],
                                        "COP [-]": averaged_results["COP"]}, ignore_index=True)

        # Print results
        print("Conditions: " + test_conditions)
        print(f" SG = {round(averaged_results['SG'], 2)}\n DG = {round(averaged_results['DG'], 2)}\n SSH = {round(averaged_results['SSH'], 2)}\n Duty = {round(averaged_results['Duty'], 2)} kW\n Flow rate = {round(averaged_results['Flow rate'], 2)} m3/h\n IE = {round(averaged_results['IE'], 2)} %\n VE = {round(averaged_results['VE'], 2)} %\n COP = {round(averaged_results['COP'], 3)}\n")


with pd.ExcelWriter(test_report) as writer:    
    # Write the dataframe into test report excel file
    output_df.to_excel(writer, sheet_name='Summary', startrow=1, startcol=1)

    # Add test log to the report
    results.to_excel(writer, sheet_name='Test_log', startrow=0, startcol=0)
    print("Writing finished")      

# Print the output dataframe
print(output_df.head())
            
