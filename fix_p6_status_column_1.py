#Fix the Status Update Needed Column in P6. 
#Map the previous weeks update status to this weeks update 
#execfile('fix_p6_status_column.py')
#path Folder X:\groupdirs\0727\NND - Planning & Scheduling\Collaborative Documentation\Weekly excel schedule look ahead\Fix 'Status Needed' field after Global Change 003 problem using Python 27
import pandas as pd
import numpy as np
import xlsxwriter
import os

from pandas import ExcelWriter

#Area to change path to files and python script you want to use
os.chdir('X:\groupdirs\0727\NND - Planning & Scheduling\Collaborative Documentation\Weekly excel schedule look ahead\Fix 'Status Needed' field after Global Change 003 problem using Python 27')

#Function to convert a single string to appropriate integer value
def convert_int(string):
	if pd.isnull(string):
		return None
	else:
		integer = int(string)
	return integer
	
#Import the data
lw_df = pd.read_excel(open('last_week.xls','rb'), sheetname='TASK')
cw_df = pd.read_excel(open('current_week.xls','rb'), sheetname='TASK')

#Save the first row to eventually be used as the header
lw_new_header = lw_df.iloc[0]
cw_new_header = cw_df.iloc[0]

#Remove the first row from the data
lw_df = lw_df[1:]
cw_df = cw_df[1:]

#Rename the header as the first row that was saved earlier
lw_df = lw_df.rename(columns = lw_new_header)
cw_df = cw_df.rename(columns = cw_new_header)

#Convert the index of the dataframe to the Activity ID column. This makes it easir to search and compare data
indexed_lw_df = lw_df.set_index('Activity ID', drop=False)
indexed_cw_df = cw_df.set_index('Activity ID', drop=False)

#Sort the index by ascending by Activity ID. Just to make things line up better
sorted_indexed_lw_df = indexed_lw_df.sort_index()
sorted_indexed_cw_df = indexed_cw_df.sort_index()

#Checks to see if the dataframes index in the first dateframe (sorted_indexed_cw_df) is in the second dataframe (sorted_indexed_lw_df)
#Then creates a new column called "Indicator" in the dataframe "sorted_indexed_cw_df" with a True or False value
#True if the index is in both dataframes or false if it is not
sorted_indexed_cw_df['Indicator'] = sorted_indexed_cw_df.index.isin(sorted_indexed_lw_df.index)

#Filter out all Indicator values that are False from the dataframe "sorted_indexed_cw_df"
#Return to the variable added_activities a dataframe with all activities that were False
cw_new_activities = sorted_indexed_cw_df[sorted_indexed_cw_df['Indicator'] == False]

#Filter out all Indicator values that are True from the dataframe "sorted_indexed_cw_df"
#Return to the variable same_activities a dataframe with all activities that were True
cw_old_activities = sorted_indexed_cw_df[sorted_indexed_cw_df['Indicator'] == True]


################################################################################################
#Now that we know what activities that were in the schedule last week and we have them separated
#in the current week, the program will now start to transfer over the UDF - Integer values
#from last week to this week.
################################################################################################


#Use "convert_int" function above to convert the column 'UDF - Interger' to an integer
field = sorted_indexed_lw_df['UDF - Interger']
sorted_indexed_lw_df['UDF - Interger'] = field.apply(convert_int)

#Drop rows from dataframe 'sorted_indexed_lw_df' that have a no value or NaN
drp_srt_indx_lw_df = sorted_indexed_lw_df[np.isfinite(sorted_indexed_lw_df['UDF - Interger'])]

#Only use the current weeks old activities to filter out activities that don't match all the activities
#in the the dataframe 'drp_srt_indx_lw_df'
drp_srtd_indx_cw_df = cw_old_activities
drp_srtd_indx_cw_df = drp_srtd_indx_cw_df.drop('Indicator',1)



#Creating a True or False value for the 'Indicator' column if one of current weeks old activities
#has the same index as last weeks activities with a value
drp_srtd_indx_cw_df['Indicator'] = drp_srtd_indx_cw_df.index.isin(drp_srt_indx_lw_df.index)

#Need to keep the False values to combine later
keep_srtd_indx_cw_df = drp_srtd_indx_cw_df[drp_srtd_indx_cw_df['Indicator'] == False]

#Creating a data frame where the 'Indicator' column is True
drp_srtd_indx_cw_df = drp_srtd_indx_cw_df[drp_srtd_indx_cw_df['Indicator'] == True]


#Getting rid of the indicator column
drp_srtd_indx_cw_df = drp_srtd_indx_cw_df.drop('Indicator',1)
keep_srtd_indx_cw_df = keep_srtd_indx_cw_df.drop('Indicator',1)
cw_new_activities = cw_new_activities.drop('Indicator',1)

#Moving over last week UDF Integer values to this week
drp_srtd_indx_cw_df['UDF - Interger'] = drp_srt_indx_lw_df['UDF - Interger']


frames = [drp_srtd_indx_cw_df, cw_new_activities, keep_srtd_indx_cw_df] 

result = pd.concat(frames)

result = result.sort_index()

writer = pd.ExcelWriter('PythonExport.xlsx', engine='xlsxwriter')
result.to_excel(writer, sheet_name="TASK")
writer.save()




