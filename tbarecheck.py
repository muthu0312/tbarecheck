import pandas as pd
from datetime import datetime, date, timedelta
import os

# set the folder path
folder_path = r'C:\rdp_connection\tba_automation\current_file'


# get today's date
today = date.today()
yesterday = today - timedelta(days=1)

# define a function to read excel files
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path, skiprows=1)
        if df.shape[0] < 1:
            return None
        date = df.iloc[0]['TBA Report Date']
        date_str = datetime.strftime(date, '%m/%d/%Y')
        date = datetime.strptime(date_str, '%m/%d/%Y').date()
        return (date, df)
    except ValueError:
        return None

# initialize variables for the file paths
file1, file2 = None, None

# loop over the files in the folder and read the dates
for file_name in os.listdir(folder_path):
    # check if the file has the correct extension
    if file_name.endswith(".xlsx"):
        # create the full file path
        file_path = os.path.join(folder_path, file_name)
        # read the file and get its date
        result = read_excel(file_path)
        if result is None:
            continue
        date, df = result
        # determine which file is file 1 and which is file 2
        if date == today:
            file1 = file_path if file2 is None else file2
            file2 = file_path if file2 is not None else file1
        else:
            file2 = file_path if file1 is None else file1
            file1 = file_path if file1 is not None else file2

# read file 1 into a data frame
tba_df1 = pd.read_excel(file1, skiprows=1) if file1 is not None else None

# read file 2 into a data frame
tba_df2 = pd.read_excel(file2, skiprows=1) if file2 is not None else None

# check for pending items
if tba_df1 is not None:
    pending_df = tba_df1[tba_df1["Status"] == "Pending"]
    grouped_df = pending_df.groupby(["Reporter", "Source Name"]).size().reset_index(name='count')
    pivot_df = grouped_df.pivot_table(index=['Reporter', 'Source Name'], values='count', aggfunc='sum', margins=True, margins_name='Grand Total')
    pivot_df.rename(columns={'count': '#Records'}, inplace=True)
    pivot_df['#Records'] = pivot_df['#Records'].astype(int)
    pending_status = f"{folder_path}/pending_status_{yesterday}.xlsx"
    writer = pd.ExcelWriter(pending_status, engine='xlsxwriter')
    if pivot_df.loc['Grand Total', '#Records'].item() == 0:
        print("No pending items.")
        writer.close()  # close the writer before removing the file
        os.remove(pending_status)
    else:
        pivot_df.to_excel(writer, sheet_name='Pending_Status', index=True)
        writer.save()
        writer.close()

print(pivot_df.loc['Grand Total', '#Records'].item())

# filter tba_df1 by items that are not marked as completed and are not pending
is_present = (tba_df1['Repeat UID'].isin(tba_df2['Repeat UID'])) & \
             (tba_df1['Status'] == 'Complete') & \
             (tba_df1['Action Taken'] != 'Kept As Is') & \
             (tba_df1['Action Taken'] != 'Reached Out to Provider')

tba_df1_filtered = tba_df1[is_present]


# is_preset = (tba_df2['Repeat UID']).isin(tba_df1_filtered['Repeat UID'])
# tba_df2_filtered = tba_df2[is_present]

# group items by reporter and source name
grouped_df = tba_df1_filtered.groupby(["Reporter", "Source Name"]).size().reset_index(name='count')

# create a pivot table of the grouped data
pivot_df1 = grouped_df.pivot_table(index=['Reporter', 'Source Name'], values='count', aggfunc='sum', margins=True, margins_name='Grand Total')

# rename the count column to #Records and convert the data type to integer
pivot_df1.rename(columns={'count': '#Records'}, inplace=True)
pivot_df1['#Records'] = pivot_df1['#Records'].astype(int)

# create an Excel writer object and write the pivot table to an Excel file
pending_status = f"{folder_path}/stillpending_{today}.xlsx"
pending_record = f"{folder_path}/stillpendingrecord_{today}.xlsx"
writer = pd.ExcelWriter(pending_status, engine='xlsxwriter')
writer2 = pd.ExcelWriter(pending_record, engine ='xlsxwriter')
if pivot_df1.loc['Grand Total', '#Records'].item() == 0:
    print("No pending items.")
    writer.close()  # close the writer before removing the file
    os.remove(pending_status)
    
else:
    tba_df1_filtered.to_excel(writer2,sheet_name='pending_record',index=True)
    pivot_df1.to_excel(writer, sheet_name='Pending_Status', index=True)
    writer.save()
    writer2.save()
    writer.close()
    writer2.close()

print(pivot_df1.loc['Grand Total', '#Records'].item())

from win32com.client import Dispatch

# initialize Outlook application
outlook = Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# set email properties

mail.To = 'spinascheduling@xperi.com'
mail.CC = 'NAManagers@xperi.com;suresh.kandsaamy@xperi.com'
mail.Subject = f'TBA report - {yesterday}'
mail.Body = 'Hi All, <br> <br> Please find the TBA status below.'


# create HTML tables
pivot_df_html = pivot_df.to_html()
pivot_df1_html = pivot_df1.to_html()


    
# add pivot tables and body text to email body
mail.HTMLBody = f"""\
<html>
  <body>
    <p>{mail.Body}</p>
    <p>Pending:</p>
    {pivot_df_html}
    <br>
    <p>Pending on today's file but updated as completed on previous day:</p>
    {pivot_df1_html}
  </body>
    <p>Thanks and Regards,<br>
        Muthukumar</p>
    
    <p>Note - This is auto generated mail.</p>
</html>
"""
# attach the still pending records Excel file
attachment = mail.Attachments.Add(pending_record)
attachment.DisplayName = os.path.basename(pending_record)
attachment = mail.Attachments.Add(file1)
attachment.DisplayName = os.path.basename(file1)

# send email
mail.Display()
