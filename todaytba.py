# group by Reporter and Status columns and count the values
tba_df2_grouped = tba_df2.groupby(['Reporter', 'Status']).size().reset_index(name='#Records')

# pivot the table to get count of each status for each Reporter
pivot_df2 = tba_df2_grouped.pivot_table(index=['Reporter'], columns='Status', values='#Records', fill_value=0)
pivot_df2 = pivot_df2.rename(columns={"Complete": "Reflow done"})

# add a column for Grand Total
pivot_df2['Grand Total'] = pivot_df2.sum(axis=1)

# add a row for Grand Total
pivot_df2.loc['Grand Total',:] = pivot_df2.sum()

# sort the pivot table by column 'Reflow done' in descending order
pivot_df2 = pivot_df2.sort_values(by='Reflow done', ascending=False)

# format the values as whole numbers
pivot_df2= pivot_df2.applymap('{:,.0f}'.format)





# display the pivot table
print(pivot_df2)


from win32com.client import Dispatch

# initialize Outlook application
outlook = Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# set email properties

mail.To = 'spinascheduling@xperi.com'
mail.CC = 'NAManagers@xperi.com;suresh.kandsaamy@xperi.com'
mail.Subject = f'TBA report - {today}'
mail.Body = 'Hi All, <br> <br> Today’s TBA report has been uploaded to the below public server'


# create HTML tables

pivot_df2_html = pivot_df2.to_html()

# add pivot tables and body text to email body
mail.HTMLBody = f"""\
<html>
  <body>
    <p>{mail.Body}</p>
    <br>
    <p>URL: https://hawkeye.spi-global.com/</p>
    {pivot_df2_html}
  </body>
    <p>Thanks and Regards,<br>
        Muthukumar</p>
    
    <p>Note - This is auto generated mail.</p>
</html>
"""
attachment = mail.Attachments.Add(file2)
attachment.DisplayName = os.path.basename(file2)

# send email
mail.Display()
