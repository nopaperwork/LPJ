# from script to excel
import sqlite3
from xlsxwriter.workbook import Workbook
workbook = Workbook('output_db.xlsx')
worksheet = workbook.add_worksheet()

conn=sqlite3.connect('dont_touch.db')
c=conn.cursor()
c.execute("select * from central_tracker")
mysel=c.execute("select * from central_tracker ")

first_row = 0
ordered_list = ['Article No','RefID','Booking Date','Destination Pincode','Addressee','Addressee Address','Weight','Net Value','LPJ Person','CaseCode','Delivery Status','User Note']
# push the header information
for header in ordered_list:
    col = ordered_list.index(header)  # We are keeping order.
    worksheet.write(
        first_row, col, header
    )  # We have written first row which is the header of worksheet also.
    
for i, row in enumerate(mysel):
    print(row)
    for j, value in enumerate(row):
        worksheet.write(i+1, j, row[j])
workbook.close()
