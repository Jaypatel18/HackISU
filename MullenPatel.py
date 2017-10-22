import MySQLdb
from xlsxwriter.workbook import Workbook


db = MySQLdb.connect(host="mysql.cs.iastate.edu",  # your host
                     user="dbu363esmullen",  # username
                     passwd="eaRTXbe3",  # password
                     db="db363esmullen")  # name of the database

# Cursor to execute queries.
cur = db.cursor()

script = "SELECT DISTINCT * FROM HackISU ORDER BY Property_Name, Loc_STATE, Loc_CITY Asc"
cur.execute(script)

# *********************WRITING TO EXCEL FILES***************************************8
workbook = Workbook('ExcelFileTest_PYTHON.xlsx')

format = workbook.add_format({'bold': True, 'italic': True, 'align': 'center'})
format2 = workbook.add_format({'align': 'center'})
format3 = workbook.add_format({'num_format': 'mm/dd/yyyy'})

sheet = workbook.add_worksheet()
sheet.write(0,0,'ID', format)
sheet.write(0,1,'BUSINESS NAME', format)
sheet.write(0,2,'STATE', format)
sheet.write(0,3,'CITY', format)
sheet.write(0,4,'DATE', format)
sheet.write(0,5,'OUTCOME', format)

for r, row in enumerate(cur.fetchall()):
    for c, col in enumerate(row):
        if c == 4:
            sheet.write(r + 1, c, col, format3)
        else:
            sheet.write(r + 1, c, col, format2)
workbook.close()

# **************************************************************************************************************************
# **************************************************************************************************************************
# **************************************************************************************************************************

