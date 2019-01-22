from bs4 import BeautifulSoup
from urllib.request import urlopen
from openpyxl import load_workbook
import xlsxwriter
import sqlite3
#sqlite database location
conn=sqlite3.connect("C:/Users/Administrator/AppData/Local/Programs/Python/Python36-32/database3.db")
#conn.execute("create table datas(url varchar(30),keywords varchar(20),density varchar(10))")
workbook1=xlsxwriter.Workbook('outputfile.xlsx')

url=input("enter a url \n")
o=url
print(o)
fo=urlopen(url)
s=fo.read()
soup=BeautifulSoup(s,"html.parser")
for script in soup(["script","style"]):
    script.extract()
text=soup.get_text()
lines=(line.strip() for line in text.splitlines())
lis=list(lines)
st="".join(lis)
q=st.split()
qlen=len(q)
wordfreq = [q.count(w) for w in q]
out={}

wrds=input("enter five words ..\n")
s2=wrds.split()
for w1 in s2:
    w1=w1.lower()
    for a,b in zip(q,wordfreq):
       a=a.lower()
       if w1==a:
           out[w1]=b
print(out)

v=list(out.values())
k=list(out.keys())

length=len(v) + 3
den=[]
j=0
for fre in v:
    den.append((fre/qlen)*100)
chart = workbook1.add_chart({'type': 'column'})
wsheet = workbook1.add_worksheet()
#writing file
wsheet.write('A1',o )
wsheet.write_column('A3', k)
wsheet.write_column('B3', v)
wsheet.write_column('C3',den)
l='=Sheet1!$B$3:$B$'+str(length)
chart.add_series({'values':l})
wsheet.insert_chart('C10', chart)

workbook1.close()
i=0
try:
    while i<=4:       
        conn.execute("insert into datas(url,keywords,density)values(?,?,?)",(o,k[i],v[i]));
        i=i+1
    conn.commit()
except BaseException:
    print("No more word(s) found ")

with open('dboutput.csv', 'w') as write_file:
    cursor = conn.cursor()
    for row in cursor.execute('SELECT * FROM datas'):        
        write_file.writelines(row)

