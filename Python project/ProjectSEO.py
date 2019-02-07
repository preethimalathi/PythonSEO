from urllib.request import urlopen
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import xlsxwriter
import sqlite3

g = load_workbook('read.xlsx')
sheet = g.active
url=sheet['A1'].value
#url=input(" Enter a url")
print(url)
u=url
fo=urlopen(url)
a=fo.read()
#print(a)
soup=BeautifulSoup(a,"html.parser")
for script in soup(["script","style"]):
    script.extract()
text=soup.get_text()
lines=(line.strip() for line in text.splitlines())
s=list(lines)
d="".join(s)
c=d.split()
#print(c)

e=len(c)
wordfreq = [c.count(w) for w in c]

g = load_workbook('read.xlsx')
sheet = g.active
key=sheet['A2'].value
#key=input(" Enter a keywords")
r=key.split(',')


dic={}
for s1 in r:
    s1=s1.lower()
    for k,v in zip(c,wordfreq):
        k=k.lower()
        if s1==k:
            dic[s1]=v              
#print(dic)
den=[]
b1=list(dic.values())
c1=list(dic.keys())
length=len(b1)+3

for d1 in b1:
    den.append((d1/e)*100)
    print(den)

workbook1=xlsxwriter.Workbook('write_excel.xlsx')
chart = workbook1.add_chart({'type': 'column'})
wsheet = workbook1.add_worksheet()

wsheet.write('A1',u )
wsheet.write_column('A3', c1)
wsheet.write_column('B3', b1)
wsheet.write_column('C3',den)
print("URL,KEYWORDS,DENSITY ARE STORED INTO SPREADSHEET SUCCESSFULLY")

l='=Sheet1!$B$3:$B$'+str(length)
chart.add_series({'values':l})
wsheet.insert_chart('C10', chart)
print("chart drawn into excel successfully")
workbook1.close()

conn=sqlite3.connect('database7.db')

#conn.execute("create table record1(url varchar(30),keywords varchar(20),density varchar(10))")
#print("Database created successfully");

i=0
try:
    conn.execute("create table record1(url varchar(30),keywords varchar(20),density varchar(10))")
    while i<=(len(b1)-1):       
        conn.execute("insert into record1(url,keywords,density)values(?,?,?)",(u,c1[i],den[i]));
        i=i+1
    conn.commit()
except BaseException:
    #print("TABLE NOT CREATED")
    conn.execute("drop table record1")
    conn.commit()
    #print("Existed Database dropped  successfully");
    conn.execute("create table record1(url varchar(30),keywords varchar(20),density varchar(10))")
    print("Database created successfully");

    i=0
    try:
        while i<=(len(b1)-1):       
            conn.execute("insert into record1(url,keywords,density)values(?,?,?)",(u,c1[i],den[i]));
            i=i+1
        conn.commit()
    except BaseException:
        print("TABLE NOT CREATED")

print("Table inserted successfully");


cur=conn.execute("select * from record1")
with open('final_excel_output.csv', 'w') as write_file:
    cursor = conn.cursor()
    for row in cursor.execute('SELECT * FROM record1'):
        write_file.writelines(row)
        print(row)
        for ro in row:
            print(ro)


    


