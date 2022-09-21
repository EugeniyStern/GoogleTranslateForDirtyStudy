import os, sys
import pylightxl as xl
dirname = 'C:\\STERNEV\\@@@Move\\$$ Dzen\\Спокойствие\\Deutsch\\Log'
files = os.listdir(dirname)

db_out = xl.Database()
db_out.add_ws(ws="Sheet1")

a = sys.getfilesystemencoding()
i = 0
print(a)
temp = map(lambda name: os.path.join(dirname, name), files)
#
#
for filename in list(temp):
    print(filename )

    if(filename == 'C:\STERNEV\@@@Move\$$ Dzen\Спокойствие\Deutsch\Log\~$log20220503-091015.xlsx'): continue

    db = xl.readxl(fn = filename)
    
    for row in db.ws(ws='Sheet1').rows:
        i = i + 1
        print(row)
        db_out.ws(ws="Sheet1").update_index(row=i, col=1, val = row[0])
        db_out.ws(ws="Sheet1").update_index(row=i, col=2, val = row[1])
        db_out.ws(ws="Sheet1").update_index(row=i, col=3, val = row[2])
        db_out.ws(ws="Sheet1").update_index(row=i, col=4, val = row[3])
    
out_filename = 'C:\\STERNEV\\@@@Move\\$$ Dzen\\Спокойствие\\Deutsch\\words.xlsx'
        
xl.writexl(db=db_out, fn=out_filename)   
    