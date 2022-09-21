import pylightxl as xl


# C:\STERNEV\@@@Move\$$ Dzen\Спокойствие\Deutsch\Log>dir *.xlsx /B >dir.xlsx

def main():
    # request = xl.readxl(fn='C:\\STERNEV\\@@@Move\\$$ Dzen\\�����������\\Deutsch\\request.xlsx')    
    dir = 'C:\\STERNEV\\@@@Move\\$$ Dzen\\Спокойствие\\Deutsch\\'
    
    # readxl returns a pylightxl database that holds all worksheets and its data
    in_fn = dir + '\\log\\dir.xlsx'
    db = xl.readxl(fn = in_fn)
    
    
    # create a blank db
    db_out = xl.Database()
    
    # add a blank worksheet to the db
    db_out.add_ws(ws="Concat")  
       
    # add a blank worksheet to the db
    db_out.add_ws(ws="Words")   
       
    line_num = 0  
    
    line_num_word = 0


    for row in db.ws(ws='dir').rows:
        print(row[0])
        in_fn = dir + '\\log\\' + row[0]
        db_one = xl.readxl(fn = in_fn)
        
        for row_one in db_one.ws(ws='Sheet1').rows:
            line_num = line_num + 1
            print(row_one[0], row_one[1], row_one[2], row_one[3])
            db_out.ws(ws="Concat").update_index(row = line_num + 1, col = 1, val = row_one[0])
            db_out.ws(ws="Concat").update_index(row = line_num + 1, col = 2, val = row_one[1])
            db_out.ws(ws="Concat").update_index(row = line_num + 1, col = 3, val = row_one[2])
            db_out.ws(ws="Concat").update_index(row = line_num + 1, col = 4, val = row_one[3])
            words = row_one[1].split()
            for word_ in words:
                line_num_word = line_num_word + 1
                db_out.ws(ws="Words").update_index(row = line_num_word  + 1, col = 1, val = word_ )
                

    
    out_fn = dir + '\\log\\out.xlsx'     
    xl.writexl(db=db_out, fn=out_fn) 
    
if __name__ == "__main__":
    main()     