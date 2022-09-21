import pyautogui
import time
import clipboard
import pylightxl as xl
import pathlib

def main():
    # request = xl.readxl(fn='C:\\STERNEV\\@@@Move\\$$ Dzen\\Спокойствие\\Deutsch\\request.xlsx')    
    dir = 'C:\\STERNEV\\@@@Move\\$$ Dzen\\Спокойствие\\Deutsch\\'
    screenWidth, screenHeight = pyautogui.size()
    currentMouseX, currentMouseY = pyautogui.position()
    # minimize window
    # pyautogui.moveTo(1750, 20)
    # pyautogui.click()
    # time.sleep(3)
    
  

    # pylightxl also supports pathlib as well
    my_pathlib = pathlib.Path(dir + 'subtitles.xlsx')
    db_sub = xl.readxl(my_pathlib)
    
    time.sleep(15)
    
    
    # create a blank db
    db = xl.Database()
    
    # add a blank worksheet to the db
    db.add_ws(ws="Sheet1")
    
    # loop to add our data to the worksheet
    # for row_id, data in enumerate(mydata, start=1):
    #     db.ws(ws="Sheet1").update_index(row=row_id, col=1, val=data)
    
    # write out the db

   
    # for k in range(1):
    #     for i in range(20):
    #
    #         for j in range(10):
    #             study(i*10 + j,db)
    #
    #         print('Resting...', i)
    #         time.sleep(25)



    shift = 0
    for k in range(1):
        for i in range(4):
    
            # sub_text = db_sub.ws(ws='Sheet1').index(row = i + 2 + shift, col=1)
            # study_by_subtitles(i, db, sub_text) 
    
            # for j in range(10):
            #     sub_text = db_sub.ws(ws='Sheet1').index(row = i*10 + j + 2 + shift, col=1)
            #     study_by_subtitles(k*40 + i*10 + j, db, sub_text)             
            #     # study(i*10 + j,db)
    
            for j in range(10):
                study(i*10 + j,db)
    
            print('Resting...', i)
            time.sleep(25)
        
    out_fn = dir + '\\log\\log'+ time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    xl.writexl(db=db, fn=out_fn)            
    
    
# this work in a loop       
def study(line_num, db):

    # click on microphone to get source language text
    pyautogui.moveTo(56, 530)
    pyautogui.click()
    time.sleep(6)
    
    # stop recording source language text
    pyautogui.moveTo(56, 560)
    pyautogui.click()
    time.sleep(0.5)
      
    # play the automatic speaker in target language
    pyautogui.moveTo(56, 760)
    pyautogui.click()  
    time.sleep(4)
      
    # play the automatic speaker in target language again   
    pyautogui.moveTo(56, 760)
    pyautogui.click()  
    time.sleep(4)
    
    # record Your speech on target language
    pyautogui.moveTo(1004, 530)
    pyautogui.click()  
    time.sleep(6)
    pyautogui.click()      
        
    # print line number in console. You really need to see this number to keep motivated
    # So you need second monitor to show screen with debug console    
    currentMouseX, currentMouseY = pyautogui.position()
    print(line_num,' ->', currentMouseX, currentMouseY)
    
    #this stuff get text from browser and store it into Excel using copy-paste method
    
    #    column 1 - source language text
    pyautogui.moveTo(420, 450)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=1, val=clipboard.paste())
    
    #    column 2 - stupid translation   
    pyautogui.moveTo(744, 760)
    time.sleep(0.3)    
    pyautogui.click()
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=2, val=clipboard.paste())
    
    
    #    column 3 - text you said on target language
    pyautogui.moveTo(1100, 450)
    time.sleep(0.3)     
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=3, val=clipboard.paste())
    
    #    column 4 - translation of your text from target language to source one
    pyautogui.moveTo(1692, 744)
    time.sleep(0.3)     
    pyautogui.click()
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=4, val=clipboard.paste())
    

def study_by_subtitles(line_num, db, sub_text):

    # push source text from parameter (Youtube subtitles - next paragraph)
    pyautogui.moveTo(56, 450)
    pyautogui.click()
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)    
    pyautogui.hotkey('ctrl', 'x')
    time.sleep(0.1) 
    
    clipboard.copy(sub_text)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.3)
    pyautogui.moveTo(500, 500)
    
    # play the automatic speaker of subtitle
    Y__ = 530 + ( len(sub_text) // 50 ) * 45
    pyautogui.moveTo(120, Y__)
    pyautogui.click()
    time.sleep(8)
    
    # pyautogui.moveTo(90, Y__)
    # pyautogui.click()
    # time.sleep(5)
    
    # # click on microphone to get source language text
    # pyautogui.moveTo(56, 530)
    # pyautogui.click()
    # time.sleep(6)
    #
    # # stop recording source language text
    # pyautogui.moveTo(56, 560)
    # pyautogui.click()
    # time.sleep(0.5)
      
    # # play the automatic speaker in target language
    # pyautogui.moveTo(56, 760)
    # pyautogui.click()  
    # time.sleep(4)
    #
    # # play the automatic speaker in target language again   
    # pyautogui.moveTo(56, 760)
    # pyautogui.click()  
    # time.sleep(4)
    
    # record Your speech on target language
    pyautogui.moveTo(1004, 530)
    pyautogui.click()  
    time.sleep(12)
    pyautogui.click()      
        
    # print line number in console. You really need to see this number to keep motivated
    # So you need second monitor to show screen with debug console    
    currentMouseX, currentMouseY = pyautogui.position()
    print(line_num,' ->', currentMouseX, currentMouseY)
    
    #this stuff get text from browser and store it into Excel using copy-paste method
    
    #    column 1 - source language text
    pyautogui.moveTo(420, 450)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=1, val=clipboard.paste())
    
    #    column 2 - stupid translation   
    pyautogui.moveTo(744, 760)
    time.sleep(0.3)    
    pyautogui.click()
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=2, val=clipboard.paste())
    
    
    #    column 3 - text you said on target language
    pyautogui.moveTo(1100, 450)
    time.sleep(0.3)     
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=3, val=clipboard.paste())
    
    #    column 4 - translation of your text from target language to source one
    pyautogui.moveTo(1692, 744)
    time.sleep(0.3)     
    pyautogui.click()
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=4, val=clipboard.paste())


if __name__ == "__main__":
    main() 
