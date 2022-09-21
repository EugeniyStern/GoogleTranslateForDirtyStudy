import pyautogui
import time
import clipboard
import pylightxl as xl

from PyQt5.QtCore import * 
from PyQt5.QtGui import * 
from PyQt5.QtWidgets import * 
from PyQt5 import QtGui, QtCore
import sys
  


def main():
    # request = xl.readxl(fn='C:\\STERNEV\\@@@Move\\$$ Dzen\\Ð¡Ð¿Ð¾ÐºÐ¾Ð¹Ñ�Ñ‚Ð²Ð¸Ðµ\\Deutsch\\request.xlsx')    
    dir = 'C:\\STERNEV\\@@@Move\\$$ Dzen\\Ð¡Ð¿Ð¾ÐºÐ¾Ð¹Ñ�Ñ‚Ð²Ð¸Ðµ\\Deutsch\\'
    screenWidth, screenHeight = pyautogui.size()
    currentMouseX, currentMouseY = pyautogui.position()
    # minimize window
    pyautogui.moveTo(1750, 20)
    pyautogui.click()
    time.sleep(3)
    
  
    # create pyqt5 app
    App = QApplication(sys.argv)
      
    # create the instance of our Window
    window = Window()
      
    # start the app
    # sys.exit(App.exec())
  
    # create a blank db
    db = xl.Database()
    
    # add a blank worksheet to the db
    db.add_ws(ws="Sheet1")
    
    # loop to add our data to the worksheet
    # for row_id, data in enumerate(mydata, start=1):
    #     db.ws(ws="Sheet1").update_index(row=row_id, col=1, val=data)
    
    # write out the db

    for i in range(5):
        study(i,db)    
        
    out_fn = dir + '\\log\\log'+ time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    xl.writexl(db=db, fn=out_fn)            
        
def study(line_num, db):


    pyautogui.moveTo(56, 530)
    pyautogui.click()
    time.sleep(4)
    
 
    pyautogui.moveTo(56, 560)
    pyautogui.click()
    time.sleep(0.5)
      
    pyautogui.moveTo(56, 760)
    pyautogui.click()  
    time.sleep(4)
      
    pyautogui.moveTo(56, 760)
    pyautogui.click()  
    time.sleep(4)
    
    pyautogui.moveTo(1004, 530)
    pyautogui.click()  
    time.sleep(6)
    pyautogui.click()      
        
    currentMouseX, currentMouseY = pyautogui.position()
    print(currentMouseX, currentMouseY)
    
    pyautogui.moveTo(420, 450)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=1, val=clipboard.paste())
    
    pyautogui.moveTo(744, 760)
    time.sleep(2)    
    pyautogui.click()
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=2, val=clipboard.paste())
    
    pyautogui.moveTo(1100, 450)
    time.sleep(2)     
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=3, val=clipboard.paste())
    
    pyautogui.moveTo(1692, 744)
    time.sleep(2)     
    pyautogui.click()
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=4, val=clipboard.paste())
    
def study2(line_num, db):


    pyautogui.moveTo(56, 530)
    pyautogui.click()
    time.sleep(5)
    
 
    pyautogui.moveTo(56, 560)
    pyautogui.click()
    time.sleep(0.3)
      
    # pyautogui.moveTo(56, 760)
    # pyautogui.click()  
    # time.sleep(3)
      
    # pyautogui.moveTo(56, 760)
    # pyautogui.click()  
    # time.sleep(4)
    
    pyautogui.moveTo(1004, 530)
    pyautogui.click()  
    time.sleep(5)

    
    pyautogui.moveTo(1004, 560)
    pyautogui.click()
    time.sleep(0.3)        
        
    currentMouseX, currentMouseY = pyautogui.position()
    print(currentMouseX, currentMouseY)
    
    pyautogui.moveTo(420, 450)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=1, val=clipboard.paste())
    
    pyautogui.moveTo(744, 760)
    time.sleep(2)    
    pyautogui.click()
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=2, val=clipboard.paste())
    
    pyautogui.moveTo(1100, 450)
    time.sleep(2)     
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=3, val=clipboard.paste())
    
    pyautogui.moveTo(1692, 744)
    time.sleep(2)     
    pyautogui.click()
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c')
    db.ws(ws="Sheet1").update_index(row=line_num + 1, col=4, val=clipboard.paste())



class Window(QMainWindow):
  
  
    def __init__(self):
        super().__init__()
  
        TEXT = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " \
        "Nullam malesuada tellus in ex elementum, vitae rutrum enim vestibulum."

  
        # set the title
        self.setWindowTitle("Python")
  
        self.setWindowOpacity(0.7)
        self.setWindowFlags(Qt.CustomizeWindowHint | Qt.FramelessWindowHint | Qt.Dialog | Qt.WindowStaysOnTopHint | Qt.Tool)
  
  
        # setting  the geometry of window
        self.setGeometry(10, 60, 1000, 100)
  
        # creating a label widget
        self.label_1 = QLabel(TEXT, self)
        self.label_1.setFont(QFont('Arial', 12))
        
  
        # making it multi line
        self.label_1.setWordWrap(True)
                
        # moving position
        self.label_1.move(5, 5)
  
        self.label_1.resize(1000, 100)
        
       

  
        # show all the widgets
        self.show()


if __name__ == "__main__":
    main() 
  
