
# importing the required libraries
  
from PyQt5.QtCore import * 
from PyQt5.QtGui import * 
from PyQt5.QtWidgets import * 
from PyQt5 import QtGui, QtCore
import sys
  



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
  
  
# create pyqt5 app
App = QApplication(sys.argv)
  
# create the instance of our Window
window = Window()
  
# start the app
sys.exit(App.exec())