# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import  QApplication
from label.label import label
import qtmodern.styles                           #DARK_STYLE
import qtmodern.windows                        #DARK_STYLE
if __name__ == "__main__": 
    import sys
#    app = QApplication(sys.argv)
    app = QApplication([])
    qtmodern.styles.dark(app)                                   #dark
    ui = label()
#    ui.show()                                           #dark
    dark_joker = qtmodern.windows.ModernWindow(ui)  #dark
    dark_joker.show()                                              #dark
    sys.exit(app.exec_())

#ME AYTO EGINE ENA PROGRAMMA
#pyinstaller --onefile --windowed label.py

#GIA ALLAGH MENOY 
#self.stackedWidget.setCurrentIndex(2)
