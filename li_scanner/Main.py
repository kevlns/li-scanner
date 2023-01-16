import sys
from PyQt5 import QtWidgets
from li_scanner.CardDesign import CardDesign

app = QtWidgets.QApplication(sys.argv)
ui = CardDesign()
ui.show()
sys.exit(app.exec_())