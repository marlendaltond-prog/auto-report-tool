import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel

app = QApplication(sys.argv)
window = QMainWindow()
window.setWindowTitle("PyQt5 Test")
window.setGeometry(100, 100, 300, 200)

label = QLabel("PyQt5 GUI is working!")
window.setCentralWidget(label)

window.show()
print("GUI window should be visible now.")

# Add a timer to close the window after 5 seconds
from PyQt5.QtCore import QTimer
QTimer.singleShot(5000, app.quit)

app.exec_()