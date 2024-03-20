import sys
from PyQt5.QtWidgets import QApplication, QMainWindow,QMessageBox
from PyQt5.QtCore import QPropertyAnimation, QEasingCurve, QUrl, Qt
from PyQt5.QtGui import QIcon, QDesktopServices, QMouseEvent
from main_ui import Ui_MainWindow  # Import the generated class

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        # Set up the user interface from the generated class
        self.setupUi(self)

        # Set flags to remove the default title bar
        self.setWindowFlags(Qt.FramelessWindowHint)
        # Set default page
        self.stackedWidget.setCurrentIndex(1)
        
        # Connect the maximizeRestoreAppBtn button to the maximize_window method
        self.maximizeRestoreAppBtn.clicked.connect(self.maximize_window)

        # Connect the closeAppBtn button to the close method
        self.closeAppBtn.clicked.connect(self.close)

        # Connect the minimizeAppBtn button to the showMinimized method
        self.minimizeAppBtn.clicked.connect(self.showMinimized)
        
        # buttons to switch pages
        self.btnOther.clicked.connect(lambda: self.change_page(2))
        self.btnBack.clicked.connect(lambda: self.change_page(1))
       
        # buttons to Save record
        self.btnStart.clicked.connect(self.save_record)

    def save_record(self):
            # Get inputs from line edits
            txtID = self.txtID.text()
            txtWO = self.txtWO.text()
            txtPT = self.txtPT.text()
            txtIssue = self.txtIssue.text()
    
            # Check if any input is blank
            if txtID == '' or txtWO == '' or txtPT == '':
                QMessageBox.warning(self, "Warning", "Please fill all required inputs.")
            else:
                # Print inputs for now
                print("Input 1:", txtID)
                print("Input 2:", txtWO)
                print("Input 3:", txtPT)
                print("Input 4:", txtIssue)
    def change_page(self, index):
        self.stackedWidget.setCurrentIndex(index)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self.dragPos = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.dragPos)
            event.accept()

    def maximize_window(self):
        # If the window is already maximized, restore it
        if self.isMaximized():
            self.showNormal()
        # Otherwise, maximize it
        else:
            self.showMaximized()
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
