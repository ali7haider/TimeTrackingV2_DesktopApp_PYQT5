# import sys
# import pandas as pd
# from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QMessageBox, QVBoxLayout, QWidget, QHBoxLayout
# from datetime import datetime

# class WorkTracker(QMainWindow):
#     def __init__(self):
#         super().__init__()

#         self.setWindowTitle("Work Tracker")
#         self.setGeometry(200, 200, 400, 200)

#         self.operator_label = QLabel("Operator ID:")
#         self.operator_input = QLineEdit()

#         self.workorder_label = QLabel("WorkOrder ID:")
#         self.workorder_input = QLineEdit()

#         self.pt_label = QLabel("PT:")
#         self.pt_input = QLineEdit()

#         self.start_button = QPushButton("Start")
#         self.start_button.clicked.connect(self.start_work)

#         self.pause_button = QPushButton("Pause")
#         self.pause_button.clicked.connect(self.pause_work)
#         self.pause_button.setDisabled(True)

#         layout = QVBoxLayout()
#         layout.addWidget(self.operator_label)
#         layout.addWidget(self.operator_input)
#         layout.addWidget(self.workorder_label)
#         layout.addWidget(self.workorder_input)
#         layout.addWidget(self.pt_label)
#         layout.addWidget(self.pt_input)
#         layout.addWidget(self.start_button)
#         layout.addWidget(self.pause_button)

#         central_widget = QWidget()
#         central_widget.setLayout(layout)
#         self.setCentralWidget(central_widget)

#         # Data logging
#         self.log_file = "work_log.xlsx"

#     def start_work(self):
#         operator_id = self.operator_input.text()
#         workorder_id = self.workorder_input.text()
#         pt = self.pt_input.text()
#         if operator_id and workorder_id and pt:
#             timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#             data = {'Operator ID': operator_id, 'WorkOrder ID': workorder_id, 'PT': pt, 'Start Time': timestamp}
#             self.log_data(data)
#             self.clear_inputs()
#             self.start_button.setDisabled(True)
#             self.pause_button.setEnabled(True)
#         else:
#             QMessageBox.warning(self, "Warning", "Please fill in all fields")

#     def pause_work(self):
#         operator_id = self.operator_input.text()
#         workorder_id = self.workorder_input.text()
#         pt = self.pt_input.text()
#         if operator_id:
#             timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#             data = {'Operator ID': operator_id, 'WorkOrder ID': workorder_id, 'PT': pt, 'Pause Time': timestamp}
#             self.log_data(data)
#             self.operator_input.clear()
#             self.workorder_input.clear()
#             self.pt_input.clear()
#             self.start_button.setEnabled(True)
#             self.pause_button.setDisabled(True)
#         else:
#             QMessageBox.warning(self, "Warning", "Please fill in Operator ID")

#     def log_data(self, data):
#         try:
#             df = pd.read_excel(self.log_file)
#         except FileNotFoundError:
#             df = pd.DataFrame(columns=['Operator ID', 'WorkOrder ID', 'PT', 'Start Time', 'Pause Time'])
#         df = df.append(data, ignore_index=True)
#         df.to_excel(self.log_file, index=False)

#     def clear_inputs(self):
#         self.operator_input.clear()
#         self.workorder_input.clear()
#         self.pt_input.clear()
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
        self.btnHome.clicked.connect(lambda: self.change_page(0))
        self.btnTimeTracking.clicked.connect(lambda: self.change_page(1))
       
        # buttons to Save record
        self.btnSaveRecord.clicked.connect(self.save_record)

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
