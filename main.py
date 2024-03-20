import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import Qt, QDate, QTime
from PyQt5.QtGui import QMouseEvent
from main_ui import Ui_MainWindow  # Import the generated class
import pandas as pd
import os

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
        # Disable buttons at the start
        self.disable_buttons()
        
        # Connect textChanged signal of other three input fields to enable/disable buttons
        self.txtID.textChanged.connect(self.check_enable_start_btn)
        self.txtWO.textChanged.connect(self.check_enable_start_btn)
        self.txtPT.textChanged.connect(self.check_enable_start_btn)
        
        # Connect textChanged signal of ID field to check existing entry and load information
        self.txtID.textChanged.connect(self.check_existing_entry)

    def check_enable_start_btn(self):
        # Check if other three input fields are not empty
        if self.txtID.text() and self.txtWO.text() and self.txtPT.text():
            # Enable the start button
            self.btnStart.setEnabled(True)
            style_sheet = """
                QPushButton {
                    background-color: #D6D3CE;
                    color: black;
                    font: bold 12pt "Arial";
                    border-radius: 0px;
                    border: none;
                }
            """   
        else:
            # Disable the start button
            self.btnStart.setEnabled(False)
            disabled_color = "#D3D3CE"  # Gray color
            color="grey"
            style_sheet = "QPushButton { background-color: " + disabled_color + ";color: " + color + ";  }"
            self.btnStart.setStyleSheet(style_sheet)

        # Update color of the start button
        self.btnStart.setStyleSheet(style_sheet)

    def disable_buttons(self):
        # Disable all buttons
        self.btnPause.setEnabled(False)
        self.btnFinish.setEnabled(False)
        self.btnStart.setEnabled(False)
        disabled_color = "#D3D3CE"  # Gray color
        color="grey"
        style_sheet = "QPushButton { background-color: " + disabled_color + ";color: " + color + ";  }"
        self.btnStart.setStyleSheet(style_sheet)
        self.btnFinish.setStyleSheet(style_sheet)
        self.btnPause.setStyleSheet(style_sheet)

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
            # Save the record to Excel and CSV files
            self.save_to_files(txtID, txtWO, txtPT, txtIssue)
            # Show successful message box
            QMessageBox.information(self, "Success", "Record saved successfully.")
            # Clear input fields
            self.clear_input_fields()

    def check_existing_entry(self):
        # Check if the provided ID exists in the CSV file and has no end time
        txtID = self.txtID.text()
        print(txtID)
        if self.entry_exists(txtID):
            # Load corresponding information from the CSV file
            print("entry")
            loaded_info = self.load_info_from_csv(txtID)
            print(loaded_info)
            if loaded_info:
                # Enable the pause and finish buttons
                self.btnPause.setEnabled(True)
                self.btnFinish.setEnabled(True)
                # Disable the start button
                self.btnStart.setEnabled(False)
                # Load information into line edits
                self.btnPause.setEnabled(True)
                self.btnFinish.setEnabled(True)
                self.btnStart.setEnabled(False)
                self.txtWO.setText(str(loaded_info['Work Order']))  # Convert to string
                self.txtPT.setText(str(loaded_info['Project Task']))  # Convert to string
                self.txtIssue.setText(str(loaded_info['Issue']))  # Convert to string
        else:
            # Disable the pause and finish buttons
            self.btnPause.setEnabled(False)
            self.btnFinish.setEnabled(False)
            # Enable the start button
            self.btnStart.setEnabled(True)
            self.txtWO.clear()
            self.txtPT.clear()
            self.txtIssue.clear()

    def entry_exists(self, txtID):
        # Check if txtID is empty
        if not txtID:
            return False
    
        # Try to convert txtID to an integer
        try:
            txtID = int(txtID)
        except ValueError:
            # If txtID cannot be converted to an integer, return False
            return False
    
        # Check if the provided ID exists in the CSV file
        df = self.load_csv_data()
        if df is not None:
            return (df['ID'] == txtID).any()
        return False

    def entry_has_end_time(self, txtID):
        # Check if the provided ID has an end time in the CSV file
        df = self.load_csv_data()
        if df is not None:
            entry = df[df['ID'] == txtID]
            if not entry.empty:
                return not entry['End Time'].isna().values[0]
        return False
    def load_csv_data(self):
        folder_path = "data"
        csv_file_path = os.path.join(folder_path, "recordCSV.csv")
        if os.path.exists(csv_file_path):
            return pd.read_csv(csv_file_path)
        return None

    def load_info_from_csv(self, txtID):
        # Check if txtID is empty
        if not txtID:
            return None
    
        # Try to convert txtID to an integer
        try:
            txtID = int(txtID)
        except ValueError:
            # If txtID cannot be converted to an integer, return None
            return None
    
        # Load information corresponding to the provided ID from the CSV file
        df = self.load_csv_data()
        if df is not None:
            entry = df[df['ID'] == txtID]
            if not entry.empty:
                return entry.iloc[0].to_dict()
        return None


    def save_to_files(self, txtID, txtWO, txtPT, txtIssue):
        # Save the record to Excel and CSV files
        current_date = QDate.currentDate().toString("yyyy-MM-dd")
        current_time = QTime.currentTime().toString("hh:mm:ss")

        data = {
            'Date': [current_date],
            'ID': [txtID],
            'Work Order': [txtWO],
            'Project Task': [txtPT],
            'Issue': [txtIssue],
            'Start Time': [current_time],
            'End Time': [''],
            'Total Time': ['']
        }
        df = pd.DataFrame(data)

        folder_path = "data"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        excel_file_path = os.path.join(folder_path, "records.xlsx")
        csv_file_path = os.path.join(folder_path, "recordCSV.csv")
        df.to_excel(excel_file_path, index=False)
        df.to_csv(csv_file_path, index=False)

   

   
    def clear_input_fields(self):
        # Clear all input fields
        self.txtID.clear()
        self.txtWO.clear()
        self.txtPT.clear()
        self.txtIssue.clear()

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
