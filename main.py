import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import Qt, QDate, QTime,QDateTime
from PyQt5.QtGui import QMouseEvent
from main_ui import Ui_MainWindow  # Import the generated class
import pandas as pd
import os
from datetime import datetime

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
    
        # Set up the user interface from the generated class
        self.setupUi(self)
    
        # Set flags to remove the default title bar
        self.setWindowFlags(Qt.FramelessWindowHint)
        # Set default page
        self.stackedWidget.setCurrentIndex(1)
        
        self.flag = 'Start'
    
        # Connect buttons to their respective methods
        self.maximizeRestoreAppBtn.clicked.connect(self.maximize_window)
        self.closeAppBtn.clicked.connect(self.close)
        self.minimizeAppBtn.clicked.connect(self.showMinimized)
        # Connect signals to check_enable_start_btn_2 method
                  
        # Set button styles
        self.enabledButtonStyle_Sheet = """
            QPushButton {
                background-color: #D6D3CE;
                color: black;
                font: bold 12pt "Arial";
                border-radius: 0px;
                border: none;
            }
        """
        self.disabledButtonStyle_Sheet = "QPushButton { background-color: #D3D3CE; color: grey; }"
        
        # Connect page-switching buttons
        self.btnOther.clicked.connect(lambda: self.change_page(2))
        self.btnBack.clicked.connect(lambda: self.change_page(1))
        self.btnPool.clicked.connect(lambda: self.change_page(3))
    
        # Connect buttons for recording time
        self.btnStart.clicked.connect(self.save_record)
        self.btnFinish.clicked.connect(self.finish_record)
        self.btnPause.clicked.connect(self.update_pause_start_time)
        
        self.btnStart_2.clicked.connect(self.save_record_2)
        self.btnFinish_2.clicked.connect(self.finish_record_2)
        self.btnPause_2.clicked.connect(self.update_pause_start_time_2)
    
        # Load values from CSV to combo box
        self.load_values_to_cmbxWO()
        
        # Disable buttons at the start
        self.disable_buttons()
        
        # Connect signals for enabling/disabling buttons based on user input
        self.txtID.textChanged.connect(self.check_enable_start_btn)
        self.txtWO.textChanged.connect(self.check_enable_start_btn)
        self.txtPT.textChanged.connect(self.check_enable_start_btn)
        self.txtID_2.textChanged.connect(self.check_enable_start_btn_2)

        self.cmbxWO.currentTextChanged.connect(self.check_enable_start_btn_2)
        self.txtID.textChanged.connect(self.check_existing_entry)
        self.txtID_2.textChanged.connect(self.check_existing_entry_2)

        # Load values from CSV to cmbxWO
    def load_values_to_cmbxWO(self):
        # Load data from CSV
        df = self.load_csv_cmbx()
        if df is not None:
            # Extract unique values from 'Work Order' column
            work_orders = df['Work Order'].unique().tolist()
            # Clear combo box
            self.cmbxWO.clear()
            # Populate combo box with unique work orders
            self.cmbxWO.addItems(work_orders)
    def check_existing_entry_2(self):
        self.flag = "Start"
        # Disconnect textChanged signals
        self.txtID_2.textChanged.disconnect(self.check_enable_start_btn_2)
        self.cmbxWO.currentTextChanged.disconnect(self.check_enable_start_btn_2)
        
        check = False
        
        # Check if the provided ID exists in the CSV file and has no end time
        txtID = self.txtID_2.text()
        print("txtID",txtID)
        if self.entry_exists(txtID):
            # Load corresponding information from the CSV file
            loaded_info = self.load_info_from_csv(txtID)
            if loaded_info:
                for entry in loaded_info:
                    end_time = entry.get('End Time')
                    endMinute = entry.get('Total Time (minutes)')
                    pause_start_time = entry.get('Pause Start Time')
                    pause_end_time = entry.get('Pause End Time')
                    Other = entry.get('Other')
                    if not (pd.isna(end_time) and pd.isna(endMinute)):
                        continue
                    if Other == "Yes":
                        check = True
                        if not pd.isna(pause_start_time) and pd.isna(pause_end_time):
                            self.flag = "Pause"
            
                            # Enable/disable buttons for paused state
                            self.btnStart_2.setEnabled(True)
                            self.btnOther.setEnabled(False)
                            self.btnPause_2.setEnabled(False)
                            self.btnFinish_2.setEnabled(True)
                            self.btnPool.setEnabled(False)
                            # Set button styles
                            self.btnStart_2.setStyleSheet(self.enabledButtonStyle_Sheet)
                            self.btnPause_2.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnFinish_2.setStyleSheet(self.enabledButtonStyle_Sheet)
                            self.btnOther.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPool.setStyleSheet(self.disabledButtonStyle_Sheet)
            
                            # Load information into line edits
                            self.cmbxWO.setCurrentText(str(entry['Work Order']))
            
                        elif pd.isna(end_time) and pd.isna(endMinute):
            
                            # Enable/disable buttons for non-paused state
                            self.flag = "Start"
                            self.btnStart_2.setEnabled(False)
                            self.btnOther.setEnabled(False)
                            self.btnPool.setEnabled(False)
                            self.btnPause_2.setEnabled(True)
                            self.btnFinish_2.setEnabled(True)
                            # Set button styles
                            self.btnStart_2.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPool.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnOther.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPause_2.setStyleSheet(self.enabledButtonStyle_Sheet)
                            self.btnFinish_2.setStyleSheet(self.enabledButtonStyle_Sheet)
            
                            # Load information into line edits
                            self.cmbxWO.setCurrentText(str(entry['Work Order']))
    
        else:
            # Disable all buttons if the entry does not exist
            self.flag = "Start"
            self.btnStart_2.setEnabled(False)
            self.btnOther.setEnabled(True)
            self.btnPool.setEnabled(True)
            self.btnPause_2.setEnabled(False)
            self.btnFinish_2.setEnabled(False)
    
            # Set button styles
            self.btnStart_2.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnOther.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.enabledButtonStyle_Sheet)
    
            self.btnPause_2.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnFinish_2.setStyleSheet(self.disabledButtonStyle_Sheet)
    
            # # Clear line edits
            # self.txtID_2.clear()
            self.cmbxWO.setCurrentText("SELECT")
        if check=="False":
            # Disable all buttons if the entry does not exist
            self.flag = "Start"
            self.btnStart_2.setEnabled(False)
            self.btnOther.setEnabled(True)
            self.btnPool.setEnabled(True)
            self.btnPause_2.setEnabled(False)
            self.btnFinish_2.setEnabled(False)
    
            # Set button styles
            self.btnStart_2.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnOther.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.enabledButtonStyle_Sheet)
    
            self.btnPause_2.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnFinish_2.setStyleSheet(self.disabledButtonStyle_Sheet)
    
            # # Clear line edits
            # self.txtID_2.clear()
            self.cmbxWO.setCurrentText("SELECT")
            
        # Reconnect textChanged signals
        self.txtID_2.textChanged.connect(self.check_enable_start_btn_2)
        self.cmbxWO.currentTextChanged.connect(self.check_enable_start_btn_2)

    def check_existing_entry(self):
        self.flag = "Start"
        check=False
        # Disconnect textChanged signals
        self.txtID.textChanged.disconnect(self.check_enable_start_btn)
        self.txtWO.textChanged.disconnect(self.check_enable_start_btn)
        self.txtPT.textChanged.disconnect(self.check_enable_start_btn)

    
    
    
        # Check if the provided ID exists in the CSV file and has no end time
        txtID = self.txtID.text()
        if self.entry_exists(txtID):
            # Load corresponding information from the CSV file
            loaded_info = self.load_info_from_csv(txtID)
    
            if loaded_info:
                for entry in loaded_info:
                    end_time = entry.get('End Time')
                    endMinute = entry.get('Total Time (minutes)')
                    pause_start_time = entry.get('Pause Start Time')
                    pause_end_time = entry.get('Pause End Time')
                    if not (pd.isna(end_time) and pd.isna(endMinute)):
                        continue
                    Other = entry.get('Other')
                    if Other=="No":
                        check=True
                        if not pd.isna(pause_start_time) and pd.isna(pause_end_time):
                            self.flag = "Pause"
            
                            # Enable/disable buttons for paused state
                            self.btnStart.setEnabled(True)
                            self.btnOther.setEnabled(False)
                            self.btnPool.setEnabled(False)
                            self.btnPause.setEnabled(False)
                            self.btnFinish.setEnabled(True)
                            # Set button styles
                            self.btnStart.setStyleSheet(self.enabledButtonStyle_Sheet)
                            self.btnOther.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPool.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPause.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnFinish.setStyleSheet(self.enabledButtonStyle_Sheet)
            
                            # Load information into line edits
                            self.txtWO.setText(str(entry['Work Order']))
                            self.txtPT.setText(str(entry['Project Task']))
                            self.txtIssue.setText(str(entry['Issue']))
            
                        elif pd.isna(end_time) and pd.isna(endMinute):
            
                            # Enable/disable buttons for non-paused state
                            self.flag = "Start"
                            self.btnStart.setEnabled(False)
                            self.btnOther.setEnabled(False)
                            self.btnPool.setEnabled(False)
                            self.btnPause.setEnabled(True)
                            self.btnFinish.setEnabled(True)
                            # Set button styles
                            self.btnStart.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPool.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnOther.setStyleSheet(self.disabledButtonStyle_Sheet)
                            self.btnPause.setStyleSheet(self.enabledButtonStyle_Sheet)
                            self.btnFinish.setStyleSheet(self.enabledButtonStyle_Sheet)
            
                            # Load information into line edits
                            self.txtWO.setText(str(entry['Work Order']))
                            self.txtPT.setText(str(entry['Project Task']))
                            self.txtIssue.setText(str(entry['Issue']))
    
        else:
            # Disable all buttons if the entry does not exist
            self.flag = "Start"
            self.btnStart.setEnabled(False)
            self.btnOther.setEnabled(True)
            self.btnPool.setEnabled(True)
            self.btnPause.setEnabled(False)
            self.btnFinish.setEnabled(False)
    
            # Set button styles
            self.btnStart.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnOther.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.enabledButtonStyle_Sheet)
    
            self.btnPause.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnFinish.setStyleSheet(self.disabledButtonStyle_Sheet)
    
            # Clear line edits
            self.txtWO.clear()
            self.txtPT.clear()
            self.txtIssue.clear()
        if check==False:
            # Disable all buttons if the entry does not exist
            self.flag = "Start"
            self.btnStart.setEnabled(False)
            self.btnOther.setEnabled(True)
            self.btnPool.setEnabled(True)
            self.btnPause.setEnabled(False)
            self.btnFinish.setEnabled(False)
    
            # Set button styles
            self.btnStart.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnOther.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.enabledButtonStyle_Sheet)
    
            self.btnPause.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnFinish.setStyleSheet(self.disabledButtonStyle_Sheet)
    
            # Clear line edits
            self.txtWO.clear()
            self.txtPT.clear()
            self.txtIssue.clear()
            
    
        # Reconnect textChanged signals
        self.txtID.textChanged.connect(self.check_enable_start_btn)
        self.txtWO.textChanged.connect(self.check_enable_start_btn)
        self.txtPT.textChanged.connect(self.check_enable_start_btn)
    def load_csv_cmbx(self):
        folder_path = "appData"
        csv_file_path = os.path.join(folder_path, "WO.csv")
        try:
            if os.path.exists(csv_file_path):
                return pd.read_csv(csv_file_path)
            else:
                raise FileNotFoundError(f"CSV file not found: {csv_file_path}")
        except Exception as e:
            error_message = "Error loading CSV file: Missing WO.csv File"
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("CSV Loading Error")
            msg_box.setText(error_message)
            msg_box.exec_()
            return None

    def check_enable_start_btn(self):
        # Check if other three input fields are not empty
        if self.txtID.text() and self.txtWO.text() and self.txtPT.text():
            # Enable the start button
            self.btnStart.setEnabled(True)
            self.btnOther.setEnabled(False)
            self.btnPool.setEnabled(False)

            self.btnStart.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnOther.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.disabledButtonStyle_Sheet)

            

        else:
            # Disable the start button
            self.btnStart.setEnabled(False)
            self.btnOther.setEnabled(True)
            self.btnPool.setEnabled(True)
            self.btnOther.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnStart.setStyleSheet(self.disabledButtonStyle_Sheet)
    def check_enable_start_btn_2(self):
        # Check if other three input fields are not empty
        cmbxValue=self.cmbxWO.currentText()
        if self.txtID_2.text() and cmbxValue!="SELECT":
            # Enable the start button
            self.btnStart_2.setEnabled(True)
            self.btnOther.setEnabled(False)
            self.btnPool.setEnabled(False)
            self.btnStart_2.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnOther.setStyleSheet(self.disabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.disabledButtonStyle_Sheet)
        else:
            # Disable the start button
            self.btnStart_2.setEnabled(False)
            self.btnOther.setEnabled(True)
            self.btnPool.setEnabled(True)
            self.btnOther.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnPool.setStyleSheet(self.enabledButtonStyle_Sheet)
            self.btnStart_2.setStyleSheet(self.disabledButtonStyle_Sheet)
            
            



    def disable_buttons(self):
        # Disable all buttons
        self.btnPause.setEnabled(False)
        self.btnFinish.setEnabled(False)
        self.btnStart.setEnabled(False)
        
        self.btnPause_2.setEnabled(False)
        self.btnFinish_2.setEnabled(False)
        self.btnStart_2.setEnabled(False)
        
        self.btnStart.setStyleSheet(self.disabledButtonStyle_Sheet)
        self.btnFinish.setStyleSheet(self.disabledButtonStyle_Sheet)
        self.btnPause.setStyleSheet(self.disabledButtonStyle_Sheet)
        
        self.btnStart_2.setStyleSheet(self.disabledButtonStyle_Sheet)
        self.btnFinish_2.setStyleSheet(self.disabledButtonStyle_Sheet)
        self.btnPause_2.setStyleSheet(self.disabledButtonStyle_Sheet)

    def save_record(self):
        # Get inputs from line edits
        if (self.flag=="Start"):
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
                QMessageBox.information(self, "Success", "Your time has started.")
                # Clear input fields
                self.clear_input_fields()
        elif (self.flag=="Pause"):
            txtID = self.txtID.text()
            self.save_pause_end_time(txtID)
    def save_record_2(self):
        # Get inputs from line edits
        if (self.flag=="Start"):
            txtID = self.txtID_2.text()
            txtWO = self.cmbxWO.currentText()
            
    
            # Check if any input is blank
            if txtID == '' or txtWO == 'SELECT' :
                QMessageBox.warning(self, "Warning", "Please fill all required inputs.")
            else:
                # Save the record to Excel and CSV files
                self.save_to_files_2(txtID, txtWO)
                # Show successful message box
                QMessageBox.information(self, "Success", "Your time has started.")
                # Clear input fields
                self.clear_input_fields_2()
        elif (self.flag=="Pause"):
            txtID = self.txtID_2.text()
            self.save_pause_end_time(txtID)
    def save_pause_end_time(self, txtID):
        try:
            # Load CSV data
            df = self.load_csv_data()
    
            # Check if CSV data is loaded
            if df is not None:
                # Get the index of the last row in the DataFrame
                index = df[df['ID'] == int(txtID)].index
                # Update pause start time with the last row
                last_index = index[-1] if len(index) > 0 else None
                if last_index is not None:
    
                    # Get current date and time
                    current_date = QDate.currentDate().toString("yyyy-MM-dd")
                    current_time = QTime.currentTime().toString("hh:mm:ss")
    
                    # Combine date and time into the desired format
                    current_datetime = current_date + " " + current_time
    
                    # Update the pause end time in the CSV data for the latest ID
                    df.loc[last_index, 'Pause End Time'] = current_time
                    df.loc[last_index, 'Pause End Date Time'] = current_datetime
    
                    # Calculate total pause minutes
                    start_time_str = df.loc[last_index, 'Pause Start Date Time']
                    end_time_str = df.loc[last_index, 'Pause End Date Time']
    
                    start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
                    end_time = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M:%S")
                    pause_duration = (end_time - start_time).total_seconds() / 60
    
                    # Update the total pause minutes in the CSV data
                    df.loc[last_index, 'Pause Duration (seconds)'] = pause_duration * 60
                    df.loc[last_index, 'Pause Duration (minutes)'] = pause_duration
    
                    # Save the updated CSV data to the file
                    self.save_to_csv(df)
    
                    # Show successful message box
                    QMessageBox.information(self, "Success", "Your time has restarted.")
                    self.clear_input_fields()
    
        except Exception as e:
            # Show error message box
            QMessageBox.warning(self, "Error", f"An error occurred in update end pause time: {str(e)}")

    def update_pause_start_time(self):
         try:
             txtID = self.txtID.text()
             df = self.load_csv_data()
             if df is not None:
                 current_date = QDate.currentDate().toString("yyyy-MM-dd")
                 current_time = QTime.currentTime().toString("hh:mm:ss")
                 current_datetime = current_date + " " + current_time
                 index = df[df['ID'] == int(txtID)].index
                 # Update pause start time with the last row
                 last_index = index[-1] if len(index) > 0 else None
                 if last_index is not None:
                     df.loc[last_index, 'Pause Start Time'] = current_time
                     df.loc[last_index, 'Pause Start Date Time'] = current_datetime
                     self.save_to_csv(df)
                     QMessageBox.information(self, "Success", "Your time has paused.")
                 else:
                     QMessageBox.warning(self, "Warning", "No matching ID found.")
             self.clear_input_fields()
             self.clear_input_fields_2()
         except Exception as e:
             QMessageBox.warning(self, "Error", f"An error occurred while updating pause start time: {str(e)}")
   







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
        folder_path = "appData"
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
            entries = df[df['ID'] == txtID].to_dict('records')
            return entries
        return None


    def save_to_files_2(self, txtID, txtWO):
        # Load existing data from CSV
        csv_file_path = os.path.join("appData", "recordCSV.csv")
        if os.path.exists(csv_file_path):
            df = pd.read_csv(csv_file_path)
        else:
            df = pd.DataFrame()
    
        # Get current date and time
        current_date = QDate.currentDate().toString("yyyy-MM-dd")
        current_time = QTime.currentTime().toString("hh:mm:ss")
        current_datetime = current_date + " " + current_time
    
        # Append new data to DataFrame
        new_data = {
            'Date': [current_date],
            'ID': [txtID],
            'Work Order': [txtWO],
            'Other': ['Yes'],
            'Project Task': [""],
            'Issue': [""],
            'Start Date Time': [current_datetime],
            'Start Time': [current_time],
            'Pause Start Time': [''],
            'Pause Start Date Time':[''],
            'Pause End Date Time': [''],
            'Pause End Time': [''],
            'Pause Duration (minutes)': [''],
            'End Date Time': [''],
            'End Time': [''],
            'Total Time (minutes)': ['']
        }
        new_df = pd.DataFrame(new_data)
    
        # Concatenate existing DataFrame with new data
        df = pd.concat([df, new_df], ignore_index=True)
    
        # Save the updated DataFrame to CSV
        df.to_csv(csv_file_path, index=False)
    def save_to_files(self, txtID, txtWO, txtPT, txtIssue):
        # Load existing data from CSV
        csv_file_path = os.path.join("appData", "recordCSV.csv")
        if os.path.exists(csv_file_path):
            df = pd.read_csv(csv_file_path)
        else:
            df = pd.DataFrame()
    
        # Get current date and time
        current_date = QDate.currentDate().toString("yyyy-MM-dd")
        current_time = QTime.currentTime().toString("hh:mm:ss")
        current_datetime = current_date + " " + current_time
    
        # Append new data to DataFrame
        new_data = {
            'Date': [current_date],
            'ID': [txtID],
            'Work Order': [txtWO],
            'Other': ['No'],
            'Project Task': [txtPT],
            'Issue': [txtIssue],
            'Start Date Time': [current_datetime],
            'Start Time': [current_time],
            'Pause Start Time': [''],
            'Pause Start Date Time':[''],
            'Pause End Date Time': [''],
            'Pause End Time': [''],
            'Pause Duration (seconds)':[''],
            'Pause Duration (minutes)': [''],
            'End Date Time': [''],
            'End Time': [''],
            'Total Time (seconds)': [''],
            'Total Time (minutes)': ['']
        }
        new_df = pd.DataFrame(new_data)
    
        # Concatenate existing DataFrame with new data
        df = pd.concat([df, new_df], ignore_index=True)
    
        # Save the updated DataFrame to CSV
        df.to_csv(csv_file_path, index=False)

    def finish_record(self):
        txtID = self.txtID.text()
        if txtID:
            # Update the end time and end date time
            self.update_end_time(txtID)
            
            # Calculate total time spent
            total_time_minutes = self.calculate_total_time(txtID)
            
            if total_time_minutes is not None:
                # Update the total time in the CSV file
                self.update_total_time(txtID, total_time_minutes)
                
                # Disable the finish button after finishing the record
                self.btnFinish.setEnabled(False)
                QMessageBox.information(self, "Success", f"Record Saved. Total time: {total_time_minutes} minutes.")
                self.clear_input_fields()
                self.clear_input_fields_2()
                self.save_to_excel(txtID)

            else:
                QMessageBox.warning(self, "Warning", "Cannot calculate total time.")
        else:
            QMessageBox.warning(self, "Warning", "ID field is empty.")
    def finish_record_2(self):
        txtID = self.txtID_2.text()
        if txtID:
            # Update the end time and end date time
            self.update_end_time(txtID)
            
            # Calculate total time spent
            total_time_minutes = self.calculate_total_time(txtID)
            
            if total_time_minutes is not None:
                # Update the total time in the CSV file
                self.update_total_time(txtID, total_time_minutes)
                
                # Disable the finish button after finishing the record
                self.btnFinish.setEnabled(False)
                QMessageBox.information(self, "Success", f"Record Saved. Total time: {total_time_minutes} minutes.")
                self.clear_input_fields()
                self.clear_input_fields_2()
                self.save_to_excel(txtID)

            else:
                QMessageBox.warning(self, "Warning", "Cannot calculate total time.")
        else:
            QMessageBox.warning(self, "Warning", "ID field is empty.")
    def update_end_time(self, txtID):
        try:
            df = self.load_csv_data()
            if df is not None:
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                current_time = QTime.currentTime().toString("hh:mm:ss")
                current_datetime = current_date + " " + current_time
                index = df[df['ID'] == int(txtID)].index
                if not index.empty:
                    # Get the last index if multiple rows exist for the same ID
                    last_index = index[-1]
                    
                    if 'Pause End Date Time' in df.columns and pd.isnull(df.loc[last_index, 'Pause End Date Time']) and not pd.isnull(df.loc[last_index, 'Pause Start Date Time']):
                        # If Pause End Date Time exists, update Pause End Time and Pause End Date Time
                        df.loc[last_index, 'Pause End Time'] = current_time
                        df.loc[last_index, 'Pause End Date Time'] = current_datetime
    
                        # Calculate pause duration
                        pause_start_time = df.loc[last_index, 'Pause Start Time']
                        pause_start_datetime = df.loc[last_index, 'Pause Start Date Time']
                        pause_start_datetime_obj = QDateTime.fromString(pause_start_datetime, "yyyy-MM-dd hh:mm:ss")
                        current_datetime_obj = QDateTime.fromString(current_datetime, "yyyy-MM-dd hh:mm:ss")
                        pause_duration_seconds = pause_start_datetime_obj.secsTo(current_datetime_obj)
                        pause_duration_minutes = pause_duration_seconds / 60
    
                        df.loc[last_index, 'Pause Duration (seconds)'] = pause_duration_seconds
                        df.loc[last_index, 'Pause Duration (minutes)'] = pause_duration_minutes
    
                    df.loc[last_index, 'End Time'] = current_time
                    df.loc[last_index, 'End Date Time'] = current_datetime
                    self.save_to_csv(df)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred: {str(e)}")

    def save_to_excel(self, txtID):
        try:
            # Load record.csv
            df_record = self.load_info_from_csv(txtID)
            
            # Check if the DataFrame is not empty
            if df_record is not None:
                if len(df_record) > 1:
                    latest_record = df_record[-1]
                else:  # If there is only one row, get that row
                    latest_record = df_record[0]               # Get the last row (latest record)
                Date = latest_record.get('Date')
                ID = latest_record.get('ID')
                WO = latest_record.get('Work Order')
                ProjectTask = latest_record.get('Project Task')
                Issue = latest_record.get('Issue')
                total = latest_record.get('Total Time (minutes)')
                print(total)
                # Define the Excel file path
                self.save_to_excelNow(Date, ID,WO,ProjectTask,Issue,total)
                
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving to Excel: {str(e)}")
    def save_to_excelNow(self, Date, ID, WO, PT, Issue, total):
        try:
            # Load existing data from Excel if it exists
            excel_file_path = os.path.join("data", "records.xlsx")
            if os.path.exists(excel_file_path):
                df = pd.read_excel(excel_file_path)
            else:
                df = pd.DataFrame()
    
            # Get current date
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
    
            # Create a new row as a dictionary
            new_row = {
                'Date Started': current_date,
                'Date Finished': Date,
                'ID': ID,
                'Work Order': WO,
                'Project Task': PT,
                'Issue': Issue,
                'Total Time (minutes)': total
            }
            print(new_row)
            # Append the new row to the list of rows
            rows_to_append = [new_row]
    
            # Convert the list of rows into a DataFrame
            df_to_append = pd.DataFrame(rows_to_append)
    
            # Append the new DataFrame to the existing DataFrame
            df = pd.concat([df, df_to_append], ignore_index=True)
    
            # Save the updated DataFrame to Excel
            df.to_excel(excel_file_path, index=False)
    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving to Excel: {str(e)}")


    def save_to_csv(self, df):
        folder_path = "appData"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
    
        csv_file_path = os.path.join(folder_path, "recordCSV.csv")
        df.to_csv(csv_file_path, index=False)

    def update_pause_start_time_2(self):
        try:
            txtID = self.txtID_2.text()
            df = self.load_csv_data()
            if df is not None:
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                current_time = QTime.currentTime().toString("hh:mm:ss")
                current_datetime = current_date + " " + current_time
                index = df[df['ID'] == int(txtID)].index
                # Update pause start time with the last row
                last_index = index[-1] 
                if last_index is not None:
                    df.loc[last_index, 'Pause Start Time'] = current_time
                    df.loc[last_index, 'Pause Start Date Time'] = current_datetime
                    self.save_to_csv(df)
                    QMessageBox.information(self, "Success", "Your time has paused.")
                else:
                    QMessageBox.warning(self, "Warning", "No matching ID found.")
            self.clear_input_fields()
            self.clear_input_fields_2()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred while updating pause start time: {str(e)}")


    def calculate_total_time(self, txtID):
        df = self.load_csv_data()
        if df is not None:
            # Filter the DataFrame for the specified ID
            entries_for_id = df[df['ID'] == int(txtID)]
            
            if not entries_for_id.empty:
                # Select the latest entry for the ID
                latest_entry = entries_for_id.iloc[-1]
                print(latest_entry)
                
                start_time_str = latest_entry['Start Date Time']
                end_time_str = latest_entry['End Date Time']
                pause_start_time_str = latest_entry['Pause Start Date Time']
                pause_end_time_str = latest_entry['Pause End Date Time']
                
                # Initialize pause_duration
                pause_duration = 0
                
                # Check for NaN values and calculate pause duration
                if pd.notna(pause_start_time_str) and pd.notna(pause_end_time_str):
                    try:
                        start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
                        end_time = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M:%S")
                        pause_start_time = datetime.strptime(pause_start_time_str, "%Y-%m-%d %H:%M:%S")
                        pause_end_time = datetime.strptime(pause_end_time_str, "%Y-%m-%d %H:%M:%S")
                        
                        # Calculate pause duration
                        pause_duration = (pause_end_time - pause_start_time).total_seconds() / 60
                        
                    except ValueError as e:
                        print(f"Error parsing datetime: {e}")
                
                # Calculate total time excluding pause duration
                try:
                    start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
                    end_time = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M:%S")
                    
                    total_time = (end_time - start_time).total_seconds() / 60 - pause_duration
                    return total_time
                    
                except ValueError as e:
                    print(f"Error parsing datetime: {e}")
    
        return None



    
    





        
    def update_total_time(self, txtID, total_time_minutes):
        try:
            df = self.load_csv_data()
            if df is not None:
                index = df[df['ID'] == int(txtID)].index
                if total_time_minutes is not None:
                    # Convert total time from minutes to seconds
                    total_time_seconds = total_time_minutes * 60
                    # Update Total Time (seconds) for the last row
                    last_index = index[-1] if len(index) > 0 else None
                    if last_index is not None:
                        df.loc[last_index, 'Total Time (seconds)'] = total_time_seconds
                    else:
                        QMessageBox.warning(self, "Warning", "No matching ID found.")
                
                # Update Total Time (minutes) for all rows with the specified ID
                df.loc[index, 'Total Time (minutes)'] = total_time_minutes
                self.save_to_csv(df)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred while updating total time: {str(e)}")
    

   

   
    def clear_input_fields(self):
        # Clear all input fields
        self.txtID.clear()
        self.txtWO.clear()
        self.txtPT.clear()
        self.txtIssue.clear()
    def clear_input_fields_2(self):
        # Clear all input fields
        self.txtID_2.clear()
        self.cmbxWO.setCurrentText("SELECT")
     

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
