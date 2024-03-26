import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import Qt, QDate, QTime,QDateTime
from PyQt5.QtGui import QMouseEvent
from main_ui import Ui_MainWindow  # Import the generated class
import pandas as pd
import os
from datetime import datetime
import sqlite3
import openpyxl

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
    
        # Set up the user interface from the generated class
        self.setupUi(self)
        self.conn = sqlite3.connect("appData/data.db")
        self.cursor = self.conn.cursor()
        self.create_data_table()
        self.create_pause_table()
    
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
    def create_data_table(self):
        # Drop the table if it exists and create a new Data table
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS Data (
                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                Operator_ID TEXT ,
                                Date TEXT,
                                Work_Order TEXT,
                                Other TEXT,
                                Project_Task TEXT,
                                Qty TEXT,
                                Issue TEXT,
                                Start_Date_Time TEXT,
                                Start_Time TEXT,
                                End_Date_Time TEXT,
                                End_Time TEXT,
                                Total_Time_seconds TEXT,
                                Total_Time_minutes TEXT
                                
                            )''')
        self.conn.commit()  # Commit the transaction to save changes to the database
    def create_pause_table(self):
        # Drop the table if it exists and create a new Pause table
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS Pause (
                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                DataID INTEGER,
                                DataIdx INTEGER,

                                Pause_Start_Date_Time TEXT,
                                Pause_End_Date_Time TEXT,
                                Pause_Start_Time TEXT,
                                Pause_End_Time TEXT,
                                Pause_Duration_seconds TEXT,
                                Pause_Duration_minutes TEXT,
                                FOREIGN KEY(DataID) REFERENCES Data(ID)
                                FOREIGN KEY(DataIdx) REFERENCES Data(ID)

                            )''')
        self.conn.commit()  # Commit the transaction to save changes to the database
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
        if self.entry_exists_in_database(txtID):

            # Load corresponding information from the CSV file
            loaded_info = self.load_info_from_database(txtID)
            if loaded_info is not None and not loaded_info.empty:
                print(loaded_info)
                for _, entry in loaded_info.iterrows():
                    end_time = entry.get('End_Time')  # Adjust the column name
                    end_minute = entry.get('Total_Time_minutes')  # Adjust the column name
                    pause_start_time = entry.get('Pause_Start_Time')  # Adjust the column name
                    pause_end_time = entry.get('Pause_End_Time')  # Adjust the column name
                    Other = entry.get('Other')
                    Work_Order = entry.get('Work_Order')
                    ID = entry.get('ID')  # Adjust the column name
                    latest_pause_data = self.get_latest_pause_data(ID)
                    if not latest_pause_data.empty:
                        pause_start_time = latest_pause_data['Pause_Start_Time'].iloc[0]
                        pause_end_time = latest_pause_data['Pause_End_Time'].iloc[0]
                        print(pause_start_time,pause_end_time)
                    if end_time!='':
                        continue
                    if Other == "Yes":
                        check = True
                        if pause_start_time!="" and pause_end_time is None:
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
                            self.cmbxWO.setCurrentText(str(Work_Order))
            
                        elif end_time=="" and end_minute=="":
            
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
                            self.cmbxWO.setCurrentText(str(Work_Order))
    
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
    def get_latest_pause_data(self, idx):
        try:
            # Execute SQL query to fetch the latest pause data for the specified operator_id
            self.cursor.execute("SELECT * FROM Pause WHERE DataIdx = ? ORDER BY ID DESC LIMIT 1", (idx,))
            
            # Fetch the result of the query
            latest_pause_data = self.cursor.fetchone()
            
            # Check if data was fetched
            if latest_pause_data:
                # Convert the fetched data into a DataFrame
                columns = [column[0] for column in self.cursor.description] 
            # Get column names
                latest_pause_df = pd.DataFrame([latest_pause_data], columns=columns)  # Create DataFrame
                return latest_pause_df
            else:
                return pd.DataFrame()  # Return an empty DataFrame if no data is fetched
        except Exception as e:
            # Handle any exceptions
            QMessageBox.warning(self, "Error", f"An error occurred while retrieving latest pause data: {str(e)}")
            return pd.DataFrame()  # Return an empty DataFrame in case of error


    def check_existing_entry(self):
        self.flag = "Start"
        check=False
        # Disconnect textChanged signals
        self.txtID.textChanged.disconnect(self.check_enable_start_btn)
        self.txtWO.textChanged.disconnect(self.check_enable_start_btn)
        self.txtPT.textChanged.disconnect(self.check_enable_start_btn)

    
    
    
        # Check if the provided ID exists in the CSV file and has no end time
        txtID = self.txtID.text()
      
        if self.entry_exists_in_database(txtID):
            # Load corresponding information from the CSV file
            # Load corresponding information from the CSV file
            loaded_info = self.load_info_from_database(txtID)
            if loaded_info is not None and not loaded_info.empty:
                for _, entry in loaded_info.iterrows():
                    ID = entry.get('ID')  # Adjust the column name
                    latest_pause_data = self.get_latest_pause_data(ID)
                    end_time = entry.get('End_Time')  # Adjust the column name
                    end_minute = entry.get('Total_Time_minutes')  # Adjust the column name
                    pause_start_time = ""  # Adjust the column name
                    pause_end_time = "" # Adjust the column name
                    Other = entry.get('Other')
                    Work_Order = entry.get('Work_Order')
                    Project_Task = entry.get('Project_Task')
                    Issue = entry.get('Issue')
                    if not latest_pause_data.empty:
                        pause_start_time = latest_pause_data['Pause_Start_Time'].iloc[0]
                        pause_end_time = latest_pause_data['Pause_End_Time'].iloc[0]
                        
                    if end_time!='':
                        continue
                    if Other == "No":
                        check=True
                        if pause_start_time!="" and pause_end_time is None:
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
                            self.txtWO.setText(str(Work_Order))
                            self.txtPT.setText(str(Project_Task))
                            self.txtIssue.setText(str(Issue))
            
                        elif end_time=="" and end_minute=="":
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
                            self.txtWO.setText(str(Work_Order))
                            self.txtPT.setText(str(Project_Task))
                            self.txtIssue.setText(str(Issue))
            
    
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
                self.save_to_database(txtID, txtWO, txtPT, txtIssue)
                # Show successful message box
                QMessageBox.information(self, "Success", "Your time has started.")
                # Clear input fields
                self.clear_input_fields()
        elif (self.flag=="Pause"):
            txtID = self.txtID.text()
            txtWO = self.txtWO.text()
            txtPT = self.txtPT.text()
            txtIssue = self.txtIssue.text()
            self.save_pause_end_time(txtID,txtWO,txtPT,txtIssue)
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
                self.save_to_database_2(txtID, txtWO)
                # Show successful message box
                QMessageBox.information(self, "Success", "Your time has started.")
                # Clear input fields
                self.clear_input_fields_2()
        elif (self.flag=="Pause"):
            txtID = self.txtID_2.text()
            cmbxWO = self.cmbxWO.currentText()
            self.save_pause_end_time_2(txtID,cmbxWO)
    def save_pause_end_time(self, txtID,txtWO,txtPT,txtIssue):
        try:
            df=self.load_matching_entries_from_db(txtID,txtIssue,txtWO,txtPT,self.cursor)
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']
                # Get current date and time
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                current_time = QTime.currentTime().toString("hh:mm:ss")
                current_datetime = current_date + " " + current_time
        
                self.cursor.execute("UPDATE Pause SET Pause_End_Time = ?, Pause_End_Date_Time = ? WHERE DataIdx = ? AND Pause_End_Date_Time IS NULL",
                    (current_time, current_datetime, int(entry_id)))

                self.cursor.execute("""
                UPDATE Pause 
                SET 
                    Pause_Duration_seconds = (strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)),
                    Pause_Duration_minutes = (CAST((strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)) AS REAL) / 60)
                WHERE 
                    DataIdx = ? AND
                    Pause_End_Date_Time IS NOT NULL AND Pause_Start_Date_Time IS NOT NULL
            """, (int(entry_id),))
    
            # Commit the transaction to save changes to the database
                self.conn.commit()

                # Commit the changes to the database
        
                # Show successful message box
                QMessageBox.information(self, "Success", "Your pause end time has been updated.")
                self.clear_input_fields()
    
        except Exception as e:
            # Show error message box
            QMessageBox.warning(self, "Error", f"An error occurred in updating end pause time: {str(e)}")
    def save_pause_end_time_2(self, txtID,cmbxWO):
        try:
            df=self.load_matching_entries_from_db_2(txtID,cmbxWO,self.cursor)
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']
                # Get current date and time
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                current_time = QTime.currentTime().toString("hh:mm:ss")
                current_datetime = current_date + " " + current_time
        
                self.cursor.execute("UPDATE Pause SET Pause_End_Time = ?, Pause_End_Date_Time = ? WHERE DataIdx = ? AND Pause_End_Date_Time IS NULL",
                    (current_time, current_datetime, ))

                self.cursor.execute("""
                UPDATE Pause 
                SET 
                    Pause_Duration_seconds = (strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)),
                    Pause_Duration_minutes = (CAST((strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)) AS REAL) / 60)
                WHERE 
                    DataIdx = ? AND
                    Pause_End_Date_Time IS NOT NULL AND Pause_Start_Date_Time IS NOT NULL
            """, (int(entry_id),))
    
            # Commit the transaction to save changes to the database
                self.conn.commit()

                # Commit the changes to the database
        
                # Show successful message box
                QMessageBox.information(self, "Success", "Your pause end time has been updated.")
                self.clear_input_fields()
    
        except Exception as e:
            # Show error message box
            QMessageBox.warning(self, "Error", f"An error occurred in updating end pause time: {str(e)}")
    
    def load_matching_entries_from_db(self,txtID, txtIssue, txtWO, txtPT, cursor):
        try:
            # Query the database to fetch matching entries
            print()
            cursor.execute("""
                SELECT * FROM Data
                WHERE Operator_ID = ? AND Issue = ? AND Work_Order = ? AND Project_Task = ?
            """, (txtID, txtIssue, txtWO, txtPT))
            
            # Fetch all matching entries
            matching_entries = cursor.fetchall()
            
            # Convert the fetched data into a DataFrame
            df = pd.DataFrame(matching_entries, columns=['ID', 'Date', 'Operaror_ID','Work_Order', 'Other', 
                                                          'Project_Task','Qty', 'Issue', 'Start_Date_Time',
                                                          'Start_Time', 'End_Date_Time', 'End_Time',
                                                          'Total_Time_seconds', 'Total_Time_minutes'])
            
            return df
        except Exception as e:
            print(f"Error loading matching entries from database: {e}")
            return pd.DataFrame()  # Return an empty DataFrame if an error occurs
    def load_matching_entries_from_db_2(self,txtID, txtWO,cursor):
        try:
            # Query the database to fetch matching entries
            cursor.execute("""
                SELECT * FROM Data
                WHERE Operator_ID = ? AND Work_Order = ?
            """, (txtID,txtWO))
            
            # Fetch all matching entries
            matching_entries = cursor.fetchall()
            
            # Convert the fetched data into a DataFrame
            df = pd.DataFrame(matching_entries, columns=['ID', 'Date', 'Operator_ID','Work_Order', 'Other', 
                                                          'Project_Task','Qty', 'Issue', 'Start_Date_Time',
                                                          'Start_Time', 'End_Date_Time', 'End_Time',
                                                          'Total_Time_seconds', 'Total_Time_minutes'])
            
            return df
        except Exception as e:
            print(f"Error loading matching entries from database: {e}")
            return pd.DataFrame()  # Return an empty DataFrame if an error occurs

    def update_pause_start_time(self):
        try:
            # Get current date and time
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
            current_time = QTime.currentTime().toString("hh:mm:ss")
            current_datetime = current_date + " " + current_time
            
            # Match current user data to find all entries with the same data
            txtID = self.txtID.text()
            txtIssue = self.txtIssue.text()
            txtWO = self.txtWO.text()
            txtPT = self.txtPT.text()
            df = self.load_matching_entries_from_db(txtID, txtIssue, txtWO, txtPT, self.cursor)
            
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                print("Last index:", last_index)
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']  # Assuming 'ID' is the column name for the primary key
                
                # Save pause data to the pause table
                self.save_pause_data_to_db(entry_id, current_time, current_datetime)
            self.clear_input_fields()
            self.clear_input_fields_2()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred while updating pause start time: {str(e)}")

    def save_pause_data_to_db(self, entry_id, current_time, current_datetime):
        try:
            print("Entry",entry_id)
            # Insert pause data into the pause table
            self.cursor.execute('''INSERT INTO Pause (DataID, DataIdx,Pause_Start_Time, Pause_Start_Date_Time) 
                                   VALUES (?, ?, ?,?)''', (entry_id,int(entry_id), current_time, current_datetime))
            self.conn.commit()
            QMessageBox.information(self, "Success", "Pause data saved successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred while saving pause data: {str(e)}")
    
    
    
    
    

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
    def entry_exists_in_database(self, txtID):
        try:
            # Execute SQL query to count entries with matching txtID in Data table
            self.cursor.execute("SELECT COUNT(*) FROM Data WHERE Operator_ID = ?", (txtID,))

            
            # Fetch the result of the query
            count = self.cursor.fetchone()[0]
            
            # Return True if count is greater than 0, indicating that entries exist
            return count > 0
        except Exception as e:
            # Handle any exceptions
            QMessageBox.warning(self, "Error", f"An error occurred while checking entry in database: {str(e)}")
            return False


    def load_data_from_database(self):
        # Query the database to load data
        self.cursor.execute("SELECT * FROM Data")
        rows = self.cursor.fetchall()
        columns = [column[0] for column in self.cursor.description]
        df = pd.DataFrame(rows, columns=columns)
        return df
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
    def save_to_database(self, txtID, txtWO, txtPT, txtIssue):
        try:
            # Get current date and time
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
            current_time = QTime.currentTime().toString("hh:mm:ss")
            current_datetime = current_date + " " + current_time
            
            # Insert new data into the Data table
            self.cursor.execute('''INSERT INTO Data (Date, Operator_ID, Work_Order, Other, Project_Task, Issue, 
                                Start_Date_Time, Start_Time, End_Date_Time, End_Time, Total_Time_seconds, 
                                Total_Time_minutes) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                (current_date, txtID, txtWO, 'No', txtPT, txtIssue, current_datetime, 
                                current_time, '', '', '', ''))
            self.conn.commit()  # Commit the transaction to save changes to the database
            
            print("Data saved to database successfully.")
        except Exception as e:
            print(f"An error occurred while saving data to the database: {str(e)}")
    def save_to_database_2(self, txtID, txtWO):
        try:
            # Get current date and time
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
            current_time = QTime.currentTime().toString("hh:mm:ss")
            current_datetime = current_date + " " + current_time
            
            # Insert new data into the Data table
            self.cursor.execute('''INSERT INTO Data (Date, Operator_ID, Work_Order, Other, Project_Task, Issue, 
                                Start_Date_Time, Start_Time, End_Date_Time, End_Time, Total_Time_seconds, 
                                Total_Time_minutes) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                (current_date, txtID, txtWO, 'Yes', '', '', current_datetime, 
                                current_time, '', '', '', ''))
            self.conn.commit()  # Commit the transaction to save changes to the database
            
            print("Data saved to database successfully.")
        except Exception as e:
            print(f"An error occurred while saving data to the database: {str(e)}")


    def finish_record(self):
        txtID = self.txtID.text()
        txtIssue = self.txtIssue.text()
        txtWO = self.txtWO.text()
        txtPT = self.txtPT.text()
        if txtID:
            # Update the end time and end date time
            self.update_end_time(txtID,txtIssue,txtWO,txtPT)
            
            # Calculate total time spent
            df=self.load_matching_entries_from_db(txtID,txtIssue,txtWO,txtPT,self.cursor)
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']
                self.calculate_total_time(entry_id)
                self.btnFinish.setEnabled(False)
                QMessageBox.information(self, "Success", f"Record Saved. Total time: minutes.")
                self.clear_input_fields()
                self.clear_input_fields_2()
                self.save_to_excel(txtID,txtIssue,txtWO,txtPT)

        else:
            QMessageBox.warning(self, "Warning", "ID field is empty.")
    def finish_record_2(self):
        txtID = self.txtID_2.text()
        cmbxWO = self.cmbxWO.currentText()
        if txtID:
            # Update the end time and end date time
            self.update_end_time_2(txtID,cmbxWO)
            # Calculate total time spent
            df=self.load_matching_entries_from_db_2(txtID,cmbxWO,self.cursor)
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']
                # Calculate total time spent
                self.calculate_total_time(entry_id)
                self.btnFinish.setEnabled(False)
                QMessageBox.information(self, "Success", f"Record Saved. Total time: minutes.")
                self.clear_input_fields()
                self.clear_input_fields_2()
                self.save_to_excel_2(txtID,cmbxWO)
        else:
            QMessageBox.warning(self, "Warning", "ID field is empty.")
    def update_end_time(self, txtID,txtIssue,txtWO,txtPT):
        try:
            df=self.load_matching_entries_from_db(txtID,txtIssue,txtWO,txtPT,self.cursor)
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']
                # Get current date and time
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                current_time = QTime.currentTime().toString("hh:mm:ss")
                current_datetime = current_date + " " + current_time
                self.cursor.execute("UPDATE Pause SET Pause_End_Time = ?, Pause_End_Date_Time = ? WHERE DataIdx = ? AND Pause_End_Date_Time IS NULL",
                    (current_time, current_datetime, int(entry_id)))
                

                self.cursor.execute("""
                UPDATE Pause 
                SET 
                    Pause_Duration_seconds = (strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)),
                    Pause_Duration_minutes = (CAST((strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)) AS REAL) / 60)
                WHERE 
                    DataIdx = ? AND
                    Pause_End_Date_Time IS NOT NULL AND Pause_Start_Date_Time IS NOT NULL
            """, (int(entry_id),))
              
                self.cursor.execute("UPDATE Data SET End_Time = ?, End_Date_Time = ? WHERE ID = ?",
                            (current_time, current_datetime, int(entry_id)))

            # Commit the transaction to save changes to the database
                self.conn.commit()

                # Commit the changes to the database
        
                # Show successful message box
                QMessageBox.information(self, "Success", "Your pause end time has been updated.")
                self.clear_input_fields()
    
        except Exception as e:
            # Show error message box
            QMessageBox.warning(self, "Error", f"An error occurred in updating end pause time: {str(e)}")
    def update_end_time_2(self, txtID,txtIssue,txtWO,txtPT):
        try:
            df=self.load_matching_entries_from_db(txtID,txtIssue,txtWO,txtPT,self.cursor)
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']
                # Get current date and time
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                current_time = QTime.currentTime().toString("hh:mm:ss")
                current_datetime = current_date + " " + current_time
        
                self.cursor.execute("UPDATE Pause SET Pause_End_Time = ?, Pause_End_Date_Time = ? WHERE DataIdx = ? AND Pause_End_Date_Time IS NULL",
                    (current_time, current_datetime, int(entry_id)))

                self.cursor.execute("""
                UPDATE Pause 
                SET 
                    Pause_Duration_seconds = (strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)),
                    Pause_Duration_minutes = (CAST((strftime('%s', Pause_End_Date_Time) - strftime('%s', Pause_Start_Date_Time)) AS REAL) / 60)
                WHERE 
                    DataIdx = ? AND
                    Pause_End_Date_Time IS NOT NULL AND Pause_Start_Date_Time IS NOT NULL
            """, (int(entry_id),))
                self.cursor.execute("UPDATE Data SET End_Time = ?, End_Date_Time = ? WHERE ID = ?",
                            (current_time, current_datetime, int(entry_id)))
    
            # Commit the transaction to save changes to the database
                self.conn.commit()

                # Commit the changes to the database
        
                # Show successful message box
                QMessageBox.information(self, "Success", "Your pause end time has been updated.")
                self.clear_input_fields()
    
        except Exception as e:
            # Show error message box
            QMessageBox.warning(self, "Error", f"An error occurred in updating end pause time: {str(e)}")


    
    def save_to_excel(self, txtID, txtIssue, txtWO, txtPT):
        try:
            df = self.load_matching_entries_from_db(txtID, txtIssue, txtWO, txtPT, self.cursor)
            if df is not None and not df.empty:
                # Extract the last entry's ID
                last_entry_id = df['ID'].iloc[-1]
                
                # Get the specific column data for the last entry
                specific_column_data = self.get_specific_column(last_entry_id)
                print("specific", specific_column_data)
                
                # Define column name mappings
                column_name_mapping = {
                    'Date': 'Date',
                    'Operaror_ID': 'Operator ID',
                    'Work_Order': 'Work Order',
                    'Project_Task': 'Project Task',
                    'Issue': 'Issue',
                    'Total_Time_seconds': 'Total Time (seconds)',
                    'Total_Time_minutes': 'Total Time (minutes)',
                }
                
                # Rename columns according to the mapping
                specific_column_data.rename(columns=column_name_mapping, inplace=True)
                
                # Add a new column "Date Ended" with today's date
                # specific_column_data['Date Ended'] = datetime.now().strftime('%Y-%m-%d')
                
                # Check if the Excel file exists
                file_path = "data/records.xlsx"
                if not os.path.exists(file_path):
                    # Create a new Excel file with custom column names
                    specific_column_data.to_excel(file_path, index=False)
                    print(f"New Excel file created at {file_path}")
                else:
                    # Load existing data from Excel file
                    existing_df = pd.read_excel(file_path)
                    
                    # Append only the new data row to the existing DataFrame
                    existing_df = existing_df.append(specific_column_data, ignore_index=True)
                    
                    # Save the updated DataFrame back to the Excel file
                    existing_df.to_excel(file_path, index=False)
                    print(f"Data appended to {file_path} successfully.")

        except Exception as e:
            print(f"An error occurred while saving data to Excel: {e}")






    def get_specific_column(self, entry_id):
        try:
            id=int(entry_id)
            # Execute SQL query to fetch the specific column for the given entry_id
            self.cursor.execute(f"""
                                SELECT Date, Operator_ID, Work_Order, Project_Task, Qty,Issue, Total_Time_seconds, Total_Time_minutes
                                FROM Data 
                                WHERE ID = ?
                                """,
                                (int(id),))
            
            # Fetch the result of the query
            result = self.cursor.fetchone()
            print("Result", result)
            
            # Convert the fetched result into a DataFrame
            column_names = ['Date', 'Operator_ID', 'Work_Order', 'Project_Task','Qty', 'Issue', 'Total_Time_seconds', 'Total_Time_minutes']
            specific_column_df = pd.DataFrame([result], columns=column_names)
            
            return specific_column_df
        except Exception as e:
            print(f"An error occurred while retrieving specific column: {e}")
            return None


    def save_to_excel_2(self, txtID, cmbxWO):
        try:
            df = self.load_matching_entries_from_db_2(txtID, cmbxWO, self.cursor)
            if df is not None and not df.empty:
                # Extract the last entry's ID
                last_entry_id = df['ID'].iloc[-1]
                
                # Get the specific column data for the last entry
                specific_column_data = self.get_specific_column(last_entry_id)
                
                # Define column name mappings
                column_name_mapping = {
                    'Date': 'Date',
                    'Operaror_ID': 'Operator ID',
                    'Work_Order': 'Work Order',
                    'Project_Task': 'Project Task',
                    'Issue': 'Issue',
                    'Total_Time_seconds': 'Total Time (seconds)',
                    'Total_Time_minutes': 'Total Time (minutes)',
                }
                
                # Rename columns according to the mapping
                specific_column_data.rename(columns=column_name_mapping, inplace=True)
                
                # Add a new column "Date Ended" with today's date
                # specific_column_data['Date Ended'] = datetime.now().strftime('%Y-%m-%d')
                
                # Check if the Excel file exists
                file_path = "data/records.xlsx"
                if not os.path.exists(file_path):
                    # Create a new Excel file with custom column names
                    specific_column_data.to_excel(file_path, index=False)
                    print(f"New Excel file created at {file_path}")
                else:
                    # Load existing data from Excel file
                    existing_df = pd.read_excel(file_path)
                    
                    # Append only the new data row to the existing DataFrame
                    existing_df = existing_df.append(specific_column_data, ignore_index=True)
                    
                    # Save the updated DataFrame back to the Excel file
                    existing_df.to_excel(file_path, index=False)
                    print(f"Data appended to {file_path} successfully.")

        except Exception as e:
            print(f"An error occurred while saving data to Excel: {e}")

    def load_info_from_database(self, txtID):
        try:
            # Check if txtID is empty or not an integer
            if not txtID:
                print("Error: ID is empty.")
                return None
            
            
            # Query the database to fetch all rows with the specified ID
            self.cursor.execute('''SELECT * FROM Data WHERE Operator_ID = ?''', (txtID,))
            rows = self.cursor.fetchall()
            
            # Check if any rows were found
            if not rows:
                print("No data found for the specified ID.")
                return None
            
            # Convert the fetched rows into a DataFrame
            columns = [desc[0] for desc in self.cursor.description]
            df = pd.DataFrame(rows, columns=columns)
            
            print("Data loaded from database successfully.")
            return df
        except Exception as e:
            print(f"An error occurred while loading data from the database: {str(e)}")
            return None

    def save_to_csv(self, df):
        folder_path = "appData"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
    
        csv_file_path = os.path.join(folder_path, "recordCSV.csv")
        df.to_csv(csv_file_path, index=False)

    
    def update_pause_start_time_2(self):
        try:
            # Get current date and time
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
            current_time = QTime.currentTime().toString("hh:mm:ss")
            current_datetime = current_date + " " + current_time
            
            # Match current user data to find all entries with the same data
            txtID = self.txtID.text()
            cmbxWO = self.txtIssue.currentText()
            
            df = self.load_matching_entries_from_db_2(txtID, cmbxWO, self.cursor)
            
            if df is not None and not df.empty:
                # Get the last index of the DataFrame
                last_index = df.index[-1]
                print("Last index:", last_index)
                
                # Extract the ID from the DataFrame
                entry_id = df.loc[last_index, 'ID']  # Assuming 'ID' is the column name for the primary key
                
                # Save pause data to the pause table
                self.save_pause_data_to_db(entry_id, current_time, current_datetime)
            self.clear_input_fields()
            self.clear_input_fields_2()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"An error occurred while updating pause start time: {str(e)}")

    def calculate_total_time(self, DataID):
        try:
            # Execute SQL query to retrieve the sum of pause duration in seconds for the given DataID
            self.cursor.execute("""
                                SELECT SUM(Pause_Duration_seconds) 
                                FROM Pause 
                                WHERE DataIdx = ?
                                """,
                                (int(DataID),))
            
            # Fetch the result of the query
            result = self.cursor.fetchone()
            
            # Extract the sum of pause duration seconds from the result
            total_pause_time = result[0] if result and result[0] is not None else 0.0  
            print(total_pause_time, 'total_pause_time')
            
            # Execute SQL query to calculate the total time duration in seconds
            self.cursor.execute("""
                                SELECT (strftime('%s', End_Date_Time) - strftime('%s', Start_Date_Time))
                                FROM Data
                                WHERE ID = ?
                                """,
                                (int(DataID),)
            )
            
            # Fetch the result of the query
            result = self.cursor.fetchone()

            # Extract the total time duration in seconds from the result
            total_time_seconds = result[0] if result and result[0] is not None else 0.0  

            # Subtract pause duration from the total time duration
            total_time_seconds -= total_pause_time

            # Execute SQL query to update Total_Time_seconds in the Data table
            self.cursor.execute("""
                                UPDATE Data 
                                SET 
                                    Total_Time_seconds = ? 
                                WHERE 
                                    ID = ?
                                """,
                                (total_time_seconds, int(DataID)))
            
            # Calculate Total_Time_minutes in points
            total_time_minutes = total_time_seconds / 60.0  # Assuming each second is equivalent to 1 point

            # Execute SQL query to update Total_Time_minutes in the Data table
            self.cursor.execute("""
                                UPDATE Data 
                                SET 
                                    Total_Time_minutes = ? 
                                WHERE 
                                    ID = ?
                                """,
                                (total_time_minutes, int(DataID)))
            
            # Commit the transaction to save changes to the database
            self.conn.commit()

        except Exception as e:
            print("An error occurred:", e)

      
        



    
    





        
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
