import sys
import requests
import os
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QDateEdit, \
    QFileDialog
from PyQt5.QtCore import QDate


class AttendanceDownloaderApp(QWidget):
    """
    A PyQt5 application to download attendance data from an API.
    """

    def __init__(self):
        """Initializes the application UI and state."""
        super().__init__()
        self.initUI()
        self.save_path = "" # Variable to store the save path

    def initUI(self):
        """Sets up the graphical user interface."""
        self.setWindowTitle("Attendance Downloader")
        self.setGeometry(100, 100, 400, 300)
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
                font-family: 'Helvetica Neue', sans-serif;
                border-radius: 8px;
            }
            QLabel {
                font-size: 16px;
                color: #333;
            }
            QLineEdit, QDateEdit {
                padding: 10px;
                border: 2px solid #007BFF;
                border-radius: 5px;
                font-size: 14px;
                background-color: #fff;
            }
            QLineEdit:focus, QDateEdit:focus {
                border: 2px solid #0056b3;
            }
            QPushButton {
                background-color: #007BFF;
                color: white;
                padding: 10px;
                border: none;
                border-radius: 5px;
                font-size: 16px;
                cursor: pointer;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004494;
            }
        """)

        # Class number input
        self.class_label = QLabel("Class Number:")
        self.class_input = QLineEdit(self)
        self.class_input.setPlaceholderText("Enter class number (e.g., 10A)")

        # Date picker input
        self.date_label = QLabel("Select Date (optional, format: DD-MM-YYYY):")
        self.date_picker = QDateEdit(self)
        self.date_picker.setDisplayFormat("dd-MM-yyyy")
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDate(QDate.currentDate())

        # Save file button
        self.save_button = QPushButton("Choose Save Location", self)
        self.save_button.clicked.connect(self.choose_save_location)

        # Download button
        self.download_button = QPushButton("Download Attendance", self)
        self.download_button.clicked.connect(self.download_attendance)

        # Layout arrangement
        layout = QVBoxLayout()
        layout.addWidget(self.class_label)
        layout.addWidget(self.class_input)
        layout.addWidget(self.date_label)
        layout.addWidget(self.date_picker)
        layout.addWidget(self.save_button)
        layout.addWidget(self.download_button)

        self.setLayout(layout)

    def choose_save_location(self):
        """Opens a file dialog for the user to select the save path."""
        # The file dialog now suggests a default filename based on current date
        from datetime import datetime
        current_date_str = datetime.now().strftime("%Y-%m-%d")
        default_filename = f"Attendance-{current_date_str}.xlsx"
        
        self.save_path, _ = QFileDialog.getSaveFileName(self, "Save File", default_filename, "Excel Files (*.xlsx);;All Files (*)")
        if self.save_path:
            QMessageBox.information(self, "Save Location Selected", f"File will be saved to: {self.save_path}")

    def download_attendance(self):
        """
        Validates user input and initiates the download process.
        Shows a success or error message upon completion.
        """
        class_number = self.class_input.text().strip()
        date = self.date_picker.text().strip()

        if not class_number:
            QMessageBox.warning(self, "Input Error", "Class number is required!")
            return

        if not self.save_path:
            QMessageBox.warning(self, "Save Location Error", "Please choose a save location!")
            return

        try:
            self.fetch_and_save_excel(class_number, date)
            QMessageBox.information(self, "Success", "Attendance downloaded successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def fetch_and_save_excel(self, class_number, date=None):
        """
        Fetches the attendance data from the API and saves it to a file.
        This function has been updated to use the correct URL format.
        """
        # The base URL now includes the class number directly in the path,
        # as expected by the API.
        url = f"https://iec-attendance-nodejs.onrender.com/faculty/dayExcel/{class_number}"
        
        params = {}
        # Only add the 'date' as a query parameter if it is provided.
        if date and date.strip():
            params["date"] = date

        response = requests.get(url, params=params)

        if response.status_code == 200:
            # Ensure the directory exists before saving the file
            save_directory = os.path.dirname(self.save_path)
            if save_directory and not os.path.exists(save_directory):
                os.makedirs(save_directory)

            with open(self.save_path, 'wb') as f:
                f.write(response.content)
        else:
            try:
                # Attempt to get a JSON error message from the response
                error_data = response.json()
                error_message = error_data.get('error', f'Unknown error occurred: {response.status_code}')
                raise Exception(f"API Error: {error_message}")
            except requests.exceptions.JSONDecodeError:
                # If JSON parsing fails, provide a generic error message
                raise Exception(f"Failed to download. Status Code: {response.status_code}")


if __name__ == '__main__':
    # This block ensures the application runs when the script is executed
    app = QApplication(sys.argv)
    ex = AttendanceDownloaderApp()
    ex.show()
    sys.exit(app.exec_())
