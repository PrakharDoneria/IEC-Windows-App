import sys
import requests
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QDateEdit, \
    QFileDialog
from PyQt5.QtCore import QDate


class AttendanceDownloaderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
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

        # Variable to store the save path
        self.save_path = ""

    def choose_save_location(self):
        self.save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx);;All Files (*)")
        if self.save_path:
            QMessageBox.information(self, "Save Location Selected", f"File will be saved to: {self.save_path}")

    def download_attendance(self):
        class_number = self.class_input.text()
        date = self.date_picker.text()

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
        base_url = "https://iec-group-of-institutions.onrender.com/raw-excel"
        params = {"class_number": class_number}

        if date and date.strip():
            params["date"] = date

        response = requests.get(base_url, params=params)

        if response.status_code == 200:
            filename = f"Attendance-{class_number}"
            if date and date.strip():
                filename += f"-{date}.xlsx"
            else:
                from datetime import datetime
                current_date = datetime.now().strftime("%d-%m-%Y")
                filename += f"-{current_date}.xlsx"

            # Combine save path with filename
            full_save_path = self.save_path if self.save_path.endswith(
                '.xlsx') else self.save_path + f"/{filename}.xlsx"

            with open(full_save_path, 'wb') as f:
                f.write(response.content)
        else:
            try:
                error_message = response.json().get('error', 'Unknown error occurred')
                raise Exception(f"Error: {error_message}")
            except:
                raise Exception(f"Failed to download. Status Code: {response.status_code}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AttendanceDownloaderApp()
    ex.show()
    sys.exit(app.exec_())
