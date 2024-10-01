# Attendance Downloader

A Python application for downloading attendance records as Excel files. This tool allows users to fetch attendance data for specific classes and dates, providing a user-friendly graphical interface built with PyQt5.

## Features

- Download raw attendance data as Excel files.
- Filter attendance by class number and optional date.
- Modern and aesthetic user interface.
- Easy-to-use file save location selection.

## Requirements

- Python 3.x
- PyQt5
- Requests

## Installation

1. **Clone the repository** (if using Git):
   ```bash
   git clone https://github.com/PrakharDoneria/IEC-Windows-App.git
   cd https://github.com/PrakharDoneria/IEC-Windows-App.git
   ```

2. **Install the required packages**:
   Ensure you have the required libraries installed. You can use pip to install them:
   ```bash
   pip install PyQt5 requests
   ```

3. **Run the Application**:
   To run the application, execute the following command:
   ```bash
   python main.py
   ```

## Usage

1. **Open the Application**: Launch the application by executing the main Python script.

2. **Enter Class Number**: Input the class number (e.g., `10A`) in the designated field.

3. **Select Date (Optional)**: Use the date picker to select a date. The expected format is `DD-MM-YYYY`. If left blank, attendance will be fetched for the current date.

4. **Choose Save Location**: Click the "Choose Save Location" button to specify where to save the downloaded Excel file.

5. **Download Attendance**: Click the "Download Attendance" button to fetch and save the attendance data. A success message will be displayed upon successful download.

## Creating an Executable

To create an executable file for Windows using PyInstaller, follow these steps:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Navigate to the project directory:
   ```bash
   cd C:\Users\lenovo\PycharmProjects\iecFrontendWindows
   ```

3. Run the PyInstaller command:
   ```bash
   pyinstaller --onefile --windowed main.py
   ```

4. After the build completes, find the executable in the `dist` folder.

## Error Handling

- If the date format is invalid, the application will prompt an error message.
- If no attendance records are found for the specified class or date, the application will inform the user accordingly.

## Acknowledgments

- Thanks to the developers of PyQt5 and Requests for their excellent libraries that make this application possible.
