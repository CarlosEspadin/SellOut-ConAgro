# SellOut-ConAgro
This repository contains the source code and executable file of the GUI project that cleans, transforms and uploads the SellOut information from Syngenta distributors to the ConAgro web service.
# Foobar

This Python application is designed to manually upload distributor information to the ConAgro service. It provides a graphical user interface (GUI) for users to input the path to an Excel file and a distributor number, processes the data according to the distributor's layout, and then uploads the data to a web service in JSON format.

## Installation

Ensure you have Python 3.11 installed on your system. You can then install the required dependencies using the following commands:

```bash
pip install pandas==1.5.3
pip install numpy==1.26.0
pip install requests==2.31.0
pip install openpyxl==3.1.2
```
Note: openpyxl is a library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

## Usage

To use this application, run the script in a Python environment. The GUI will prompt you to enter the path to the Excel file containing the distributor's data and the distributor's number.
Start the application.
Enter the path to the Excel file when prompted.
Enter the distributor number when prompted.
Click the 'Aceptar' button to process and upload the data.
The application will process the data, convert it to JSON, and upload it to the ConAgro web service.

## Contributing

Contributions to this project are welcome. To contribute:
1. Fork the repository.
2. Create a new branch for your feature (git checkout -b feature/fooBar).
3. Commit your changes (git commit -am 'Add some fooBar').
4. Push to the branch (git push origin feature/fooBar).
5. Create a new Pull Request.
