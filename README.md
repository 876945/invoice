ProInvoice Generator
A modern professional invoice generator desktop application built with Python.
Features:


Create professional invoices


Export invoices as PDF


Export unpaid invoices to Excel


Customer management


Analytics dashboard


Company logo support


SQLite database storage


Dark modern UI



Requirements
Install Python libraries before running the project.
pip install reportlab pillow openpyxl matplotlib
Tkinter and SQLite are included with Python by default.

Run the Application
python main.py

Create Virtual Environment (Recommended)
Windows
python -m venv venvvenv\Scripts\activate
Install Dependencies
pip install -r requirements.txt

Generate requirements.txt
pip freeze > requirements.txt

Build EXE File
Install PyInstaller:
pip install pyinstaller
Create executable:
pyinstaller --onefile --windowed main.py
Generated EXE file will be inside:
dist/

Features


Professional invoice PDF generation


Customer database


Invoice tracking


Paid/Unpaid status


Revenue analytics


Monthly charts


Excel export


Company settings


Logo upload system



Technologies Used


Python


Tkinter


SQLite


ReportLab


OpenPyXL


Matplotlib


Pillow


