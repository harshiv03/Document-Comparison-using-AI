PDF to Excel Table Extraction using Azure Form Recognizer
This project provides a simple solution for extracting tables from PDF documents and saving them into an Excel workbook. It utilizes the Azure Form Recognizer service to detect and extract tables from each page of the PDF and then writes these tables to an Excel file using the openpyxl library.

Table of Contents
Overview
Features
Prerequisites
Installation
Usage
Project Structure
License
Overview
In many cases, data is trapped in PDFs, making it difficult to extract and analyze. This project addresses that issue by leveraging Azure's AI capabilities to automatically detect tables within a PDF file and extract them into an easily accessible Excel format.

Features
PDF Splitting: Automatically splits a multi-page PDF into individual pages.
Table Detection: Uses Azure Form Recognizer to identify tables on each PDF page.
Excel Output: Saves extracted tables from each page into separate sheets of an Excel workbook.
Prerequisites
Before running the project, ensure you have the following:

Azure Account: An active Azure account to access Form Recognizer.
Python 3.x: Make sure Python is installed on your machine.
Azure Form Recognizer Resource: You'll need the endpoint and key for the Azure Form Recognizer service.

Usage
Configure Azure Credentials:

Replace the endpoint and key placeholders in the script with your Azure Form Recognizer service credentials.

Prepare the PDF File:

Place the PDF file you want to process in the project directory and update the script with the correct file name.

Run the Script:

Execute the script to extract tables and save them to an Excel workbook:

bash
Copy code
python extract_tables.py
The script will:

Split the PDF into separate pages.
Analyze each page using Azure Form Recognizer to detect tables.
Save the detected tables into an Excel workbook (book.xlsx), with each page's tables on separate sheets.
Project Structure
extract_tables.py: Main script for processing the PDF and extracting tables.
example.pdf: Example PDF file for testing (optional, replace with your file).
book.xlsx: Output Excel file containing extracted tables.
