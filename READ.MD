# File Converter and Merger
This Python application provides a graphical interface to convert .docx files to PDF and merge multiple PDF files into one. It uses the tkinter library for the GUI, PyPDF2 for merging PDFs, and docx2pdf for converting .docx files to PDF.

## Features
* Convert .docx files to PDF.
* Merge multiple PDF files into a single PDF file.
* User-friendly interface to select files and specify output locations.

## Prerequisites
* Python 3.x
* tkinter library (comes pre-installed with Python)
* PyPDF2 library
* docx2pdf library

## Installation
* Install the required Python libraries: pip install PyPDF2 docx2pdf
* Save the Python script to a file, for example, file_converter_and_merger.py.

## Usage
* Run the script: python file_converter_and_merger.py

## Interface Overview
* File Conversion Section (Left Panel):
1) Select Word Files: Choose .docx files you want to convert to PDF.
2) Convert to PDF: Convert the selected .docx files to PDF.

* PDF Merge Section (Right Panel):
1) Select PDF Files: Choose PDF files you want to merge.
2) Output PDF file name: Specify the name and location of the merged PDF file.
3) Merge PDF Files: Merge the selected PDF files into a single PDF.

## Example
* Here is an example of how to use the application:
1) Run the script.
2) In the left panel, click "Select Word Files" to choose .docx files for conversion.
3) Click "Convert to PDF" to convert the selected .docx files to PDF.
4) In the right panel, click "Select PDF Files" to choose PDF files for merging.
5) Enter the output file name in the entry box or click "Save As..." to specify the location.
6) Click "Merge PDF Files" to merge the selected PDF files into one.