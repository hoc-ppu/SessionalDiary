# Sessional Diary
Create the House of Commons sessional diary PDF from an Excel file. A python script takes the Excel file as input and creates several XML files. These XML files can be imported into Adobe InDesign templates (.idml files) and then a PDF can be exported.

You can see previously published versions of both the excel files and PDFs on the [UK parliament website](https://www.parliament.uk/business/publications/commons/sessional-diary/).

## Before you start
If you want to create a PDF of a sessional diary from an Excel file you will need to have [Python](https://www.python.org/downloads/) installed and working on your computer and a working recent version of [Adobe InDesign](https://www.adobe.com/products/indesign.html)

You should also clone this repository. [Here is a guide to cloning](https://www.youtube.com/watch?v=CKcqniGu3tA). Or if you do not have git installed you could download and extract the zip file.

## Python script installation
### *Optionally* create and activate a python virtual environment.
To create a virtual environment run the following in PowerShell on Windows or in the terminal on Unix (Mac or Linux).

#### On Windows

Create:
```bash
python -m venv sdenv
```

To activate on Windows, run:
```powershell
sdenv\Scripts\activate.bat
```

#### On Unix

Create:
```bash
python3 -m venv sdenv
```

To activate on Unix, run:
```bash
source sdenv/bin/activate
```

### Install the dependencies (Required)
```bash
pip install -r requirements.txt
```

## Python script usage
First copy or move the Excel file into the same folder as the Sessional_Diary.py python script file.

To run the python script use the following command in your terminal or PowerShell. (replacing input_file_name with your Excel file name):
```bash
python Sessional_Diary.py input_file_name
```

If there are spaces in your Excel file name you may need to use quotes like in the following example:

```bash
python Sessional_Diary.py "2021-22 sessional diary data.xlsx"
```

There are some options you can use to change what is output. Discover them using:
```bash
python Sessional_Diary.py --help
```

The script should output XML files for the 4 main sections of the diary as well as XML for the table of contents. The XML files should appear in the same folder as your input Excel file. *By default* the script will also output an Excel version of the analysis sections.

## InDesign instructions
Open all template .idml files with InDesign. Immediately SaveAs, use a name with the session in it e.g. "sessional-diary-2021-22_Part-1.indd". Close the open .idml files.

You should use an InDesign book as this will give you correct page numbering and make exporting a concatenated PDF easier. Follow the "Create an InDesign book file" and "Add documents to a book file" in [this tutorial](https://redokun.com/blog/indesign-book#toc-3) if you are unsure. Make sure you add all 5 .indd files. You cannot add .idml files to an InDesign book.

Import the XML for parts 1-4 into the relevant InDesign file. See gif below.


![](https://github.com/hoc-ppu/SessionalDiary/blob/main/Import_xml.gif)


If the text does not flow to multiple pages automatically, follow the instructions in this [youtube video](https://youtu.be/jUP1kMsIYV0?t=97) (from about 1:37 on).

For the contents use InDesign's built in Table of Contents feature. Then turn this into a table. Add new text frames off the page. Import XML for the contents into these off page text frames. Add additional columns to the contents table that was generated by indesign. Copy the relevant columns from the tables you just imported via XML into the contents table. Format to match previous style.

To output the PDF, use the 'Export Book to PDF...' option in the InDesign book panel menu.

