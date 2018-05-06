# DocParser
Python script for parsing docx files and automatically create the references for sections and 3 levels of subsections, figures, tables and citations. The references need to be declared inside the text in a specific format.

## Declarations
Create a label for the section, figures, etc. using the following format:

### Section
**^sec0{`< label >`}^**

### LVL 1 subsections
**^sec1{`< label >`}^**

### LVL 2 subsections
**^sec2{`< label >`}^**

### LVL 3 subsections
**^sec3{`< label >`}^**

### Figures
**^fig{`< label >`}^**

### Tables
**^tbl{`< label >`}^**

### Citations
**^cite{`< label >`}^**

Using the above declarations a table of all the references is created and printed on the command window.


**NOTICE**

The numbering follows the same order as the declarations, the declarations could take place anywhere in the text (as long as it have the correct order).

## Referencing

To create a reference inside the text, a reference need to be created. To do that the following format is needed to be inserted in the text.

### For sections, subsections, figures and tables
**ref`< label >`**

### For citations
**cit`< label >`**

Those should be written as one word.

## Running the script

In order to run the script there are the following requirements.
- Python 3
- comtypes package
- python-docx package

The python packages can be installed by using **pip**

To run the script run the following command:

`$ python Doc_parse_v2 <source .docx> <destination .pdf> <temporary .docx file>`


## Packages needed
- comtypes
- python-docx

To install in windows use pip3.exe install <package name> from a command window. For mac or linux you can use the pip3 command from the terminal.
