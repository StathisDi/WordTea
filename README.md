# WordTea
Python script for parsing word .docx files and automatically create the references for sections and 2 levels of subsections, figures, tables and citations. The references need to be declared inside the text in a specific format.
This tool enables latex like cross-reference capabilities inside MS Word documents without the use of Add-Ins.

## Declarations
Create a label for the section, figures, etc. using the following format:

### Section
**^sec1{`< label >`}^**

### LVL 1 subsections
**^sec2{`< label >`}^**

### LVL 2 subsections
**^sec3{`< label >`}^**

### Figures
**^fig{`< label >`}^**

### Tables
**^tbl{`< label >`}^**

### Equations
**^eq{`< label >`}^**

### Citations
**^cite{`< label >`}^**

Using the above declarations a table of all the references is created and printed on the command window.


**NOTICE**

The numbering follows the same order as the declarations, the declarations could take place anywhere in the text (as long as it have the correct order).

## Referencing

To create a cross-reference inside the text, a reference need to be created. To do that the following format has to be inserted in the text. The **\`** symbol is part of the format and has to be included in the text

### For Figures
- **\`fig{`< label >`}\`**

### For Sections and Subsections
- **\`sec1{`< label >`}\`**
- **\`sec2{`< label >`}\`**
- **\`sec3{`< label >`}\`**

### For Tables
- **\`tbl{`< label >`}\`**

### For Equations
- **\`eq{`< label >`}\`**

### For citations
- **\`cite{`< label >`}\`**


## Running the script

Download all three files from the src folder.

**!!!All three files should be in the same folder for the script to run!!!**

To run the script run the following command:

`$ python WordTea.py <source .docx> <destination .pdf> <temporary .docx file>`

Use `$ python WordTea.py --help` for details on extra options.

Some available options:
- -s1       : Format of the section 1 reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- -s2       : Format of the section 2 reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- -s3       : Format of the section 3 reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- -table    : Format of the table reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- --verbose : Enable Verbose level 2, extreme error print for script debug.
- --silent  : Disable Verbose level 1, basic status print for missed references and citations inside the document.

### Note:
1. The external file option is not supported and it is under development.
2. To run the script, all three files must be in the same folder!

## Known bugs and issues:

1. The script removes footnotes, so the have to be added again later in the post-processed temporary docx file.
2. The script sometimes get stack in the saving of the pdf or docx file. In that case close the terminal and terminate the Microsoft-word from the task manager.

## TODOs
- Add a try-catch or other error handling code around the saving of the file to fix bug (2).
- Fix the external file options.
- Create a utility class that will include all the utility functions.
- Find out why the footnotes are removed.

## Requirements
In order to run the script there are the following requirements.
- Python 3
- comtypes package
- python-docx package

The python packages can be installed by using **pip**

To install in windows use pip3.exe install <package name> from a command window. For mac or linux you can use the pip3 command from the terminal.

# Version Details:
 Author &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : Dimitrios Stathis </br>
 email &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : stathis@kth.se, sta.dimitris@gmail.com </br>
 Last edited : 21/12/2019 </br>
 version &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: 2.0</br>
 &copy; Copyright 2017, All rights reserved.

# LICENSE

WordTea is free software: you can redistribute it and/or modify   
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or   
(at your option) any later version.                                 

WordTea is distributed in the hope that it will be useful,        
but WITHOUT ANY WARRANTY; without even the implied warranty of      
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the       
GNU General Public License for more details.                 

You should have received a copy of the GNU General Public License   
along with WordTea, see `COPYING` file.  If not, see <https://www.gnu.org/licenses/>.
