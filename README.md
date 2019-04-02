# DocParser
Python script for parsing docx files and automatically create the references for sections and 2 levels of subsections, figures, tables and citations. The references need to be declared inside the text in a specific format.

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


To run the script run the following command:

`$ python Doc_parse_v2 <source .docx> <destination .pdf> <temporary .docx file>`

Use `$ python Doc_parse_v2 --help` for details on extra options.

Some available options:
- -s1       : Format of the section 1 reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- -s2       : Format of the section 2 reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- -s3       : Format of the section 3 reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- -table    : Format of the table reference style. Use 1 for normal numbering, 2 for Latin, 3 for small letter, 4 for capital letter, default 1.
- --verbose : Enable Verbose level 2, extreme error print for script debug.
- --silent  : Disable Verbose level 1, basic status print for missed references and citations inside the document.

### Note:
The external file option is not fully supported and it is under development.

## Requirements 
In order to run the script there are the following requirements.
- Python 3
- comtypes package
- python-docx package

The python packages can be installed by using **pip**

To install in windows use pip3.exe install <package name> from a command window. For mac or linux you can use the pip3 command from the terminal.


# LICENSE

DocParser is free software: you can redistribute it and/or modify   
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or   
(at your option) any later version.                                 
                                                                    
DocParser is distributed in the hope that it will be useful,        
but WITHOUT ANY WARRANTY; without even the implied warranty of      
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the       
GNU General Public License for more details.                 
                                                                    
You should have received a copy of the GNU General Public License   
along with DocParser, see `COPYING` file.  If not, see <https://www.gnu.org/licenses/>. 