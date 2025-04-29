# Origin file conversion to open formats

This Python tool can be used to convert OriginLab (.opju) file contents 
to .csv, .txt and .png files.

## How to cite this tool

If you use this tool, please cite the following source:

Niskanen, J (2025). JNisk/convert-opju (v1.0). Zenodo.
https://doi.org/10.5281/zenodo.15300716

## Table of contents

- Requirements
- Installation
- Usage
- Known issues

## Requirements

This Python tool has been tested with Origin 2024b and Python 3.11.0 
in Windows 11. Compatibility with other versions of Origin, Python or 
Windows is not guaranteed.

Software required:
 - Windows 11
 - Origin 2024b
 - Python 3.11.0

Python packages required:
 - originpro==1.1.10
 - numpy==2.2.3
 - pandas==2.2.3
 - pillow==11.1.0
 - PyGetWindow==0.0.9
 - PyRect==0.2.0
 - python-dateutil==2.9.0.post0
 - pytz==2025.1
 - six==1.17.0
 - tzdata==2025.1

## Installation

This tool does not need to be separately installed; it can simply be 
downloaded and run. To acquire a copy of this tool, you can download 
it from this repository or use git clone:

`git clone https://github.com/JNisk/convert-opju.git`

To install the required packages in Origin 2024b, go to Connectivity > 
Python Packages > Install and enter the package names. For example, to 
install PyGetWindow 0.0.9, enter `PyGetWindow==0.0.9` and click OK.

## Usage

To use this tool, open your .opju file in Origin 2024b and go to 
Window > Script Window. In the script window, write the following
command, where `(script)` needs to be replaced with the full path of 
the `convert_opju.py` file:

`run -pyf (script)`

To run the command, press Enter. The tool will report the type and 
path of objects converted. After the run has finished, Origin 2024b 
can be closed. Due to temporary windows created during the conversion 
process, Origin 2024b may prompt to save any changes, but the file can 
be closed without saving as the file objects have not been modified. 

## Output

The following objects in an .opju file will be converted:
- Workbook -> .csv (one file per worksheet)
- Matrix -> .csv (one file per matrix object)
- Graph -> .png
- Image -> .png
- Notes (Text, Origin Rich Text) -> .txt
- Notes (Markdown) -> .md
- Notes (HTML) -> .html

The objects will be saved with the same name as in the .opju file, 
with the exception that any forward or backward slashes in the 
object's name will be converted to underscore.

The files will be placed in a folder named after the .opju file, and 
subfolder structure of the file is maintained in the output. For 
example, if the Origin file is named `MyProject.opju`, the output 
files are placed in a folder named `MyProject`. If the folder already 
exists, the tool will ask the user if files should be overwritten or 
not.

## Known issues

In worksheet with merged cells, only the contents of the merged cell 
may be visible in Origin 2024b. However, the content entered in 
adjacent cells is retained in the .opju file despite not being 
visible, and this content is exported by Python into the .csv output.
