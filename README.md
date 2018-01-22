Outlook MSG Converter

This script will convert MS Outlook .msg files to .txt for analysis, preserving headers and body

PARAMETER -f
    Specifies the file for analysis if it exists in the current working directory
    
PARAMETER -p
    Specifies an absolute path to the file for analysis
    
PARAMETER -o
    Specifies a path for the generated analysis file
    
EXAMPLE
    PS C:\Users\Bob\Desktop> .\MSG_Convert.ps1 -f test.msg
    
    This will convert test.msg on Bob's Desktop and output to his Desktop
