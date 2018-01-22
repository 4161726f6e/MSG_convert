<#
.SYNOPSIS
    Outlook MSG Converter
.DESCRIPTION
    This script will convert MS Outlook .msg files to .txt for analysis, preserving headers and body
.PARAMETER -f
    Specifies the file for analysis if it exists in the current working directory
.PARAMETER -p
    Specifies an absolute path to the file for analysis
.PARAMETER -o
    Specifies a path for the generated analysis file
.EXAMPLE
    PS C:\Users\Bob\Desktop> .\MSG_Convert.ps1 -f test.msg
    This will convert test.msg on Bob's Desktop and output to his Desktop
.NOTES
    Author: https://github.com/4161726f6e
    Date:   12/19/2017    
#>


Param( 
        [string]$p = "",
        [string]$f = "",
        [string]$o = ""
        )

# Error handling if neither -p or -f were defined
If ($p -eq "" -and $f -eq ""){
    write-warning "The specified input file could not be found: $checkFile"
    Exit 
}


# Determine if the user defined a path to the file for analysis
If ($p -eq ""){
    $checkFile = $f                  
    $fname = $checkFile
    } Else {
    $checkFile = $p                
    $fname = Split-Path $checkFile -leaf
}

# Error handling if file for analysis DNE
If(!(test-path $checkFile)){
    write-warning "The specified input file could not be found: $checkFile"
    Exit
}

# Invoke COM to utilize Outlook for .msg file manipulation
$ol = New-Object -ComObject Outlook.Application
If ($p -eq ""){
    $path = (Get-Item -Path ".\" -Verbose).FullName
    $msg = $ol.CreateItemFromTemplate($path + "\" + $checkFile)
    } Else {
    $fname = Split-Path $checkFile -leaf
    $msg = $ol.CreateItemFromTemplate($checkFile)
}

# If user defined output file path, write to $outPath and add trailing "/" if it DNE
If ($o -eq ""){
    $outPath = $path
    } Else {
    $outPath = $o
    if($outPath[-1] -ne '\') {
        $outPath += '\'
    }
}

# Create user-defined output directory if it DNE
If(!(test-path $outPath)){
    New-Item -ItemType Directory -Force -Path $outPath
}

# Define output analysis file name
$outFile = $outPath + $fName + "_for_analysis.txt"

# Write email headers to analysis file
$headers = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
$headers >> $outFile

# write email body to analysis file
$mystring = $msg.body
$mystring >> $outFile