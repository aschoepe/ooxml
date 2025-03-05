#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
namespace import ::ooxml::docx::docx

set docx [docx new]
# Importing numbering definition from another docx
$docx import numbering docx-example4-in.docx

# Some other random paragraph style definition
$docx style paragraph mystyle -color 00FF00 -fontsize 8pt

# Use the imported style
$docx paragraph "First" -numberingStyle 1 
$docx paragraph "Second" -numberingStyle 1 -level 0
$docx paragraph "Sub Second frist" -numberingStyle 1 -level 1
$docx paragraph "Sub Second second" -numberingStyle 1 -level 1 -style mystyle
$docx paragraph "Third" -numberingStyle 1 -level 0

# Add some space
$docx paragraph ""
$docx paragraph ""

# Create a numbering definition from scratch
$docx numbering abstractNum 3 {
    {
        -numberFormat upperRoman
        -levelText "%1."
        -start 3
        -indentation {start 1cm hanging 0.5cm}
        -fontsize 18pt
    }
    {
        -numberFormat chicago
        -levelText "%2"
        -indentation {start 2cm hanging 0.5cm}
        -fontsize 8pt
    }
}
# Use it
$docx paragraph "First" -numberingStyle 3
$docx paragraph "Second" -numberingStyle 3 -level 0
$docx paragraph "Sub Second frist" -numberingStyle 3 -level 1
$docx paragraph "Sub Second second" -numberingStyle 3 -level 1 -style mystyle
$docx paragraph "Third" -numberingStyle 3 -level 0


$docx write docx-example5.docx
$docx destroy
