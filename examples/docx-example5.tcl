#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
namespace import ::ooxml::docx::docx

set docx [docx new]
$docx import numbering docx-example4-in.docx
$docx paragraph "First" -styleNr 1 
$docx paragraph "Second" -styleNr 1 -level 0
$docx paragraph "Sub Second frist" -styleNr 1 -level 1
$docx paragraph "Sub Second second" -styleNr 1 -level 1
$docx paragraph "Third" -styleNr 1 -level 0

$docx write docx-example5.docx
$docx destroy
