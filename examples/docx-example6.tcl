#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
namespace import ::ooxml::docx::docx

set docx [docx new]
$docx paragraph "A new paragraph"
$docx image book.jpg -dimension {width 3cm height 3cm} -bwMode black
$docx paragraph "A new paragraph"
$docx image book.jpg -dimension {width 3cm height 3cm} -bwMode black

$docx write docx-example6.docx
$docx destroy
