#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx
set docx [docx new]
$docx import styles docx-example4-in.docx
$docx paragraph "Heading" -pstyle "Heading"
$docx paragraph $loreipsum

$docx write docx-example4.docx
$docx destroy
