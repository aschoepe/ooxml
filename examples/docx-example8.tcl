#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx

set docx [docx new -creator foo -created 2025-02-10]

$docx pagesetup -fn_numFmt lowerLetter

$docx paragraph $loreipsum
$docx footnote -verticalAlign subscript {
    $docx paragraph " My first footnote."
}
$docx append " $loreipsum"

$docx paragraph $loreipsum
$docx footnote -verticalAlign subscript {
    $docx paragraph " My "
    $docx append "second" -underline double
    $docx append " footnote."
}

$docx write docx-example8.docx
$docx destroy

