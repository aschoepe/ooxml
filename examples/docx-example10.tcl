#!/bin/sh
#\
exec tclsh "$0" "$@"

cd [file dirname [info script]]
source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx
set docx [docx new]

featuresCovered "The replace method."
set docxname [file join [file dir [info script]] docx-example10]
$docx paragraph $loreipsum
$docx simplecomment $loreipsum
$docx footnote {
    $docx paragraph $loreipsum
}

$docx write $docxname
$docx destroy

set docx [docx new $docxname.docx]
$docx replace e !

$docx write $docxname-1
$docx destroy

set docx [docx new $docxname.docx]
$docx replace e ! {word/document.xml word/footnotes.xml}

$docx write $docxname-2
$docx destroy
