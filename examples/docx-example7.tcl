#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

cd [file dirname [info script]]
source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx

set docx [docx new -creator foo -created 2025-02-10]

featuresCovered "Fields (page number, author etc.), document configuration."

set fields {
    author
    createdate
    date
    filesize
    page
    savedate
    section
    time
    title
    username
}
  
$docx sectionstart
$docx paragraph $loreipsum\n
foreach field $fields {
    $docx append "\n"
    $docx append "$field "
    $docx field $field
}
$docx append $loreipsum\n
$docx sectionstart
$docx paragraph "$loreipsum"
$docx field PAGE
$docx field SECTION
$docx pagebreak
$docx paragraph "$loreipsum"
$docx field page
$docx append " "
$docx field date "DATE \\@ \"yyyy-MM-dd\"" -bold 1 -color 0B24F8

$docx configure -creator bar -title "ooxml-docx Example File"

$docx write docx-example7.docx
$docx destroy

