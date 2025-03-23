#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx

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
  
set docx [docx new -creator foo -created 2025-02-10]
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

$docx configure -creator bar -title "ooxml-docx Example File"

$docx write docx-example7.docx
$docx destroy

