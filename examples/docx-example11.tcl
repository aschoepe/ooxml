#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx
set docx [docx new]

$docx paragraph "Some text "
$docx commentstart mycommentid
$docx append "read more"
$docx commentstart mycommentid2
$docx append " about this"
$docx commentend $mycommentid -author foo {
    $docx paragraph "What do you mean?"
}
$docx append " even more text"
$docx commentend $mycommentid2 -author bar {
    $docx paragraph "Two "
    $docx append "overlapping" -italic on
    $docx append " comments"
}
$docx append " and again even more text "
$docx commentstart mycommentid3
$docx append "and again even more text"
$docx simplecommentend $mycommentid3 "A bit repititive?" -bold on -author bar -initials b -date "2025-11-15T03:00:00"
$docx append " and again even more text "



$docx write docx-example11
$docx destroy
