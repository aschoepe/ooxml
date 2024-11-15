
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx_write new]
$docx text "My first heading" -style Berschrift1
$docx text "My first paragraph "
$docx appendText "build in parts" -bold on
$docx appendText " serveral parts indeed" -italic on

$docx write testout.docx

