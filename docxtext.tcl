
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx_write new]
$docx text "My first heading" -style Berschrift1
$docx text "My second paragraph "
$docx appendText "build in parts"
$docx appendText " serveral parts indeed"

$docx write testout.docx

