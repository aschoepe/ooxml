
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx_write new]
$docx text "My first paragraph"
$docx text "My second paragraph"
$docx write testout.docx

