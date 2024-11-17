
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx_write new]
$docx paragraph "My first heading" -style Heading 
$docx paragraph "My first paragraph "
$docx appendToParagraph "build in parts" -bold on
$docx appendToParagraph " serveral parts indeed" -italic on

$docx write testout.docx

