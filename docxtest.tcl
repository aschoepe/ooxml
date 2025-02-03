
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx_write new]
$docx paragraph "My first heading" -style Heading 
$docx paragraph "My first paragraph "
$docx append "build in parts" -bold on
$docx append " serveral parts indeed" -italic on
$docx append " and different fonts" -font "Utopia"
$docx simpletable {
    {1 2 3}
    {aaaaaa bbbbb ccccc}
    {"foo bar" "grill baz" "lore ipsum"}
}
$docx style paragraph Mystyle -font Utopia -bold 1 -italic true
$docx paragraph "Another paragraph with its own style" -style Mystyle
puts [$docx style ids]
$docx write testout.docx
$docx destroy

