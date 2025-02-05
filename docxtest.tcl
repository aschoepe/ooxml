
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx_write new]
$docx paragraph "My first heading" -style Heading 
$docx paragraph [string repeat "My first paragraph " 50]
$docx append "build in parts" -bold on
$docx append " serveral parts indeed" -italic on
$docx append " and different fonts" -font "Utopia"
$docx simpletable {
    {1 2 3}
    {aaaaaa bbbbb ccccc}
    {"foo bar" "grill baz" "lore ipsum"}
}
$docx style paragraph Mystyle -font Utopia -bold 1 -italic true \
    -spacing {before 120 after 60}
$docx paragraph "Another paragraph with its own style" -style Mystyle
$docx paragraph [string repeat "Next paragraph, back to default style, with local changes" 20] -spacing {line 400}
puts [$docx style ids]
$docx write testout.docx
$docx destroy

