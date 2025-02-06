
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx new]
foreach type {paragraph character} {
    foreach styleid [$docx style ids $type] {
        $docx style delete $type $styleid
    }
}
$docx style characterdefault -fontsize 20 -font "Liberation Serif"
$docx style paragraph Heading1 -fontsize 28 -bold on
$docx paragraph "My first heading" -style Heading1 -spacing {before 240 after 120}
$docx paragraph [string repeat "My first paragraph " 50]
$docx append "build in parts" -bold on -fontsize 16
$docx append " serveral parts indeed" -italic on
$docx append " and different fonts" -font "Utopia"
$docx style character myCharacterStyle -bold on -italic true -font "Utopia"
$docx append " and a part with character style" -style myCharacterStyle
$docx simpletable {
    {1 2 3}
    {aaaaaa bbbbb ccccc}
    {"foo bar" "grill baz" "lore ipsum"}
}
$docx style paragraph Mystyle -font Utopia -bold 1 -italic true \
    -spacing {before 120 after 60} -color ff0000 -underline wave
$docx paragraph "Another paragraph with its own style" -style Mystyle
$docx paragraph [string repeat "Next paragraph, back to default style, with a few local changes. " 20] -spacing {line 400} -align center
$docx paragraph "Another paragraph with its own style (local applied)" -align right
$docx style paragraph RigthAlign -align right
$docx paragraph "Another paragraph with its own style (applied by style)" -style RigthAlign
puts "paragraph style ids: [$docx style ids paragraph]"
puts "character style ids: [$docx style ids character]"
$docx write testout.docx
$docx writepart word/document.xml document.xml
$docx readpart word/document.xml document.xml
$docx write testout1.docx
$docx destroy

