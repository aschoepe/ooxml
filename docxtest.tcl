
source ./ooxml.tcl
source ./ooxml-docx.tcl

set docx [::ooxml::docx::docx new]
# foreach type {paragraph character} {
#     foreach styleid [$docx style ids $type] {
#         $docx style delete $type $styleid
#     }
# }
$docx pagesetup -sizeAndOrientaion {width 16838 height 23811 orientation landscape} -margins {left 2000 right 2000} -paperSource {first 1 other 2}
$docx style characterdefault -fontsize 20 -font "Liberation Serif"
$docx style paragraph Heading1 -fontsize 28 -bold on
$docx paragraph "My first heading" -style Heading1 -spacing {before 240 after 120}
$docx paragraph [string repeat "My first paragraph " 50] -spacing {after 120}
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
$docx paragraph [string repeat "Next paragraph with stupid text. " 20] \
    -indentation {firstLine 3cm hanging 2cm start 1cm end 4cm}
$docx url "click this" "https://www.staatstheater-stuttgart.de/" -underline single
$docx url "and click that" "https://core.tcl-lang.org/" -underline single
$docx append " more text"
$docx paragraph "A new paragraph"
# Prevent error if someone follow this actually testing
if {[file exists book.jpg]} {
    #$docx image book.jpg -dimension {width 2872105 height 2872105}
    $docx image book.jpg -dimension {width 3cm height 3cm}
} else {
    puts "Missing book.jpg to include into docx"
}
puts "paragraph style ids: [$docx style ids paragraph]"
puts "character style ids: [$docx style ids character]"
$docx write testout.docx
$docx writepart word/styles.xml document.xml
$docx readpart word/styles.xml document.xml
$docx write testout1.docx
$docx destroy

