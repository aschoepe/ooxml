#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx
set docx [docx new]

# The rId returned by the header/footer methods is needed later. You
# may store it like:
set defaultHeader [$docx header {
    $docx field page
}]

# Or you use the optional last argument pointing to a variable to
# store the rId into:
$docx header {
    $docx paragraph "even page " -align end
    $docx field page
} evenHeader
$docx footer {
    $docx paragraph "footer default " -align center
    $docx field page
    $docx append " "
    $docx image book.jpg inline -dimension {width 1cm height 1cm}
    $docx append " "
    $docx field numpages
} defaultFooter

# Set a default style
$docx style paragraphdefault -spacing {before 240 after 120}

# And a heading style
$docx style paragraph Heading1 -fontsize 28 -spacing {before 480 after 240}

# Set the margins
#$docx pagesetup 
$docx pagesetup -sizeAndOrientaion {width 15840 height 12240} \
    -margins {left 1cm right 1cm top 1cm bottom 1cm} \
    -topBorder {type dashed borderwidth 10} \
    -defaultHeader $defaultHeader \
    -evenHeader $evenHeader \
    -firstHeader $defaultHeader \
    -defaultFooter $defaultFooter

$docx paragraph "Chapter one" -pstyle Heading1
# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"

# Start a new section with different page setup
$docx sectionstart -sizeAndOrientaion {width 12240 height 15840} \
    -margins {left 3cm right 3cm top 4cm bottom 4cm} \
    -leftBorder {type thick borderwidth 20 color 00ff00 space 20} \
    -pageNumbering {fmt upperRoman start 222}

$docx paragraph "Chapter two" -pstyle Heading1
# Add some paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
for {set i 0} {$i < 8} {incr i} {
    $docx paragraph "The next paragraph: $loreipsum"
}

# End the section
$docx sectionend

# We are back to the pagesetup
$docx paragraph "Chapter three" -pstyle Heading1
# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"
$docx pagebreak
$docx paragraph "The next paragraph: $loreipsum"
$docx pagebreak
$docx paragraph "The next paragraph: $loreipsum"
$docx pagebreak
$docx paragraph "The next paragraph: $loreipsum"

$docx header {
    $docx paragraph "Header with textbox"
    $docx textbox -name myTextBox -dimension {width 3cm height 3cm} \
        -positionH margin \
        -posOffsetH 3cm \
        -positionV margin \
        -posOffsetV 3cm \
        -wrapMode square \
        -wrapData {wrapText bothSides} {
            $docx paragraph "First paragraph in a textbox" -spacing {before 0}
            $docx paragraph "Second paragraph in a textbox" -align center
        }
} headerWithAnchoredTextbox
$docx pagebreak
$docx sectionstart \
    -defaultHeader $headerWithAnchoredTextbox \
    -evenHeader $headerWithAnchoredTextbox

$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx sectionend

$docx write docx-example2.docx
$docx destroy
