#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx
set docx [docx new]

# Set a default style
$docx style paragraphdefault -spacing {before 240 after 120}
# And a heading style
$docx style paragraph Heading1 -fontsize 28 -spacing {before 480 after 240}

# Set the margins
#$docx pagesetup 
$docx pagesetup -sizeAndOrientaion {width 15840 height 12240} \
    -margins {left 1cm right 1cm top 1cm bottom 1cm} \
    -topPageBorder {type dashed borderwidth 10}

$docx paragraph "Chapter one" -style Heading1
# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"

# Start a new section with different page setup
$docx sectionstart -sizeAndOrientaion {width 12240 height 15840} \
    -margins {left 3cm right 3cm top 4cm bottom 4cm} \
    -leftPageBorder {type thick borderwidth 20 color 00ff00 space 20}

$docx paragraph "Chapter two" -style Heading1
# Add some paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
for {set i 0} {$i < 8} {incr i} {
    $docx paragraph "The next paragraph: $loreipsum"
}

# End the section
$docx sectionend

# We are back to the pagesetup
$docx paragraph "Chapter three" -style Heading1
# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"

$docx write docx-example2.docx
$docx destroy
