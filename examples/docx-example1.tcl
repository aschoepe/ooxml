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

# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"

# What about headings and such? Well, WordprocessingML doesn't really
# has a basic concept for that. They are (for now) just another
# paragraph with special settings. Let's define a style for heading1:
$docx style paragraph Heading1 -fontsize 28 -spacing {before 480 after 240}
$docx paragraph "Looks like a Heading" -style Heading1

# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"

# Next Heading
$docx paragraph "The next Heading" -style Heading1

# Certain parts of a paragraph may have local formating. This is done
# by starting a paragraph and then append to that.
$docx paragraph "The start of a paragraph. "
$docx append "Another chunk of the paragraph, just with the same formating. "
$docx append "This chunk of the paragraph is obviously locally formated." \
      -color 0000ff
$docx append "You have certain effects already on your hand "
foreach textOnOff {
        bold
        dstrike
        italic 
        strike
} {
    $docx append "$textOnOff " -$textOnOff on
}
$docx append "color " -color 00ff00
foreach text {Serveral kinds of underline} underline {
    single
    words
    double
    thick
    dotted
} {
    $docx append "$text $text " -underline $underline
}

$docx append " or fontsize " -fontsize 18pt
$docx append " and you can freely combine them " -fontsize 15pt \
    -underline double -color ff0000
$docx append " which isn't so important because using them all would be ugly.\
               Rest of the paragraph $loreipsum"


# Style heritage:
$docx style paragraph Base -font "Latin Modern Mono Caps" -color 101010 -fontsize 12pt
$docx style paragraph H1 -basedon Base -fontsize 18pt
$docx style paragraph H2 -basedon Base -fontsize 16pt

$docx paragraph "Toplevel Heading" -style H1
$docx paragraph "Second Level Heading 1" -style H2
$docx paragraph $loreipsum -style Base
$docx paragraph "Second Level Heading 2" -style H2
$docx paragraph $loreipsum -style Base


$docx paragraph $loreipsum -align both \
    -indentation {firstLine 3cm hanging 2cm start 1cm end 4cm} \
    -spacing {before 240 after 120}

$docx paragraph $loreipsum -leftBorder {type thick borderwidth 20 color 00ff00 space 20}

$docx style paragraph withBorder -rightBorder {type thick borderwidth 20 color 00ff00 space 20}
$docx paragraph $loreipsum -style withBorder
$docx paragraph $loreipsum -style withBorder

$docx paragraph "Tabs test $loreipsum" -tabs {3cm 6cm 9cm 12cm}

$docx write docx-example1.docx
$docx destroy
