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
$docx paragraph "Looks like a Heading" -pstyle Heading1

# Add two paragraphs
$docx paragraph "A very simple monoton paragraph: $loreipsum"
$docx paragraph "The next paragraph: $loreipsum"

# Next Heading
$docx paragraph "The next Heading" -pstyle Heading1

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
$docx append " or highlight color " -highlight green
$docx append " superscript " -verticalAlign superscript
$docx append " normal text "
$docx append " subscript " -verticalAlign subscript
$docx append " normal text "
$docx append " emboss text " -emboss true
$docx append " and you can freely combine them " -fontsize 15pt \
    -underline double -color ff0000
$docx append " which isn't so important because using them all would be ugly.\
               $loreipsum "
$docx mark mymark
$docx append "At the start of this sentence is a jump mark."

# Style heritage:
$docx style paragraph Base -font "Arial" -color 101010 -fontsize 12pt
$docx style paragraph H1 -basedon Base -fontsize 18pt
$docx style paragraph H2 -basedon Base -fontsize 16pt

$docx paragraph "Toplevel Heading" -pstyle H1
$docx paragraph "Second Level Heading 1" -pstyle H2
$docx paragraph $loreipsum -pstyle Base
$docx paragraph "Second Level Heading 2" -pstyle H2
$docx paragraph $loreipsum -pstyle Base
$docx simplecomment "This is ridiculous." -color 00ff00 \
      -date 2025-05-29T14:20:00 \
      -author "Rolf Ade"

$docx paragraph $loreipsum -align both \
    -indentation {firstLine 3cm hanging 2cm start 1cm end 4cm} \
    -spacing {before 240 after 120}

$docx paragraph $loreipsum -leftBorder {type thick borderwidth 20 color 00ff00 space 20}

$docx style paragraph withBorder -rightBorder {type thick borderwidth 20 color 00ff00 space 20}
$docx paragraph $loreipsum -pstyle withBorder
$docx paragraph $loreipsum -pstyle withBorder

$docx paragraph "Tabs test $loreipsum" -tabs {3cm 6cm 9cm 12cm}
$docx url "Link zu core.tcl-lang.org" https://core.tcl-lang.org -underline single
$docx comment -author "Document Creator" \
      -date [clock format [clock seconds] -format %Y-%m-%dT%H:%M:%SZ] \
      -initals "dc" {
    $docx paragraph "You have "
    $docx append "every" -bold 1
    $docx append " local formating at hand to format the comment text."
    $docx paragraph "Or several paragraphs. Though no tables. Even if the xsd allows them (if I read it right) neither libreoffice nor word grok this (and doesn't allow it via GUI)." -spacing {before 240}
}
$docx append " "
$docx jumpto "This is a link to a place inside the document." mymark -underline double
$docx append " "
$docx url "The url method also allows mailto links." mailto:pointsman@gmx.net -color 0000ff
$docx append " "
$docx url "The url method also allows mailto links with pre-set subject." "mailto:pointsman@gmx.net?subject=Feeback to your silly example document" -color 0000ff


$docx paragraph "This is a text frame. $loreipsum" -textframe {
    width 3500
    vSpace 300
    hSpace 300
    wrap auto
    vAnchor text
    hAnchor text
    xAlign inside
    yAlign inside
}

$docx style paragraph my_p_style -align both \
      -indentation {firstLine 3cm hanging 2cm start 1cm end 4cm} \
      -spacing {before 240 after 120}
$docx style character my_c_style -highlight green \
      -fontsize 15pt \
      -underline double \
      -color ff0000

$docx paragraph $loreipsum -pstyle my_p_style -cstyle my_c_style
$docx pagebreak
$docx textbox -dimension {width 3cm height 3cm} \
      -anchorData {locked 1} \
      -positionH page \
      -posOffsetH 6cm \
      -positionV page \
      -posOffsetV 6cm \
      -wrapMode none {
    $docx paragraph "First paragraph in a textbox" -spacing {before 0}
    $docx paragraph "Second paragraph in a textbox" -align center
}
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"
$docx paragraph "$loreipsum"

$docx write docx-example1.docx
$docx destroy
