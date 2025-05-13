#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx

set docx [docx new]
$docx paragraph $loreipsum
$docx image book.jpg anchor -dimension {width 3cm height 3cm} -wrapMode square -wrapData {wrapText bothSides distL 8cm distR 8cm}

$docx paragraph $loreipsum
$docx image book.jpg anchor -dimension {width 8cm height 8cm} -wrapMode square -wrapData {wrapText left distR 8cm}

$docx paragraph $loreipsum
$docx image book.jpg anchor -dimension {width 8cm height 8cm} -wrapMode square -wrapData {wrapText right distR 8cm}

$docx paragraph $loreipsum
$docx image book.jpg anchor -dimension {width 8cm height 8cm} -wrapMode square -wrapData {wrapText largest distR 8cm}

$docx append $loreipsum

$docx pagebreak

$docx paragraph $loreipsum
$docx image book.jpg anchor -dimension {width 3cm height 3cm} \
      -positionH page \
      -posOffsetH 0cm \
      -positionV page \
      -posOffsetV 0cm \
      -wrapMode square \
      -wrapData {wrapText bothSides}
$docx append $loreipsum
$docx image book.jpg anchor -dimension {width 3cm height 3cm} \
      -positionH page \
      -posOffsetH 3cm \
      -positionV page \
      -posOffsetV 3cm \
      -wrapMode square \
      -wrapData {wrapText bothSides}
$docx append $loreipsum
$docx image book.jpg anchor -dimension {width 3cm height 3cm} \
      -positionH page \
      -posOffsetH 6cm \
      -positionV page \
      -posOffsetV 6cm \
      -wrapMode square \
      -wrapData {wrapText bothSides}
$docx append $loreipsum



# $docx paragraph $loreipsum
# $docx image book.jpg inline -dimension {width 3cm height 3cm} -bwMode black
# $docx append $loreipsum

# $docx paragraph "A new paragraph"
# $docx image book.jpg -dimension {width 8cm height 8cm} -bwMode black

$docx write docx-example6.docx
$docx destroy
