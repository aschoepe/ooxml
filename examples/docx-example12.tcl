#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx

set docx [docx new]

$docx paragraph $loreipsum

set tableToAppend [$docx table {
    set rowToAppend [$docx tablerow {
        $docx tablecell -vspan continue {
            $docx paragraph "3.1"
        }
        set cellToAppend [$docx tablecell {
            $docx paragraph "3.2"
        }]
        $docx tablecell {
            $docx paragraph "3.3"
        }
    }]
}]

$docx paragraph "Next paragraph "
set comment1 [$docx comment {
    $docx paragraph "Comment started."
}]
$docx append $loreipsum

$docx appendTo $tableToAppend {
    $docx tablerow {
        $docx tablecell {
            $docx paragraph "Later"
        }
        $docx tablecell {
            $docx paragraph " added "
        }
        $docx tablecell {
            $docx paragraph "row"
        }
    }
}

$docx paragraph "Append to "
$docx commentrangestart commentrange
$docx append "document"
set comment2 [$docx commentrangeend $commentrange {
    $docx paragraph "Another comment started."
}]
$docx append " $loreipsum"

$docx appendTo $rowToAppend {
    $docx tablecell {
        $docx paragraph "Later added cell to row."
    }
}

$docx paragraph "A second table:"

$docx table {
    $docx tablerow {
        set cell2 [$docx tablecell {
            $docx paragraph "We append later to this cell."
        }]
        $docx tablecell {
            $docx paragraph "foo"
        }
    }
    $docx tablerow {
        $docx tablecell {
            $docx paragraph "bar" 
        }
        $docx tablecell {
            $docx paragraph "and grill"
        }
    }
}
$docx paragraph "Append"
set footnote1 [$docx footnote {
    $docx paragraph "Footnote one."
}]
$docx append " to"
set footnote2 [$docx footnote {
    $docx paragraph "Footnote two."
}]
$docx append " document"
set endnote1 [$docx endnote {
    $docx paragraph "Endnote one."
}]
$docx append " $loreipsum"
               
$docx appendTo $endnote1 {
    $docx paragraph "Added to endnote."
}
$docx appendTo $footnote2 {
    $docx paragraph "Added to footnote two."
}
$docx appendTo $footnote1 {
    $docx paragraph "Added to footnote one."
}
              
$docx appendTo $comment2 {
    $docx paragraph "Later added"
}

$docx appendTo $comment1 {
    $docx paragraph "Later added"
}

$docx appendTo $cell2 {
    $docx paragraph "This was added later" -italic 1 
}
$docx appendTo $cellToAppend {
    $docx append " Later added text"
    $docx paragraph "Later added paragraph."
}

$docx appendTo $cellToAppend {
    $docx append " Later added text"
    $docx paragraph "Later added paragraph."
}

$docx header {
    $docx paragraph "This header"
} defaultHeader defaultHeaderObjId
$docx footer {
    $docx field page
} defaultFooter defaultFooterObjId
$docx pagesetup -defaultHeader $defaultHeader \
    -defaultFooter $defaultFooter

$docx paragraph "Append to document $loreipsum"

$docx appendTo $defaultHeaderObjId {
    $docx append " later on added to header"
}
$docx appendTo $defaultFooterObjId {
    $docx append " later added to footer" -bold 1
}

$docx pagebreak

set textbox1 [$docx textbox -dimension {width 3cm height 3cm} \
      -anchorData {distL 1cm distR 1cm distT 1cm} \
      -positionH page \
      -posOffsetH 6cm \
      -positionV page \
      -posOffsetV 6cm \
      -wrapMode square \
      -wrapData {wrapText bothSides} {
          $docx paragraph "Second paragraph in a textbox" -align center
      }]

$docx paragraph $loreipsum

$docx appendTo $textbox1 {
    $docx paragraph "Later added paragraph to the textbox" -spacing {before 0}
}

$docx write docx-example12.docx
$docx destroy
