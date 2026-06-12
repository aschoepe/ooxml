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
    
$docx paragraph "Next paragraph $loreipsum"

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

$docx paragraph "Append to document $loreipsum"

$docx appendTo $rowToAppend {
    $docx tablecell {
        $docx paragraph "Later added cell to row."
    }
}

$docx paragraph "Append to document $loreipsum"

$docx appendTo $cellToAppend {
    $docx append " Later added text"
    $docx paragraph "Later added paragraph."
}

$docx write docx-example12.docx
$docx destroy
