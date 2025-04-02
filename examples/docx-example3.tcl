#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx
set docx [docx new]

set simpledata {{a b c} {1 2 3} {I II III}}
# Set a default style
$docx style paragraphdefault -spacing {before 240 after 120}

$docx paragraph "Covered simple tables."
$docx paragraph "A very simple table:"
$docx simpletable $simpledata -columnwidths {30pt 40pt 50pt}

$docx paragraph "Something more verbose please."
set table [list]
lappend table [list "Less words" "Also" $loreipsum]
lappend table [list "Less words" $loreipsum "Also"]
lappend table [list $loreipsum "Also" "Less words" ]
$docx simpletable $table

$docx paragraph "I want borders."
$docx paragraph "Sure thing:"
$docx simpletable $simpledata \
    -insideVBorder {type single space 8} -insideHBorder {type single space 8}

$docx paragraph "I meant outside."

$docx simpletable $simpledata \
    -topBorder {type double} \
    -bottomBorder {type double} \
    -startBorder {type double} \
    -endBorder {type double} 

$docx style table myTable \
    -topBorder {type single borderwidth 320} \
    -bottomBorder {type single borderwidth 320} \
    -startBorder {type single borderwidth 320} \
    -endBorder {type single borderwidth 320} 
$docx paragraph "A table with -style"
$docx simpletable $simpledata -style myTable -width 100pt


$docx style paragraph tableFirstRow -align center -bold on
$docx style paragraph tableLastRow -align end
$docx paragraph "A simple table with other style in first and last row."
$docx simpletable $simpledata -firstStyle tableFirstRow \
    -lastStyle tableLastRow

$docx paragraph "A first start of a complex table"
$docx table {
    $docx tablerow {
        $docx tablecell {
            $docx paragraph "A normal paragraph with all paragraph styling" -fontsize 18pt
        }
        $docx tablecell {
            $docx paragraph "Another cell" -underline single
        }
    }
    $docx tablerow {
        $docx tablecell {
            $docx simpletable $simpledata -columnwidths {30pt 40pt 50pt}
        }
        $docx tablecell {
            $docx paragraph "Fill grid"
        }
    }
}

$docx write docx-example3.docx
$docx destroy
