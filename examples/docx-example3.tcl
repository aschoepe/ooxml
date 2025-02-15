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

$docx paragraph "A very simple table:"
$docx simpletable $simpledata

$docx paragraph "Something more verbose please."
set table [list]
lappend table [list "Less words" "Also" $loreipsum]
lappend table [list "Less words" $loreipsum "Also"]
lappend table [list $loreipsum "Also" "Less words" ]
$docx simpletable $table

$docx paragraph "I want borders."
$docx paragraph "Sure thing:"
$docx simpletable $simpledata \
    -insideVborder {type single space 8} -insideHborder {type single space 8}

$docx paragraph "I meant outside."

$docx simpletable $simpledata \
    -topborder {type double space 8} \
    -bottomborder {type double space 8} \
    -startborder {type double space 8} \
    -endborder {type double space 8} \

$docx write docx-example3.docx
$docx destroy
