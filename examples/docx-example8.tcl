#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl

package require fileutil::traverse

namespace eval ::docx::vars {
    variable foo "bar baz booooo"
    variable bar $::loreipsum
}

namespace import ::ooxml::docx::*

processdocx docx-example8-in.docx docx-example8-out.docx ::docx::vars
exec libreoffice docx-example8-out.docx
