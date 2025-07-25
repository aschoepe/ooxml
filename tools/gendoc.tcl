#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

if {$argc != 1} {
    error "usage: $argv0 <argumentstring>"
}

namespace eval ooxml::docx {
    variable docgen 1
}
source ../ooxml.tcl
source ../ooxml-docx.tcl
package require ooxml::docx

set docx [::ooxml::docx::docx new]
#$docx paragraph "foo" -underline foo
if {![catch {eval $docx [lindex $argv 0]} errMsg]} {
    error "execution did not fail; no error message to parse"
}
if {![regexp {^.*===(.*)===(.*)$} $errMsg -> typename allowedValues]} {
    error "could not parse the error message \"$errMsg\""
}

puts "$typename"
foreach allowedValue [split $allowedValues |] {
    puts $allowedValue
}

