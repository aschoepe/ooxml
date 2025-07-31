#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

if {$argc < 1} {
    error "usage: $argv0 method ?prelist? ?postlist?"
}

namespace eval ooxml::docx {
    variable docgen 1
}
source ../ooxml.tcl
source ../ooxml-docx.tcl
package require ooxml::docx

lassign $argv method prelist postlist
set docx [::ooxml::docx::docx new]
if {![catch {eval $docx $method {*}$prelist -foo foo {*}$postlist} errMsg1]} {
    error "execution did not fail; no error message to parse"
}
if {![regexp {^.*===(.*)===(.*)$} $errMsg1 -> typename allowedOptions]} {
    error "could not parse the error message \"$errMsg1\""
}

foreach allowedOption [split $allowedOptions |] {
    if {![catch {eval $docx $method {*}$prelist $allowedOption foo {*}$postlist} errMsg2]} {
        puts "$allowedOption string"
        continue
    }
    if {![regexp {^.*===(.*)===(.*)$} $errMsg2 -> typename allowedValues]} {
        if {[regexp {^.*expected a (.*)$} $errMsg2 -> expectedValue]} {
            puts "$allowedOption $expectedValue"
            continue
        } elseif {[regexp {^.*match the regular expression (.*)$} $errMsg2 -> expectedValue]} {
            puts "$allowedOption $expectedValue"
            continue
        }                
        puts "$allowedOption could not parse the error message \"$errMsg2\""
        continue
    }
    puts $allowedOption
    if {[regexp {expected is a key value pair} $errMsg2]} {
        foreach allowedValue [split $allowedValues |] {
            if {![catch {eval $docx $method {*}$prelist $allowedOption [list [list $allowedValue foo]] {*}$postlist} errMsg3]} {
                puts "    $allowedValue string"
                continue
            }
            if {![regexp {^.*===(.*)===(.*)$} $errMsg3 -> keyType keyValues]} {
                puts "$allowedValue could not parse the error message \"$errMsg3\""
                continue
            }
            puts "    $allowedValue ($keyType)"
            foreach keyValue [split $keyValues |] {
                puts "        $keyValue"
            }
        }
    } elseif {[regexp {expected one of} $errMsg2]} {
        foreach allowedValue [split $allowedValues |] {
            puts "    $allowedValue"
        }
    }
}

