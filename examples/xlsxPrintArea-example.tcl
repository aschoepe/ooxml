#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
    source ../ooxml.tcl
}

set spreadsheet [::ooxml::xl_write new -creator {Creator Name}]
for {set s 1} {$s <= 4} {incr s} {
    if {[set sheet [$spreadsheet worksheet "Sheet$s"]] > -1} {
        for {set r 1} {$r < 6} {incr r} {
            $spreadsheet row $sheet
            for {set c 1} {$c < 6} {incr c} {
                $spreadsheet cell $sheet $s,$c,$r
            }
        }
        switch -- $s {
            1 {
                $spreadsheet printarea $sheet A1 C3
            }
            2 {
                #$spreadsheet printarea $sheet B2 C3
                $spreadsheet printarea $sheet 1,1 2,2
            }
            4 {
                $spreadsheet printarea $sheet A3 A4
            }
        }
    }
}
$spreadsheet write printArea.xlsx
$spreadsheet destroy