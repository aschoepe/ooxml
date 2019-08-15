#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

set spreadsheet [::ooxml::xl_write new -creator {User A} -created {2019-08-10 10:01:30} -modifiedby {User B} -modified {2019-08-10 12:30:01} -application {Tcl Example Script 9}]
set wrap [$spreadsheet style -wrap]

if {[set sheet [$spreadsheet worksheet {Sheet 1}]] > -1} {
  $spreadsheet cell $sheet {my text the will be automatically wrapped by excel} -index A1 -style $wrap
  $spreadsheet write export9.xlsx
}
$spreadsheet destroy
