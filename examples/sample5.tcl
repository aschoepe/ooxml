#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

set spreadsheet [::ooxml::xl_write new -creator {Alexander Schöpe}]
if {[set sheet [$spreadsheet worksheet {Tabelle 1}]] > -1} {
  set date [$spreadsheet style -numfmt [$spreadsheet numberformat -datetime]]
  $spreadsheet defaultdatestyle $date
  # 2018-03-02 17:39 -> 43161.73542
  $spreadsheet cell $sheet "2018-03-02 17:39" -index 0,0
}
$spreadsheet write export5.xlsx
$spreadsheet destroy
