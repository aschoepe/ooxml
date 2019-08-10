#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

source array.tcl

set spreadsheet [::ooxml::xl_write new -creator {Alexander Schöpe}]
if {[set sheet [$spreadsheet worksheet {Tabelle 1}]] > -1} {
  $spreadsheet row $sheet
  $spreadsheet cell $sheet 3
  $spreadsheet cell $sheet 5
  $spreadsheet cell $sheet {} -formula A1+B1

  $spreadsheet write export4.xlsx
}
$spreadsheet destroy
