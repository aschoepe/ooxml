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
  # single formula autoincrement column index

  $spreadsheet row $sheet
  $spreadsheet cell $sheet 3
  $spreadsheet cell $sheet 5
  $spreadsheet cell $sheet {} -formula A1+B1

  # vertical shared formula C3 to C5

  $spreadsheet cell $sheet 1 -index 2,0
  $spreadsheet cell $sheet 2 -index 2,1
  $spreadsheet cell $sheet {} -index 2,2 -formula {A3+B3} -formularef C3:C5 -formulaidx 0

  $spreadsheet cell $sheet 2 -index 3,0
  $spreadsheet cell $sheet 3 -index 3,1
  $spreadsheet cell $sheet {} -index 3,2 -formulaidx 0

  $spreadsheet cell $sheet 3 -index A5
  $spreadsheet cell $sheet 4 -index B5
  $spreadsheet cell $sheet {} -index C5 -formulaidx 0

  # horizontal shared formula A9 to C9

  $spreadsheet cell $sheet 1 -index 6,0
  $spreadsheet cell $sheet 2 -index 7,0
  $spreadsheet cell $sheet {} -index A9 -formula {A7+A8} -formularef 8,0:8,2 -formulaidx 1

  $spreadsheet cell $sheet 2 -index 6,1
  $spreadsheet cell $sheet 3 -index 7,1
  $spreadsheet cell $sheet {} -index 8,1 -formulaidx 1

  $spreadsheet cell $sheet 3 -index C7
  $spreadsheet cell $sheet 4 -index C8
  $spreadsheet cell $sheet {} -index C9 -formulaidx 1
}
$spreadsheet write export4.xlsx
$spreadsheet destroy
