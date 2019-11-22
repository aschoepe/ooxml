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
  set center [$spreadsheet style -horizontal center]
  set date [$spreadsheet style -numfmt [$spreadsheet numberformat -datetime]]
  set decimal [$spreadsheet style -numfmt [$spreadsheet numberformat -decimal -separator]]
  set text [$spreadsheet style -numfmt [$spreadsheet numberformat -string]]

  $spreadsheet column $sheet
  $spreadsheet column $sheet -width 30
  $spreadsheet column $sheet
  $spreadsheet column $sheet
  $spreadsheet column $sheet -style $decimal
  $spreadsheet column $sheet -style $date -width [::ooxml::CalcColumnWidth 16]
  $spreadsheet column $sheet

  $spreadsheet autofilter $sheet 0,0 0,6
  $spreadsheet freeze $sheet 1,2

  foreach {name value} [array get data] {
    lassign [split $name ,] row col
    switch -- $col {
      2 {
	$spreadsheet cell $sheet $value -index $name -string
      }
      default {
	$spreadsheet cell $sheet $value -index $name -globalstyle
      }
    }
  }
}
$spreadsheet write export1.xlsx
$spreadsheet destroy
