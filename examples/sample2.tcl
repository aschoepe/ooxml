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
  set decimal [$spreadsheet style -numfmt [$spreadsheet numberformat -decimal -separator -red]]
  set text [$spreadsheet style -numfmt [$spreadsheet numberformat -string]]

  $spreadsheet column $sheet -width 30 -index 1
  $spreadsheet column $sheet -style $decimal -index 4
  $spreadsheet column $sheet -style $date -width [::ooxml::CalcColumnWidth 16]

  $spreadsheet autofilter $sheet 0,0 0,6
  $spreadsheet freeze $sheet C2

  set lastRow -1
  foreach name [lsort -dictionary [array names data]] {
    lassign [split $name ,] row col
    if {$row != $lastRow} {
      set lastRow $row
      $spreadsheet row $sheet
    }
    switch -- $col {
      2 {
	$spreadsheet cell $sheet $data($name) -index $col -string
      }
      default {
	$spreadsheet cell $sheet $data($name) -index $col -globalstyle
      }
    }
  }
}
$spreadsheet write export2.xlsx
$spreadsheet destroy
