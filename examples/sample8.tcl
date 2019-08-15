#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

array set workbook [ooxml::xl_read form8.xlsx]

set data(NAME) {Erika Mustermann}
set data(ANSCHRIFT) {Heidestrasse 17}
set data(PLZORT) {51147 Köln}
set data(positionen) 3
set data(0,BEZEICHUNG) {Kopierpapier 80g 500 Blatt}
set data(0,MENGE) 5
set data(1,BEZEICHUNG) {Ordner A4 breit}
set data(1,MENGE) 1
set data(2,BEZEICHUNG) {Haftnotizen 5x5cm 40 Blatt}
set data(2,MENGE) 2

foreach name {NAME ANSCHRIFT PLZORT POSITION BEZEICHUNG MENGE} {
  set position($name) {}
}
foreach {name value} [array get workbook 0,v,*] {
  if {[set value [string trim $value]] in {NAME ANSCHRIFT PLZORT POSITION BEZEICHUNG MENGE}} {
    lassign [split $name ,] sheet tag row column
    set position($value) [list sheet $sheet row $row column $column]
  }
}

if 0 {
  set workbook([dict get $position(NAME) sheet],v,[dict get $position(NAME) row],[dict get $position(NAME) column]) $data(NAME)
  set workbook([dict get $position(ANSCHRIFT) sheet],v,[dict get $position(ANSCHRIFT) row],[dict get $position(ANSCHRIFT) column]) $data(ANSCHRIFT)
  set workbook([dict get $position(PLZORT) sheet],v,[dict get $position(PLZORT) row],[dict get $position(PLZORT) column]) $data(PLZORT)

  for {set i 0} {$i < $data(positionen)} {incr i} {
    set workbook([dict get $position(POSITION) sheet],v,[expr {[dict get $position(POSITION) row] + $i}],[dict get $position(POSITION) column]) [expr {$i + 1}]
    set workbook([dict get $position(BEZEICHUNG) sheet],v,[expr {[dict get $position(BEZEICHUNG) row] + $i}],[dict get $position(BEZEICHUNG) column]) $data($i,BEZEICHUNG)
    set workbook([dict get $position(MENGE) sheet],v,[expr {[dict get $position(MENGE) row] + $i}],[dict get $position(MENGE) column]) $data($i,MENGE)
  }
}

proc Cell { *workbook *position item } {
  upvar ${*workbook} workbook
  upvar ${*position} position

  return [dict get $workbook(sheetmap) [dict get $position($item) sheet]]
}

proc Options { *workbook *position item {rowoffset 0} } {
  upvar ${*workbook} workbook
  upvar ${*position} position

  set sheet [dict get $position($item) sheet]
  set row [expr {[dict get $position($item) row] + $rowoffset}]
  set column [dict get $position($item) column]

  set options [list -index $row,$column]
  if {[info exists workbook($sheet,s,$row,$column)]} {
    lappend options -style $workbook($sheet,s,$row,$column)
  }

  return $options
}

set spreadsheet [::ooxml::xl_write new]
$spreadsheet presetstyles workbook
$spreadsheet presetsheets workbook
if 1 {
  $spreadsheet cell [Cell workbook position NAME] $data(NAME) {*}[Options workbook position NAME]
  $spreadsheet cell [Cell workbook position ANSCHRIFT] $data(ANSCHRIFT) {*}[Options workbook position ANSCHRIFT]
  $spreadsheet cell [Cell workbook position PLZORT] $data(PLZORT) {*}[Options workbook position PLZORT]
  for {set i 0} {$i < $data(positionen)} {incr i} {
    $spreadsheet cell [Cell workbook position POSITION] [expr {$i + 1}] {*}[Options workbook position POSITION $i]
    $spreadsheet cell [Cell workbook position BEZEICHUNG] $data($i,BEZEICHUNG) {*}[Options workbook position BEZEICHUNG $i]
    $spreadsheet cell [Cell workbook position MENGE] $data($i,MENGE) {*}[Options workbook position MENGE $i]
  }
}
$spreadsheet write export8.xlsx
$spreadsheet destroy

