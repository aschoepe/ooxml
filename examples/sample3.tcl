#!/bin/sh
#\
exec wish8.6 "$0" "$@"

package require Tk
package require tablelist

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

source array.tcl

# build Tablelist from array

set lb .lb
tablelist::tablelist $lb -height 16 -width 100 -showseparators 1 -titlecolumns 3 -labelcommand tablelist::sortByColumn
grid $lb -row 0 -column 0 -sticky nwse
ttk::scrollbar .x -orient horizontal -command [list $lb xview]
$lb configure -xscrollcommand [list .x set]
grid .x -row 1 -column 0 -sticky we
ttk::scrollbar .y -orient vertical -command [list $lb yview]
$lb configure -yscrollcommand [list .y set]
grid .y -row 0 -column 1 -sticky ns
grid columnconfigure . 0 -weight 1
grid rowconfigure . 0 -weight 1
ttk::button .export -text Export -command [list ::ooxml::tablelist_to_xl $lb -file export3.xlsx -globalstyle]
grid .export -row 2 -column 0

set list {}
set cols -1
foreach name [lsort -dictionary [array names data]] {
  lassign [split $name ,] row col
  if {$row == 0} {
    switch -- $data($name) {
      {order number} {
        set justify center
      }
      item - price {
        set justify right
      }
      default {
        set justify left
      }
    }
    $lb insertcolumns end 0 $data($name) $justify
    incr cols
  } elseif {$row > 0} {
    lappend rows $row
  }
}

foreach row [lsort -integer -unique $rows] {
  set list {}
  for {set col 0} {$col <= $cols} {incr col} {
    if {[info exists data($row,$col)]} {
      lappend list $data($row,$col)
    } else {
      lappend list {}
    }
  }
  $lb insert end $list
}


# this sample callback is the default source code if nothing else is set

proc TablelistToXlCallback { spreadsheet sheet maxcol column title width align sortmode hide } {
  set left 0
  set center [$spreadsheet style -horizontal center]
  set right [$spreadsheet style -horizontal right]
  set date [$spreadsheet style -numfmt [$spreadsheet numberformat -datetime]]
  set decimal [$spreadsheet style -numfmt [$spreadsheet numberformat -decimal -red]]
  set text [$spreadsheet style -numfmt [$spreadsheet numberformat -string]]

  if {$column == -1} {
    $spreadsheet defaultdatestyle $date
  } else {
    switch -- $align {
      center {
        $spreadsheet column $sheet -index $column -style $center
      }
      right {
        $spreadsheet column $sheet -index $column -style $right
      }
      default {
        $spreadsheet column $sheet -index $column -style $left
      }
    }
  }
}

