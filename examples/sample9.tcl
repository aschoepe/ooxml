#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

source array.tcl

set spreadsheet [::ooxml::xl_write new -creator {Steve Landers}]
if {[set sheet [$spreadsheet worksheet {Table 1}]] > -1} {
  $spreadsheet row $sheet
  $spreadsheet cell $sheet {} -hyperlink {https://fossil.sowaswie.de/ooxml/timeline}
  $spreadsheet row $sheet
  $spreadsheet cell $sheet {ooXML Home} -hyperlink {https://fossil.sowaswie.de/ooxml/}
  $spreadsheet row $sheet
  $spreadsheet cell $sheet {check tooltip} -hyperlink {https://fossil.sowaswie.de/ooxml/wiki?name=man-page} -tooltip {ooXML Manual Page}
}
$spreadsheet write export9.xlsx
$spreadsheet destroy
