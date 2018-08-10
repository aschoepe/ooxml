#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

#package require ooxml
source ooxml.tcl

array set workbook [ooxml::xl_read original_excel.xlsx]

set spreadsheet [::ooxml::xl_write new]
$spreadsheet presetstyles workbook
$spreadsheet presetsheets workbook
$spreadsheet write export7.xlsx
$spreadsheet destroy

