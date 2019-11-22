#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

set auto_path [linsert $auto_path 0 ..]
if {[catch {package require ooxml}]} {
  source ../ooxml.tcl
}

set spreadsheet [::ooxml::xl_write new -creator {Alexander Schöpe}]
if {[set sheet [$spreadsheet worksheet {Blatt 1}]] > -1} {
  set bold [$spreadsheet style -font [$spreadsheet font -bold]]
  set italic [$spreadsheet style -font [$spreadsheet font -italic]]
  set underline [$spreadsheet style -font [$spreadsheet font -underline]]

  set red [$spreadsheet style -font [$spreadsheet font -color Red]]

  set font9 [$spreadsheet style -font [$spreadsheet font -size 9]]
  set font18 [$spreadsheet style -font [$spreadsheet font -size 18]]

  set rotate90 [$spreadsheet style -rotate 90 -horizontal center]
  set rotate45 [$spreadsheet style -rotate 45]

  set left [$spreadsheet style -horizontal left]
  set center [$spreadsheet style -horizontal center]
  set right [$spreadsheet style -horizontal right]
  set top [$spreadsheet style -vertical top]
  set vcenter [$spreadsheet style -vertical center]
  set bottom [$spreadsheet style -vertical bottom]
  set hvcenter [$spreadsheet style -horizontal center -vertical center]

  set yellow [$spreadsheet style -fill [$spreadsheet fill -fgcolor FFFFFF00 -bgcolor 64 -patterntype solid]]

  set bLeft [$spreadsheet style -border [$spreadsheet border -leftstyle thin]]
  set bBottom [$spreadsheet style -border [$spreadsheet border -bottomstyle thin]]
  set bBottom2 [$spreadsheet style -border [$spreadsheet border -bottomstyle double]]
  set bBottomB [$spreadsheet style -border [$spreadsheet border -bottomstyle medium]]
  set bBottomD [$spreadsheet style -border [$spreadsheet border -bottomstyle dashed]]
  set bDiagonal [$spreadsheet style -border [$spreadsheet border -diagonalstyle medium -diagonaldirection up]]

  set dec2 [$spreadsheet style -numfmt [$spreadsheet numberformat -decimal]]
  set dec2t [$spreadsheet style -numfmt [$spreadsheet numberformat -decimal -separator]]
  set dec3 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {#,##0.000}]]
  set dec3t [$spreadsheet style -numfmt [$spreadsheet numberformat -format {#,##0.000_ ;[Red]\-#,##0.000\ }]]
  set currency [$spreadsheet style -numfmt [$spreadsheet numberformat -format {#,##0.00\ "€"}]]
  set iso8601 [$spreadsheet style -numfmt [$spreadsheet numberformat -iso8601]]
  set date [$spreadsheet style -numfmt [$spreadsheet numberformat -date]]
  set date2 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {[$-F800]dddd\,\ mmmm\ dd\,\ yyyy}]]
  set date3 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {[$-407]d/\ mmmm\ yyyy;@}]]
  set date4 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {d/m/yy\ h:mm;@}]]
  set time [$spreadsheet style -numfmt [$spreadsheet numberformat -time]]
  set time2 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {[$-F400]h:mm:ss\ AM/PM}]]
  set percent [$spreadsheet style -numfmt [$spreadsheet numberformat -percent]]
  set percent2 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {0.00%}]]
  set scientific [$spreadsheet style -numfmt [$spreadsheet numberformat -scientific]]
  set fraction [$spreadsheet style -numfmt [$spreadsheet numberformat -fraction]]
  set fraction2 [$spreadsheet style -numfmt [$spreadsheet numberformat -format {# ??/??}]]
  set text [$spreadsheet style -numfmt [$spreadsheet numberformat -string]]
  set wrap [$spreadsheet style -wrap]

  $spreadsheet column $sheet -index 0 -width 17.33203125 -bestfit
  $spreadsheet column $sheet -index 1 -width 20.5 -bestfit
  $spreadsheet column $sheet -index 4 -width 31.1640625 -bestfit
  $spreadsheet column $sheet -index 7 -width 11.1640625 -bestfit ;# -style 19
  $spreadsheet column $sheet -index 8 ;# -style 15

  $spreadsheet cell $sheet Standard -index 0,0 -string
  $spreadsheet cell $sheet 3.1415 -index 0,1
  $spreadsheet cell $sheet 1 -index 0,8 ;# -style 15

  $spreadsheet cell $sheet Standard -index 1,0 -string
  $spreadsheet cell $sheet Text -index 1,1 -string

  $spreadsheet cell $sheet {Zahl 2} -index 2,0 -string
  $spreadsheet cell $sheet 3.1415 -index 2,1 -style $dec2
  $spreadsheet cell $sheet 1 -index 2,2
  $spreadsheet cell $sheet 2 -index 2,3
  $spreadsheet cell $sheet 0.00 -index 2,4 -string ;# -style 15

  $spreadsheet cell $sheet {Zahl 2 T} -index 3,0 -string
  $spreadsheet cell $sheet 3.1415 -index 3,1 -style $dec2t
  $spreadsheet cell $sheet 2 -index 3,2
  $spreadsheet cell $sheet 4 -index 3,3
  $spreadsheet cell $sheet #,##0.00 -index 3,4 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 3,7 -style $date4

  $spreadsheet cell $sheet {Zahl 3} -index 4,0 -string
  $spreadsheet cell $sheet 3.1415 -index 4,1 -style $dec3
  $spreadsheet cell $sheet 17 -index 4,2
  $spreadsheet cell $sheet 174 -index 4,3
  $spreadsheet cell $sheet #,##0.000 -index 4,4 -string

  $spreadsheet cell $sheet {Zahl 3 C} -index 5,0 -string
  $spreadsheet cell $sheet -3.1415 -index 5,1 -style $dec3t
  $spreadsheet cell $sheet 16 -index 5,2
  $spreadsheet cell $sheet 173 -index 5,3
  $spreadsheet cell $sheet {#,##0.000_ ;[Red]\-#,##0.000\ } -index 5,4 -string

  $spreadsheet cell $sheet Währung -index 6,0 -string
  $spreadsheet cell $sheet 3.1415 -index 6,1 -style $currency
  $spreadsheet cell $sheet 3 -index 6,2
  $spreadsheet cell $sheet 166 -index 6,3
  $spreadsheet cell $sheet {#,##0.00\ "€"} -index 6,4 -string

  $spreadsheet cell $sheet tt.mm.jj -index 7,0 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 7,1 -style $date
  $spreadsheet cell $sheet 4 -index 7,2
  $spreadsheet cell $sheet 14 -index 7,3
  $spreadsheet cell $sheet mm-dd-yy -index 7,4 -string

  $spreadsheet cell $sheet {tttt, t.mmmm.jj} -index 8,0 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 8,1 -style $date2
  $spreadsheet cell $sheet 5 -index 8,2
  $spreadsheet cell $sheet 167 -index 8,3
  $spreadsheet cell $sheet {[$-F800]dddd\,\ mmmm\ dd\,\ yyyy} -index 8,4 -string

  $spreadsheet cell $sheet t.mmmm.jjjj -index 9,0 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 9,1 -style $date3
  $spreadsheet cell $sheet 6 -index 9,2
  $spreadsheet cell $sheet 168 -index 9,3
  $spreadsheet cell $sheet {[$-407]d/\ mmmm\ yyyy;@} -index 9,4 -string

  $spreadsheet cell $sheet hh:mm:ss -index 10,0 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 10,1 -style $time2
  $spreadsheet cell $sheet 7 -index 10,2
  $spreadsheet cell $sheet 169 -index 10,3
  $spreadsheet cell $sheet {[$-F400]h:mm:ss\ AM/PM} -index 10,4 -string

  $spreadsheet cell $sheet hh:mm -index 11,0 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 11,1 -style $time
  $spreadsheet cell $sheet 8 -index 11,2
  $spreadsheet cell $sheet 170 -index 11,3
  $spreadsheet cell $sheet {h:mm;@} -index 11,4 -string

  $spreadsheet cell $sheet {jjjj-mm-tt hh:mm:ss} -index 12,0 -string
  $spreadsheet cell $sheet {2018-02-09 16:17:00} -index 12,1 -style $iso8601
  $spreadsheet cell $sheet 9 -index 12,2
  $spreadsheet cell $sheet 171 -index 12,3
  $spreadsheet cell $sheet {yyyy\-mm\-dd\ hh:mm:ss;@} -index 12,4 -string

  $spreadsheet cell $sheet {Prozent 2} -index 13,0 -string
  $spreadsheet cell $sheet 3.1415 -index 13,1 -style $percent2
  $spreadsheet cell $sheet 10 -index 13,2
  $spreadsheet cell $sheet 10 -index 13,3
  $spreadsheet cell $sheet 0.00% -index 13,4 -string

  $spreadsheet cell $sheet {Bruch 1} -index 14,0 -string
  $spreadsheet cell $sheet 3.1415 -index 14,1 -style $fraction
  $spreadsheet cell $sheet 12 -index 14,2
  $spreadsheet cell $sheet 12 -index 14,3
  $spreadsheet cell $sheet {# ?/?} -index 14,4 -string

  $spreadsheet cell $sheet {Bruch 2} -index 15,0 -string
  $spreadsheet cell $sheet 3.1415 -index 15,1 -style $fraction2
  $spreadsheet cell $sheet 13 -index 15,2
  $spreadsheet cell $sheet 13 -index 15,3
  $spreadsheet cell $sheet {# ??/??} -index 15,4 -string

  $spreadsheet cell $sheet {Wissenschaftl. 2} -index 16,0 -string
  $spreadsheet cell $sheet 3.1415 -index 16,1 -style $scientific
  $spreadsheet cell $sheet 14 -index 16,2
  $spreadsheet cell $sheet 11 -index 16,3
  $spreadsheet cell $sheet 0.00E+00 -index 16,4 -string

  $spreadsheet cell $sheet Text -index 17,0 -string
  $spreadsheet cell $sheet 3,1415 -index 17,1 -string -style $text
  $spreadsheet cell $sheet 15 -index 17,2
  $spreadsheet cell $sheet 49 -index 17,3
  $spreadsheet cell $sheet @ -index 17,4 -string

  $spreadsheet cell $sheet {Filter A} -index 19,5 -string ;# -style 19
  $spreadsheet cell $sheet {Filter B} -index 19,6 -string ;# -style 15

  $spreadsheet cell $sheet unten -index 20,0 -string
  $spreadsheet cell $sheet {} -index 20,1 -style $bBottom

  $spreadsheet cell $sheet links -index 21,0 -string
  $spreadsheet cell $sheet {} -index 21,1 -style $bLeft
  $spreadsheet cell $sheet fett -index 21,4 -string -style $bold

  $spreadsheet cell $sheet {unten doppelt} -index 22,0 -string
  $spreadsheet cell $sheet {} -index 22,1 -style $bBottom2
  $spreadsheet cell $sheet kursiv -index 22,4 -string -style $italic

  $spreadsheet cell $sheet {unten mittel} -index 23,0 -string
  $spreadsheet cell $sheet {} -index 23,1 -style $bBottomB
  $spreadsheet cell $sheet unterstrichen -index 23,4 -string -style $underline

  $spreadsheet cell $sheet {diagonal mittel} -index 24,0 -string
  $spreadsheet cell $sheet {} -index 24,1 -style $bDiagonal

  $spreadsheet cell $sheet {unten gestrichelt} -index 25,0 -string
  $spreadsheet cell $sheet {} -index 25,1 -style $bBottomD

  $spreadsheet cell $sheet {vorne rot} -index 27,0 -string
  $spreadsheet cell $sheet rot -index 27,1 -string -style $red

  $spreadsheet cell $sheet {hinten gelb} -index 28,0 -string
  $spreadsheet cell $sheet gelb -index 28,1 -string -style $yellow

  $spreadsheet cell $sheet links -index 30,0 -string
  $spreadsheet cell $sheet links -index 30,1 -string -style $left

  $spreadsheet cell $sheet mitte -index 31,0 -string
  $spreadsheet cell $sheet mitte -index 31,1 -string -style $center

  $spreadsheet cell $sheet rechts -index 32,0 -string
  $spreadsheet cell $sheet rechts -index 32,1 -string -style $right

  $spreadsheet cell $sheet oben -index 33,0 -string
  $spreadsheet cell $sheet oben -index 33,1 -string -style $top

  $spreadsheet cell $sheet mitte -index 34,0 -string
  $spreadsheet cell $sheet mitte -index 34,1 -string -style $vcenter
  $spreadsheet cell $sheet {Calibri 9} -index 34,4 -string -style $font9

  $spreadsheet cell $sheet unten -index 35,0 -string
  $spreadsheet cell $sheet unten -index 35,1 -string -style $bottom

  $spreadsheet cell $sheet {Text 90} -index 38,0 -string -style $rotate90
  $spreadsheet cell $sheet {Text 45} -index 38,1 -string -style $rotate45
  $spreadsheet cell $sheet {Calibri 18} -index 38,4 -string -style $font18

  $spreadsheet cell $sheet {12 Zellen} -index 39,3 -string -style $hvcenter

  $spreadsheet cell $sheet {3 Spalten} -index 40,0 -string
  $spreadsheet cell $sheet {3 Zeilen} -index 41,1 -string

  $spreadsheet cell $sheet {this text will be automatically wrapped by excel} -index H27 -style $wrap

  $spreadsheet cell $sheet a -index F21 -style $center
  $spreadsheet cell $sheet a -index F22 -style $center
  $spreadsheet cell $sheet b -index F23 -style $center
  $spreadsheet cell $sheet 1 -index G21 -style $center
  $spreadsheet cell $sheet 2 -index G22 -style $center
  $spreadsheet cell $sheet 2 -index G23 -style $center

  $spreadsheet row $sheet -index 10
  $spreadsheet cell $sheet 3 -index G
  $spreadsheet cell $sheet 5
  $spreadsheet cell $sheet {} -formula G11+H11

  $spreadsheet freeze $sheet 20,5

  $spreadsheet autofilter $sheet 19,5 19,6

  $spreadsheet rowheight $sheet 33 20
  $spreadsheet rowheight $sheet 34 20
  $spreadsheet rowheight $sheet 35 20

  $spreadsheet merge $sheet 40,0 40,2
  $spreadsheet merge $sheet 41,1 43,1
  $spreadsheet merge $sheet 39,3 42,5
}
$spreadsheet write export6.xlsx
$spreadsheet destroy
