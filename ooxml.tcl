#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  $Id: ooxml.tcl,v 1.28 2018/06/20 15:01:38 alex Exp $
#
#  Copyright (C) 2018 Alexander Schoepe, Bochum, DE, <schoepe@users.sourceforge.net>
#  All rights reserved.
#
#  Redistribution and use in source and binary forms, with or without modification,
#  are permitted provided that the following conditions are met:
#
#  1. Redistributions of source code must retain the above copyright notice, this
#     list of conditions and the following disclaimer.
#  2. Redistributions in binary form must reproduce the above copyright notice,
#     this list of conditions and the following disclaimer in the documentation
#     and/or other materials provided with the distribution.
#  3. Neither the name of the project nor the names of its contributors may be used
#     to endorse or promote products derived from this software without specific
#     prior written permission.
#
#  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY
#  EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
#  OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT
#  SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
#  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
#  TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
#  BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
#  CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
#  ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF
#  SUCH DAMAGE.
#


# 
# INDEX and ID are zero based
# 
# 
# BORDERLINESTYLE
#   dashDot | dashDotDot | dashed | dotted | double | hair | medium |
#   mediumDashDot | mediumDashDotDot | mediumDashDotDot | none | slantDashDot | thick | thin
# 
# COLOR
#   0-65
#   Aqua | Black | Blue | BlueRomance | Canary | CarnationPink | Citrus | Cream | DarkSlateBlue | DeepSkyBlue |
#   Eucalyptus | Fuchsia | Gray | Green | Karaka | LavenderBlue | LightCoral | LightCyan | LightSkyBlue | Lime |
#   Lipstick | Maroon | Mauve | MediumTurquoise | Myrtle | Navy | NavyBlue | NightRider | Nobel | Olive |
#   OrangePeel | PeachOrange | Portage | PrussianBlue | Purple | Red | RoyalBlue | SaddleBrown | SafetyOrange |
#   Scampi | Silver | TangerineYellow | Teal | White | Yellow |
#   SystemBackground | SystemForeground
#   RGB
#   ARGB
# 
# DEGREE
#   0-360
# 
# DIAGONALDIRECTION
#   up | down
# 
# HORIZONTAL
#   left | center | right
# 
# PATTERNTYPE
#   darkDown | darkGray | darkGrid | darkHorizontal | darkTrellis | darkUp | darkVertical |
#   gray0625 | gray125 | lightDown | lightGray |
#   lightGrid | lightHorizontal | lightTrellis | lightUp | lightVertical | mediumGray | none | solid
# 
# VERTICAL
#   top | center | bottom
# 
# 
# ::ooxml::Default name value
#   name = path
#
# 
# ::ooxml::RowColumnToString rowcol
#   return name
# 
# 
# ::ooxml::StringToRowColumn name
#   return rowcol
# 
# 
# ::ooxml::CalcColumnWidth numberOfCharacters {maximumDigitWidth 7} {pixelPadding 5}
#   return width
# 
# 
# ::ooxml::xl_sheets file
#   return sheetInformation
# 
# 
# ::ooxml::xl_read file
#   -valuesonly -keylist -sheets PATTERN -sheetnames PATTERN -datefmt FORMAT
#   return workbookData
# 
# 
# ::ooxml::xl_write
# 
#   constructor args
#     -creator CREATOR
#     return class
# 
#   method numberformat args
#     -format FORMAT -general -date -time -datetime -iso8601 -number -decimal -red -separator -fraction -scientific -percent -text -string
#     return NUMFMTID
#
#   method defaultdatestyle STYLEID
# 
#   method font args
#     -list -name NAME -family FAMILY -size SIZE -color COLOR -scheme SCHEME -bold -italic -underline -color COLOR
#     return FONTID
# 
#   method fill args
#     -list -patterntype PATTERNTYPE -fgcolor COLOR -bgcolor COLOR
#     return FILLID
# 
#   method border args
#     -list -leftstyle BORDERLINESTYLE -leftcolor COLOR -rightstyle BORDERLINESTYLE -rightcolor COLOR -topstyle BORDERLINESTYLE -topcolor COLOR
#     -bottomstyle BORDERLINESTYLE -bottomcolor COLOR -diagonalstyle BORDERLINESTYLE -diagonalcolor COLOR -diagonaldirection DIAGONALDIRECTION
#     return BORDERID
# 
#   method style args
#     -list -numfmt NUMFMTID -font FONTID -fill FILLID -border BORDERID -xf XFID -horizontal HORIZONTAL -vertical VERTICAL -rotate DEGREE
#     return STYLEID
# 
#   method worksheet name
#     return SHEETID
# 
#   method column sheet args
#     -index INDEX -to INDEX -width WIDTH -style STYLEID -bestfit -customwidth -string -nozero -calcfit
#     autoincrement of column if INDEX not applied
#     return column
# 
#   method row sheet args
#     -index INDEX -height HEIGHT
#     autoincrement of row if INDEX not applied
#     return row
# 
#   method cell sheet {data {}} args
#     -index INDEX -style STYLEID -formula FORMULA -string -nozero -globalstyle -height HEIGHT
#     autoincrement of column if INDEX not applied
#     return row,column
# 
#   method autofilter sheet indexFrom indexTo
# 
#   method freeze sheet index
# 
#   method presetstyles
#
#   method presetsheets
#
#   method write filename
# 
#
# ::ooxml::tablelist_to_xl lb args
#   -callback CALLBACK -path PATH -file FILENAME -creator CREATOR -name NAME -rootonly -addtimestamp
#   Callback arguments
#     spreadsheet sheet maxcol column title width align sortmode hide
#


package require Tcl 8.6
package require vfs::zip
package require tdom 0.9.0-
package require msgcat


namespace eval ::ooxml {
  namespace export xl_sheets xl_read xl_write

  variable defaults
  variable predefNumFmts
  variable predefColors
  variable predefColorsName
  variable predefColorsARBG
  variable predefBorderLineStyles
  variable predefPatternType

  set defaults(path) {.}

  set defaults(numFmts,start) 166
  set defaults(cols,width) 10.83203125

  # predefined formats
  array set predefNumFmts {
    0 {dt 0 fmt {General}}
    1 {dt 0 fmt {0}}
    2 {dt 0 fmt {0.00}}
    3 {dt 0 fmt {#,##0}}
    4 {dt 0 fmt {#,##0.00}}
    9 {dt 0 fmt {0%}}
    10 {dt 0 fmt {0.00%}}
    11 {dt 0 fmt {0.00E+00}}
    12 {dt 0 fmt {#\ ?/?}}
    13 {dt 0 fmt {#\ ??/??}}
    14 {dt 1 fmt {mm-dd-yy}}
    15 {dt 1 fmt {d-mmm-yy}}
    16 {dt 1 fmt {d-mmm}}
    17 {dt 1 fmt {mmm-yy}}
    18 {dt 1 fmt {h:mm\ AM/PM}}
    19 {dt 1 fmt {h:mm:ss\ AM/PM}}
    20 {dt 1 fmt {h:mm}}
    21 {dt 1 fmt {h:mm:ss}}
    22 {dt 1 fmt {m/d/yy h:mm}}
    37 {dt 0 fmt {#,##0\ ;(#,##0)}}
    38 {dt 0 fmt {#,##0\ ;[Red](#,##0)}}
    39 {dt 0 fmt {#,##0.00;(#,##0.00)}}
    40 {dt 0 fmt {#,##0.00;[Red](#,##0.00)}}
    45 {dt 1 fmt {mm:ss}}
    46 {dt 1 fmt {[h]:mm:ss}}
    47 {dt 1 fmt {mmss.0}}
    48 {dt 0 fmt {##0.0E+0}}
    49 {dt 0 fmt {@}}
  }

  array set predefColors {
    0 {argb 00000000 name Black}
    1 {argb 00FFFFFF name White}
    2 {argb 00FF0000 name Red}
    3 {argb 0000FF00 name Lime}
    4 {argb 000000FF name Blue}
    5 {argb 00FFFF00 name Yellow}
    6 {argb 00FF00FF name Fuchsia}
    7 {argb 0000FFFF name Aqua}
    8 {argb 00000000 name Black}
    9 {argb 00FFFFFF name White}
    10 {argb 00FF0000 name Red}
    11 {argb 0000FF00 name Lime}
    12 {argb 000000FF name Blue}
    13 {argb 00FFFF00 name Yellow}
    14 {argb 00FF00FF name Fuchsia}
    15 {argb 0000FFFF name Aqua}
    16 {argb 00800000 name Maroon}
    17 {argb 00008000 name Green}
    18 {argb 00000080 name Navy}
    19 {argb 00808000 name Olive}
    20 {argb 00800080 name Purple}
    21 {argb 00008080 name Teal}
    22 {argb 00C0C0C0 name Silver}
    23 {argb 00808080 name Gray}
    24 {argb 009999FF name Portage}
    25 {argb 00993366 name Lipstick}
    26 {argb 00FFFFCC name Cream}
    27 {argb 00CCFFFF name LightCyan}
    28 {argb 00660066 name Purple}
    29 {argb 00FF8080 name LightCoral}
    30 {argb 000066CC name NavyBlue}
    31 {argb 00CCCCFF name LavenderBlue}
    32 {argb 00000080 name Navy}
    33 {argb 00FF00FF name Fuchsia}
    34 {argb 00FFFF00 name Yellow}
    35 {argb 0000FFFF name Aqua}
    36 {argb 00800080 name Purple}
    37 {argb 00800000 name Maroon}
    38 {argb 00008080 name Teal}
    39 {argb 000000FF name Blue}
    40 {argb 0000CCFF name DeepSkyBlue}
    41 {argb 00CCFFFF name LightCyan}
    42 {argb 00CCFFCC name BlueRomance}
    43 {argb 00FFFF99 name Canary}
    44 {argb 0099CCFF name LightSkyBlue}
    45 {argb 00FF99CC name CarnationPink}
    46 {argb 00CC99FF name Mauve}
    47 {argb 00FFCC99 name PeachOrange}
    48 {argb 003366FF name RoyalBlue}
    49 {argb 0033CCCC name MediumTurquoise}
    50 {argb 0099CC00 name Citrus}
    51 {argb 00FFCC00 name TangerineYellow}
    52 {argb 00FF9900 name OrangePeel}
    53 {argb 00FF6600 name SafetyOrange}
    54 {argb 00666699 name Scampi}
    55 {argb 00969696 name Nobel}
    56 {argb 00003366 name PrussianBlue}
    57 {argb 00339966 name Eucalyptus}
    58 {argb 00003300 name Myrtle}
    59 {argb 00333300 name Karaka}
    60 {argb 00993300 name SaddleBrown}
    61 {argb 00993366 name Lipstick}
    62 {argb 00333399 name DarkSlateBlue}
    63 {argb 00333333 name NightRider}
    64 {argb {} name SystemForeground}
    65 {argb {} name SystemBackground}
  }
  set predefColorsName {}
  set predefColorsARBG {}
  foreach idx [lsort -integer [array names predefColors]] {
    lappend predefColorsName [dict get $predefColors($idx) name]
    lappend predefColorsARBG [dict get $predefColors($idx) argb]
  }

  set predefPatternType {
    darkDown
    darkGray
    darkGrid
    darkHorizontal
    darkTrellis
    darkUp
    darkVertical
    gray0625
    gray125
    lightDown
    lightGray
    lightGrid
    lightHorizontal
    lightTrellis
    lightUp
    lightVertical
    mediumGray
    none
    solid
  }

  set predefBorderLineStyles {
    dashDot
    dashDotDot
    dashed
    dotted
    double
    hair
    medium
    mediumDashDot
    mediumDashDotDot
    mediumDashDotDot
    none
    slantDashDot
    thick
    thin
  }

  msgcat::mcset ar Book \u0627\u0644\u0643\u062a\u0627\u0628
  msgcat::mcset ar Worksheets "\u0623\u0648\u0631\u0627\u0642 \u0627\u0644\u0639\u0645\u0644"
  msgcat::mcset ar Sheet \u0627\u0644\u0648\u0631\u0642\u0629
  msgcat::mcset cs Book Ses\u030cit
  msgcat::mcset cs Worksheets Listy
  msgcat::mcset cs Sheet List
  msgcat::mcset da Book Mappe
  msgcat::mcset da Worksheets Regneark
  msgcat::mcset da Sheet Ark
  msgcat::mcset de Book Mappe
  msgcat::mcset de Worksheets Arbeitsbl\u00e4tter
  msgcat::mcset de Sheet Blatt
  msgcat::mcset el Book \u0392\u03b9\u03b2\u03bb\u03b9\u0301\u03bf
  msgcat::mcset el Worksheets "\u03a6\u03cd\u03bb\u03bb\u03b1 \u03b5\u03c1\u03b3\u03b1\u03c3\u03af\u03b1\u03c2"
  msgcat::mcset el Sheet \u03a6\u03cd\u03bb\u03bb\u03bf
  msgcat::mcset en Book Book
  msgcat::mcset en Worksheets Worksheets
  msgcat::mcset en Sheet Sheet
  msgcat::mcset es Book Libro
  msgcat::mcset es Worksheets "Hojas de c\u00e1lculo"
  msgcat::mcset es Sheet Hoja
  msgcat::mcset fi Book Tyo\u0308kirja
  msgcat::mcset fi Worksheets Laskentataulukot
  msgcat::mcset fi Sheet Taulukko
  msgcat::mcset fr Book Classeur
  msgcat::mcset fr Worksheets "Feuilles de calcul"
  msgcat::mcset fr Sheet Feuil
  msgcat::mcset he Book \u05d7\u05d5\u05d1\u05e8\u05ea
  msgcat::mcset he Worksheets "\u05d2\u05dc\u05d9\u05d5\u05e0\u05d5\u05ea \u05e2\u05d1\u05d5\u05d3\u05d4"
  msgcat::mcset he Sheet \u05d2\u05d9\u05dc\u05d9\u05d5\u05df
  msgcat::mcset hu Book Munkafu\u0308zet
  msgcat::mcset hu Worksheets Munkalapok
  msgcat::mcset hu Sheet Munkalap
  msgcat::mcset it Book Cartel
  msgcat::mcset it Worksheets "Fogli di lavoro"
  msgcat::mcset it Sheet Foglio
  msgcat::mcset ja Book Book
  msgcat::mcset ja Worksheets \u30ef\u30fc\u30af\u30b7\u30fc\u30c8
  msgcat::mcset ja Sheet Sheet
  msgcat::mcset ko Book "\u1110\u1169\u11bc\u1112\u1161\u11b8 \u1106\u116e\u11ab\u1109\u1165"
  msgcat::mcset ko Worksheets \uc6cc\ud06c\uc2dc\ud2b8
  msgcat::mcset ko Sheet \uc2dc\ud2b8
  msgcat::mcset nl Book Map
  msgcat::mcset nl Worksheets Werkbladen
  msgcat::mcset nl Sheet Blad
  msgcat::mcset no Book Bok
  msgcat::mcset no Worksheets Regneark
  msgcat::mcset no Sheet Ark
  msgcat::mcset pl Book Skoroszyt
  msgcat::mcset pl Worksheets Arkusze
  msgcat::mcset pl Sheet Arkusz
  msgcat::mcset pt Book Livro
  msgcat::mcset pt Worksheets "Folhas de C\u00e1lculo"
  msgcat::mcset pt Sheet Folha
  msgcat::mcset ru Book \u041a\u043d\u0438\u0433\u0430
  msgcat::mcset ru Worksheets \u041b\u0438\u0441\u0442\u044b
  msgcat::mcset ru Sheet \u041b\u0438\u0441\u0442
  msgcat::mcset sl Book Zos\u030cit
  msgcat::mcset sl Worksheets H\u00e1rky
  msgcat::mcset sl Sheet H\u00e1rok
  msgcat::mcset sv Book Bok
  msgcat::mcset sv Worksheets Kalkylblad
  msgcat::mcset sv Sheet Blad
  msgcat::mcset th Book \u0e2a\u0e21\u0e38\u0e14\u0e07\u0e32\u0e19
  msgcat::mcset th Worksheets \u0e40\u0e27\u0e34\u0e23\u0e4c\u0e01\u0e0a\u0e35\u0e15
  msgcat::mcset th Sheet \u0e41\u0e1c\u0e48\u0e19\u0e07\u0e32\u0e19
  msgcat::mcset tr Book Kitap
  msgcat::mcset tr Worksheets "\u00c7al\u0131\u015fma Sayfalar\u0131"
  msgcat::mcset tr Sheet Sayfa
  msgcat::mcset zh Book \u5de5\u4f5c\u7c3f
  msgcat::mcset zh Worksheets \u5de5\u4f5c\u8868
  msgcat::mcset zh Sheet \u5de5\u4f5c\u8868
}


proc ::ooxml::Default { name value } {
  variable defaults

  switch -- $name {
    path {
      set defaults($name) [string trim $value]
      if {$value eq {}} {
	set defaults($name) .
      }
    }
    default {
    } 
  }
}


proc ::ooxml::Getopt { *result declaration to_parse } {
  if {${*result} ne {}} {
    upvar 1 ${*result} result
  }

  set error 0
  set errmsg {} 
  array set result {}

  set optopts {}
  set argcnt [llength $declaration]
  for {set argc 0} {$argc < $argcnt} {incr argc} {
    set opt [lindex $declaration $argc]
    if {[lsearch -exact {. -} [string index $opt 0]] > -1} {
      set error 1
      lappend errmsg "option name can not start with '.' or '-'"
      set result(-error) $error
      set result(-errmsg) $errmsg
      return $error
    }
    if {[regsub -- {\..*$} $opt {} name]} {
      regsub -- (${name}) $opt {} opt
    }
    if {![regsub -- {\.arg\M} $opt {} opt]} {
      dict set optopts $name arg 0
      set result($name) 0
    } else {
      dict set optopts $name arg 1
      incr argc
      if {$argc < $argcnt} {
	set result($name) [lindex $declaration $argc]
      } else {
	set result($name) {}
	set error 1
	lappend errmsg "declaration of '$name' missing default value"
	set result(-error) $error
	set result(-errmsg) $errmsg
	return $error
      }
    }
  }

  set argcnt [llength $to_parse]
  for {set argc 0} {$argc < $argcnt} {incr argc} {
    set opt [lindex $to_parse $argc]
    if {$opt eq {--}} {
      set result(--) [lrange $to_parse ${argc}+1 end]
      break
    } elseif {[string index $opt 0] eq {-}} {
      set opt [string range $opt 1 end]
      if {[dict keys $optopts $opt] eq $opt || [llength [set opt [dict keys $optopts ${opt}*]]] == 1} {
	if {[dict get $optopts $opt arg]} {
	  incr argc
	  if {$argc < $argcnt} {
	    set result($opt) [lindex $to_parse $argc]
	  } else {
	    set error 1
	    lappend errmsg "${opt}: missing argument"
	  }
	} else {
	  set result($opt) 1
	}
      } else {
	set error 1
	lappend errmsg "[lindex $to_parse $argc]: unknown or ambiguous option"
      }
    } else {
      set error 1
      lappend errmsg "${opt}: syntax error"
    }
  }

  if {![info exists result(--)]} {
    set result(--) {}
  }

  set result(-error) $error
  set result(-errmsg) $errmsg

  return $error
}


proc ::ooxml::ScanDateTime { scan {iso8601 0} } {
  set d  1
  set m  1
  set ml {}
  set y  1970
  set H  0
  set M  0
  set S  0
  set F  0

  if {[regexp {^(\d+)\.(\d+)\.(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$} $scan all d m y H M S F x a b] ||
      [regexp {^(\d+)-(\d+)-(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$} $scan all y m d H M S F x a b] ||
      [regexp {^(\d+)-(\w+)-(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$} $scan all d ml y H M S F x a b] ||
      [regexp {^(\d+)/(\d+)/(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$} $scan all m d y H M S F x a b]} {
    scan $y %u y

    if {[string is integer -strict $y] && $y >= 0 && $y <= 2038} {
      switch -- [string tolower $ml] {
	jan -
	ene -
	gen -
	tam {set m 1}
	feb -
	fev -
	fév -
	hel {set m 2}
	mrz -
	mar -
	mär -
	maa {set m 3}
	apr -
	avr -
	abr -
	huh {set m 4}
	mai -
	may -
	mei -
	mag -
	maj -
	tou {set m 5}
	jun -
	jui -
	giu -
	kes {set m 6}
	jul -
	jui -
	lug -
	hei {set m 7}
	aug -
	aou -
	aoû -
	ago -
	elo {set m 8}
	sep -
	set -
	syy {set m 9}
	okt -
	oct -
	out -
	ott -
	lok {set m 10}
	nov -
	mar {set m 11}
	dez -
	dec -
	déc -
	dic -
	des -
	jou {set m 12}
	default { set m [string trimleft $m 0] }
      }

      foreach name {y m d H M S F a b} {
	upvar 0 $name var
	set var [string trimleft $var 0]
	if {![string is integer -strict $var]} {
	  set var 0
	}
      }

      if {$y < 100} {
	if {$y < 50} {
	  incr y 2000
	} else {
	  incr y 1900
	}
      }
      if {$y < 1900} {
	return {}
      }

      set Y [format %04u $y]
      set y [format %02u [expr {$y - int($y / 100) * 100}]]
      set m [format %02u $m]
      set d [format %02u $d]
      set H [format %02u $H]
      set M [format %02u $M]
      set S [format %02u $S]

      if {$iso8601} {
	return [list ${Y}-${m}-${d}T${H}:${M}:${S}]
      }
      return [set ole [expr {[clock scan ${Y}${m}${d}T${H}${M}${S} -gmt 1] / 86400.0 + 25569}]]
    }
  }
  return {}
}


proc ::ooxml::Zip { zipfile directory files } {
  array set v { fd {} base {} toc {} }

  # this code is a rewrite and extension of the zipper code found
  # at http://equi4.com/critlib/ and http://wiki.tcl.tk/36689
  # by Tom Krehbiel 2012 krehbiel.tom at gmail dot com

  proc Initialize { *v file } {
    upvar ${*v} v

    set fd [open $file w]
    set v(fd) $fd
    set v(base) [tell $fd]
    set v(toc) {}
    fconfigure $fd -translation binary -encoding binary
  }

  proc Emit { *v s } {
    upvar ${*v} v

    puts -nonewline $v(fd) $s
  }

  proc DosTime { sec } {
    set f [clock format $sec -format {%Y %m %d %H %M %S} -gmt 1]
    regsub -all { 0(\d)} $f { \1} f
    foreach {Y M D h m s} $f break
    set date [expr {(($Y-1980)<<9) | ($M<<5) | $D}]
    set time [expr {($h<<11) | ($m<<5) | ($s>>1)}]
    return [list $date $time]
  }

  proc AddEntry { *v name contents {date {}} {force 0} } {
    upvar ${*v} v

    if {$date eq {}} {
      set date [clock seconds]
    }
    lassign [DosTime $date] date time
    set flag 0
    set type 0 ;# stored
    set fsize [string length $contents]
    set csize $fsize
    set fnlen [string length $name]
    
    if {$force > 0 && $force != [string length $contents]} {
      set csize $fsize
      set fsize $force
      set type 8 ;# if we're passing in compressed data, it's deflated
    }
    
    if {[catch {zlib crc32 $contents} crc]} {
      set crc 0
    } elseif {$type == 0} {
      set cdata [zlib deflate $contents 9]
      if {[string length $cdata] < [string length $contents]} {
	set contents $cdata
	set csize [string length $cdata]
	set type 8 ;# deflate
      }
    }
    
    lappend v(toc) "[binary format a2c6ssssiiiss4ii PK {1 2 20 0 20 0} $flag $type $time $date $crc $csize $fsize $fnlen {0 0 0 0} 128 [tell $v(fd)]]$name"
    
    Emit v [binary format a2c4ssssiiiss PK {3 4 20 0} $flag $type $time $date $crc $csize $fsize $fnlen 0]
    Emit v $name
    Emit v $contents
  }

  proc AddDirectory { *v name {date {}} {force 0} } {
    upvar ${*v} v

    set name "${name}/"
    if {$date eq {}} {
      set date [clock seconds]
    }
    lassign [DosTime $date] date time
    set flag 0
    set type 0 ;# stored
    set fsize 0
    set csize 0
    set fnlen [string length $name]
    set crc 0
    
    lappend v(toc) "[binary format a2c6ssssiiiss4ii PK {1 2 20 0 20 0} $flag $type $time $date $crc $csize $fsize $fnlen {0 0 0 0} 128 [tell $v(fd)]]$name"
    
    Emit v [binary format a2c4ssssiiiss PK {3 4 20 0} $flag $type $time $date $crc $csize $fsize $fnlen 0]
    Emit v $name
  }

  proc Finalize { *v } {
    upvar ${*v} v

    set pos [tell $v(fd)]
    set ntoc [llength $v(toc)]
    foreach x $v(toc) {
      Emit v $x
    }
    set v(toc) {}
    
    set len [expr {[tell $v(fd)] - $pos}]
    incr pos -$v(base)
    
    Emit v [binary format a2c2ssssiis PK {5 6} 0 0 $ntoc $ntoc $len $pos 0]
    
    close $v(fd)
  }

  Initialize v $zipfile
  foreach file $files {
    regsub {^\./} $file {} to
    set from [file join [file normalize $directory] $to]
    if {[file isfile $from]} {
      set fd [open $from r]
      fconfigure $fd -translation binary -encoding binary
      AddEntry v $to [read $fd] [file mtime $from]
      close $fd
    } elseif {[file isdir $from]} {
      AddDirectory v $to [file mtime $from]
      lappend dirs $file
    }
  }
  Finalize v
}


proc ::ooxml::RowColumnToString { rowcol } {
  proc Column { col } {
    set name {}
    while {$col >= 0} {
      set char [binary format c [expr {($col % 26) + 65}]]
      set name $char$name
      set col [expr {$col / 26 -1}]
    }
    return $name
  }
  lassign [split $rowcol ,] row col
  return [Column $col][incr row 1]
}


proc ::ooxml::StringToRowColumn { name } {
  set row 0
  set col 0
  binary scan [string toupper $name] c* vals
  foreach val $vals {
    if {$val < 58} {
      # 0-9, "0" = 48
      set row [expr {$row * 10 + ($val-48)}]
    } else {
      # A-Z, "A" = 65 (-1 zero based shift)
      set col [expr {$col * 26 + ($val-64)}]
    }
  }
  return [incr row -1],[incr col -1]
}


proc ::ooxml::IndexToString { index } {
  lassign [split $index ,] row col
  if {[string is integer -strict $row] && [string is integer -strict $col] && $row > -1 && $col > -1} {
    return [RowColumnToString $index]
  } else {
    lassign [split [StringToRowColumn $index] ,] row col
    if {[string is integer -strict $row] && [string is integer -strict $col] && $row > -1 && $col > -1} {
      return $index
    }
  }
  return {}
}


proc ::ooxml::CalcColumnWidth { numberOfCharacters {maximumDigitWidth 7} {pixelPadding 5} } {
  return [expr {int(($numberOfCharacters * $maximumDigitWidth + $pixelPadding + 0.0) / $maximumDigitWidth * 256.0) / 256.0}]
}


# Seite 3947
# <xsd:complexType name="CT_Color">
#   <xsd:attribute name="auto" type="xsd:boolean" use="optional"/>
#   <xsd:attribute name="indexed" type="xsd:unsignedInt" use="optional"/>
#   <xsd:attribute name="rgb" type="ST_UnsignedIntHex" use="optional"/>
#   <xsd:attribute name="theme" type="xsd:unsignedInt" use="optional"/>
#   <xsd:attribute name="tint" type="xsd:double" use="optional" default="0.0"/>
# </xsd:complexType>

proc ::ooxml::Color { color } {
  variable predefColors
  variable predefColorsName
  variable predefColorsARBG

  if {[string trim $color] eq {}} {
    return {}
  } elseif {$color in {auto none}} {
    return [list $color 1]
  } elseif {[string is integer -strict $color] && $color >= 0 && $color <= 65} {
    return [list indexed $color]
  } elseif {[set idx [lsearch -exact -nocase [array names predefColors] $color]] && $idx > -1} {
    return [list indexed $idx]
  } elseif {[set idx [lsearch -exact -nocase $predefColorsName $color]] && $idx > -1} {
    return [list indexed $idx]
  }
  if {[string is xdigit -strict $color]} {
    if {[string length $color] == 6} {
      set color 00$color
    }
    if {[set idx [lsearch -exact -nocase $predefColorsARBG $color]] && $idx > -1} {
      return [list indexed $idx]
    } else {
      return [list rgb $color]
    }
  }
  return {}
}


#
# ooxml::xl_sheets
#

proc ::ooxml::xl_sheets { file } {
  set sheets {}

  set mnt [vfs::zip::Mount $file xlsx]

  set rels 0
  if {![catch {open xlsx/xl/_rels/workbook.xml.rels r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} rdoc]} {
      set rels 1
      set relsroot [$rdoc documentElement]
    }
    close $fd
  }

  if {![catch {open xlsx/xl/workbook.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      set root [$doc documentElement]
      set idx -1
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:workbook/X:sheets/X:sheet}] {
	if {[$node hasAttribute sheetId] && [$node hasAttribute name]} {
	  set sheetId [$node getAttribute sheetId]
	  set name [$node getAttribute name]
	  set rid [$node getAttribute r:id]
	  foreach node [$relsroot selectNodes -namespaces [list X [$relsroot namespaceURI]] [subst -nobackslashes -nocommands {/X:Relationships/X:Relationship[@Id="$rid"]}]] {
	    if {[$node hasAttribute Target]} {
	      lappend sheets [incr idx] [list sheetId $sheetId name $name rId $rid]
	    }
	  }
	}
      }
      $doc delete
    }
    close $fd
  }

  if {$rels} {
    $rdoc delete
  }

  vfs::zip::Unmount $mnt xlsx

  return $sheets
}


#
# ooxml::xl_read
#

proc ::ooxml::xl_read { file args } {
  variable predefNumFmts

  array set cellXfs {}
  array set numFmts [array get predefNumFmts]
  array set sharedStrings {}
  set sheets {}

  if {[::ooxml::Getopt opts {valuesonly keylist sheets.arg {} sheetnames.arg {} datefmt.arg {%Y-%m-%d %H:%M:%S} as.arg array} $args]} {
    error $opts(-errmsg)
  }
  if {[string trim $opts(sheets)] eq {} && [string trim $opts(sheetnames)] eq {}} {
    set opts(sheetnames) *
  }


  set mnt [vfs::zip::Mount $file xlsx]

  set rels 0
  if {![catch {open xlsx/xl/_rels/workbook.xml.rels r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} rdoc]} {
      set rels 1
      set relsroot [$rdoc documentElement]
    }
    close $fd
  }

  if {![catch {open xlsx/xl/workbook.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      set root [$doc documentElement]
      set idx -1
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:workbook/X:sheets/X:sheet}] {
	if {[$node hasAttribute sheetId] && [$node hasAttribute name]} {
	  set sheetId [$node getAttribute sheetId]
	  set name [$node getAttribute name]
	  set rid [$node getAttribute r:id]
	  foreach node [$relsroot selectNodes -namespaces [list X [$relsroot namespaceURI]] [subst -nobackslashes -nocommands {/X:Relationships/X:Relationship[@Id="$rid"]}]] {
	    if {[$node hasAttribute Target]} {
	      lappend sheets [incr idx] $sheetId $name $rid [$node getAttribute Target]
	    }
	  }
	}
      }
      $doc delete
    }
    close $fd
  }

  if {$rels} {
    $rdoc delete
  }

  if {![catch {open xlsx/xl/sharedStrings.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      set root [$doc documentElement]
      set idx -1
      foreach shared [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:sst/X:si}] {
	incr idx
	foreach node [$shared selectNodes -namespaces [list X [$shared namespaceURI]] {X:t/text()}] {
	  append sharedStrings($idx) [$node nodeValue]
	}
	foreach node [$shared selectNodes -namespaces [list X [$shared namespaceURI]] {*/X:t/text()}] {
	  append sharedStrings($idx) [$node nodeValue]
	}
      }
      $doc delete
    }
    close $fd
  }


  if {![catch {open xlsx/xl/styles.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      set root [$doc documentElement]
      set idx -1
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:numFmts/X:numFmt}] {
        incr idx
	if {[$node hasAttribute numFmtId] && [$node hasAttribute formatCode]} {
	  set numFmtId [$node getAttribute numFmtId]
	  set formatCode [$node getAttribute formatCode]
	  set datetime 0
	  foreach tag {*y* *m* *d* *h* *s*} {
	    if {[string match -nocase $tag [string map {Black {} Blue {} Cyan {} Green {} Magenta {} Red {} White {} Yellow {}} $formatCode]]} {
	      set datetime 1
	      break
	    }
	  }
	  set numFmts($numFmtId) [list dt $datetime fmt $formatCode]
	}
      }
      set idx -1
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:cellXfs/X:xf}] {
        incr idx
	if {[$node hasAttribute numFmtId]} {
	  set numFmtId [$node getAttribute numFmtId]
	  if {[$node hasAttribute applyNumberFormat]} {
	    set applyNumberFormat [$node getAttribute applyNumberFormat]
	  } else {
	    set applyNumberFormat 0
	  }
	  set cellXfs($idx) [list nfi $numFmtId anf $applyNumberFormat]
	}
      }

      ### READING KNOWN FORMATS AND STYLES ###

      set wb(s,@) {}

      array unset a *
      set a(max) 0
      set wb(s,numFmtsIds) {}
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:numFmts/X:numFmt}] {
        if {[$node hasAttribute numFmtId] && [$node hasAttribute formatCode]} {
	  set wb(s,numFmts,[set idx [$node getAttribute numFmtId]]) [$node getAttribute formatCode]
	  lappend wb(s,numFmtsIds) $idx
	  if {$idx > $a(max)} {
	    set a(max) $idx
	  }
	}
      }
      if {$a(max) < $::ooxml::defaults(numFmts,start)} {
        set a(max) $::ooxml::defaults(numFmts,start)
      }
      lappend wb(s,@) numFmtId [incr a(max)]


      set idx -1
      array unset a *
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:fonts/X:font}] {
	incr idx
	array set a {name {} family {} size {} color {} scheme {} bold 0 italic 0 underline 0 color {}}
	foreach node1 [$node childNodes] {
	  switch -- [$node1 nodeName] {
	    b {
	      set a(bold) 1
	    }
	    i {
	      set a(italic) 1
	    }
	    u {
	      set a(underline) 1
	    }
	    sz {
	      if {[$node1 hasAttribute val]} {
		set a(size) [$node1 getAttribute val]
	      }
	    }
	    color {
	      if {[$node1 hasAttribute auto]} {
		set a(color) [list auto [$node1 getAttribute auto]]
	      } elseif {[$node1 hasAttribute rgb]} {
		set a(color) [list rgb [$node1 getAttribute rgb]]
	      } elseif {[$node1 hasAttribute indexed]} {
		set a(color) [list indexed [$node1 getAttribute indexed]]
	      } elseif {[$node1 hasAttribute theme]} {
		set a(color) [list theme [$node1 getAttribute theme]]
	      }
	    }
	    name {
	      if {[$node1 hasAttribute val]} {
		set a(name) [$node1 getAttribute val]
	      }
	    }
	    family {
	      if {[$node1 hasAttribute val]} {
		set a(family) [$node1 getAttribute val]
	      }
	    }
	    scheme {
	      if {[$node1 hasAttribute val]} {
		set a(scheme) [$node1 getAttribute val]
	      }
	    }
	  }
	}
	set wb(s,fonts,$idx) [array get a]
      }
      lappend wb(s,@) fonts [incr idx]


      set idx -1
      array unset a *
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:fills/X:fill}] {
	incr idx
	array set a {patterntype {} fgcolor {} bgcolor {}}
	foreach node1 [$node childNodes] {
	  switch -- [$node1 nodeName] {
	    patternFill {
	      if {[$node1 hasAttribute patternType]} {
		set a(patterntype) [$node1 getAttribute patternType]
	      }
	      foreach node2 [$node1 childNodes] {
		if {[$node2 nodeName] in { fgColor bgColor}} {
		  if {[$node2 hasAttribute auto]} {
		    set a([string tolower [$node2 nodeName]]) [list auto [$node2 getAttribute auto]]
		  } elseif {[$node2 hasAttribute rgb]} {
		    set a([string tolower [$node2 nodeName]]) [list rgb [$node2 getAttribute rgb]]
		  } elseif {[$node2 hasAttribute indexed]} {
		    set a([string tolower [$node2 nodeName]]) [list indexed [$node2 getAttribute indexed]]
		  } elseif {[$node2 hasAttribute theme]} {
		    set a([string tolower [$node2 nodeName]]) [list theme [$node2 getAttribute theme]]
		  }
		}
	      }
	    }
	  }
	}
	set wb(s,fills,$idx) [array get a]
      }
      lappend wb(s,@) fills [incr idx]


      set idx -1
      unset -nocomplain d
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:borders/X:border}] {
	incr idx
	set d {left {style {} color {}} right {style {} color {}} top {style {} color {}} bottom {style {} color {}} diagonal {style {} color {} direction {}}}
	foreach node1 [$node childNodes] {
	  if {[$node1 hasAttribute style]} {
	    set style [$node1 getAttribute style]
	  } else {
	    set style {}
	  }
	  set color {}
	  foreach node2 [$node1 childNodes] {
	    if {[$node2 nodeName] eq {color}} {
	      if {[$node2 hasAttribute auto]} {
		set color [list auto [$node2 getAttribute auto]]
	      } elseif {[$node2 hasAttribute rgb]} {
		set color [list rgb [$node2 getAttribute rgb]]
	      } elseif {[$node2 hasAttribute indexed]} {
		set color [list indexed [$node2 getAttribute indexed]]
	      } elseif {[$node2 hasAttribute theme]} {
		set color [list theme [$node2 getAttribute theme]]
	      }
	    }
	  }
	  if {[$node1 nodeName] in {left right top bottom diagonal}} {
	    if {$style ne {}} {
	      dict set d [$node1 nodeName] style $style
	    }
	    if {$color ne {}} {
	      dict set d [$node1 nodeName] color $color
	    }
	  }
	}
	if {[$node hasAttribute diagonalUp]} {
	  dict set d diagonal direction up
	} elseif {[$node hasAttribute diagonalDown]} {
	  dict set d diagonal direction down
	}
	set wb(s,borders,$idx) $d
      }
      lappend wb(s,@) borders [incr idx]


      set idx -1
      array unset a *
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:styleSheet/X:cellXfs/X:xf}] {
	incr idx
	array set a {numfmt 0 font 0 fill 0 border 0 xf 0 horizontal {} vertical {} rotate {}}
        if {[$node hasAttribute numFmtId]} {
	  set a(numfmt) [$node getAttribute numFmtId]
	}
        if {[$node hasAttribute fontId]} {
	  set a(font) [$node getAttribute fontId]
	}
        if {[$node hasAttribute fillId]} {
	  set a(fill) [$node getAttribute fillId]
	}
        if {[$node hasAttribute borderId]} {
	  set a(border) [$node getAttribute borderId]
	}
        if {[$node hasAttribute xfId]} {
	  set a(xf) [$node getAttribute xfId]
	}
	foreach node1 [$node childNodes] {
	  switch -- [$node1 nodeName] {
	    alignment {
	      if {[$node1 hasAttribute horizontal]} {
		set a(horizontal) [$node1 getAttribute horizontal]
	      }
	      if {[$node1 hasAttribute vertical]} {
		set a(vertical) [$node1 getAttribute vertical]
	      }
	      if {[$node1 hasAttribute textRotation]} {
		set a(rotate) [$node1 getAttribute textRotation]
	      }
	    }
	  }
	}
	set wb(s,styles,$idx) [array get a]
      }
      lappend wb(s,@) styles [incr idx]

      $doc delete
    }
    close $fd
  }


  ### SHEET AND DATA ###

  array set wb {}

  foreach {sheet sid name rid target} $sheets {
    set read false
    if {$opts(sheets) ne {}} {
      foreach pat $opts(sheets) {
	if {[string match $pat $sheet]} {
	  set read true
	  break
	}
      }
    }
    if {!$read && $opts(sheetnames) ne {}} {
      foreach pat $opts(sheetnames) {
	if {[string match $pat $name]} {
	  set read true
	  break
	}
      }
    }
    if {!$read} continue

    lappend wb(sheets) $sheet
    set wb($sheet,n) $name
    set wb($sheet,max_row) -1
    set wb($sheet,max_column) -1

    if {![catch {open [file join xlsx/xl $target] r} fd]} {
      fconfigure $fd -encoding utf-8
      if {![catch {dom parse [read $fd]} doc]} {
	set root [$doc documentElement]
	set idx -1
	foreach col [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:worksheet/X:cols/X:col}] {
	  incr idx
	  foreach item {min max width style bestFit customWidth} {
	    if {[$col hasAttribute $item]} {
	      switch -- $item {
	        min - max {
		  lappend wb($sheet,col,$idx) [string tolower $item] [expr {[$col getAttribute $item] - 1}]
		}
		default {
		  lappend wb($sheet,col,$idx) [string tolower $item] [$col getAttribute $item]
		}
	      }
	    } else {
	      lappend wb($sheet,col,$idx) [string tolower $item] 0
	    }
	  }
	  lappend wb($sheet,col,$idx) string 0 nozero 0 calcfit 0
	}
	set wb($sheet,cols) [incr idx]
	foreach cell [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:worksheet/X:sheetData/X:row/X:c}] {
	  if {[$cell hasAttribute t]} {
	    set type [$cell getAttribute t]
	  } else {
	    set type n
	  }
	  set value {}
	  set datetime {}
	  switch -- $type {
	    n - b - d - str {
	      # number (default), boolean, iso-date, formula string
	      if {[set node [$cell selectNodes -namespaces [list X [$cell namespaceURI]] X:v/text()]] ne {}} {
		set value [$node nodeValue]
		if {$type eq {n} && [$cell hasAttribute s] && [string is double -strict $value]} {
		  set idx [$cell getAttribute s]
		  if {[dict exists $cellXfs($idx) nfi]} {
		    set numFmtId [dict get $cellXfs($idx) nfi]
		    if {[info exists numFmts($numFmtId)] && [dict exists $numFmts($numFmtId) dt] && [dict get $numFmts($numFmtId) dt]} {
		      set datetime $value
		      catch {clock format [expr {int(($value - 25569) * 86400.0)}] -format $opts(datefmt) -gmt 1} value
		    }
		  }
		} 
	      } else {
		if {![$cell hasAttribute s]} continue
	      }
	    }
	    s {
	      # shared string
	      if {[set node [$cell selectNodes -namespaces [list X [$cell namespaceURI]] X:v/text()]] ne {}} {
		set index [$node nodeValue]
		if {[info exists sharedStrings($index)]} {
		  set value $sharedStrings($index)
		}
	      } else {
		if {![$cell hasAttribute s]} continue
	      }
	    }
	    inlineStr {
	      # inline string
	      if {[set string [$cell selectNodes -namespaces [list X [$cell namespaceURI]] X:is]] ne {}} {
		foreach node [$string selectNodes -namespaces [list X [$string namespaceURI]] {X:t/text()}] {
		  append value [$node nodeValue]
		}
		foreach node [$string selectNodes -namespaces [list X [$string namespaceURI]] {*/X:t/text()}] {
		  append value [$node nodeValue]
		}
	      } else {
		if {![$cell hasAttribute s]} continue
	      }
	    }
	    e {
	      # error
	    }
	  }

	  if {[$cell hasAttribute r]} {
	    if {!$opts(valuesonly)} {
	      set wb($sheet,c,[StringToRowColumn [$cell getAttribute r]]) [$cell getAttribute r]
	    }
	    if {!$opts(valuesonly)} {
	      if {[$cell hasAttribute s]} {
		set wb($sheet,s,[StringToRowColumn [$cell getAttribute r]]) [$cell getAttribute s]
	      }
	    }
	    if {!$opts(valuesonly)} {
	      if {[$cell hasAttribute t]} {
		set wb($sheet,t,[StringToRowColumn [$cell getAttribute r]]) [$cell getAttribute t]
	      }
	    }
	    set wb($sheet,v,[StringToRowColumn [$cell getAttribute r]]) $value
	    if {!$opts(valuesonly) && $datetime ne {}} {
	      set wb($sheet,d,[StringToRowColumn [$cell getAttribute r]]) $datetime
	    }
	    if {!$opts(valuesonly) && [set node [$cell selectNodes -namespaces [list X [$cell namespaceURI]] X:f/text()]] ne {}} {
	      set wb($sheet,f,[StringToRowColumn [$cell getAttribute r]]) [$node nodeValue]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach row [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:worksheet/X:sheetData/X:row}] {
	    if {[$row hasAttribute r] && [$row hasAttribute ht] && [$row hasAttribute customHeight] && [$row getAttribute customHeight] == 1} {
	      dict set wb($sheet,rowheight) [expr {[$row getAttribute r] - 1}] [$row getAttribute ht]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach freeze [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:worksheet/X:sheetViews/X:sheetView/X:pane}] {
	    if {[$freeze hasAttribute topLeftCell] && [$freeze hasAttribute state] && [$freeze getAttribute state] eq {frozen}} {
	      set wb($sheet,freeze) [$freeze getAttribute topLeftCell]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach filter [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:worksheet/X:autoFilter}] {
	    if {[$filter hasAttribute ref]} {
	      lappend wb($sheet,filter) [$filter getAttribute ref]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach merge [$root selectNodes -namespaces [list X [$root namespaceURI]] {/X:worksheet/X:mergeCells/X:mergeCell}] {
	    if {[$merge hasAttribute ref]} {
	      lappend wb($sheet,merge) [$merge getAttribute ref]
	    }
	  }
	}
	$doc delete
      }
      close $fd
    }
  }

  vfs::zip::Unmount $mnt xlsx

  foreach cell [lsort -dictionary [array names wb *,v,*]] {
    lassign [split $cell ,] sheet tag row column
    if {$opts(keylist)} {
      dict lappend wb($sheet,k) $row $column
    }
    if {$row > $wb($sheet,max_row)} {
      set wb($sheet,max_row) $row
    }
    if {$column > $wb($sheet,max_column)} {
      set wb($sheet,max_column) $column
    }
  }
  return [array get wb]
}


#
# ooxml::xl_write
#


oo::class create ooxml::xl_write {
  constructor { args } {
    my variable obj
    my variable cells
    my variable sharedStrings
    my variable fonts
    my variable numFmts
    my variable styles
    my variable fills
    my variable borders
    my variable cols

    if {[::ooxml::Getopt opts {creator.arg {unknown}} $args]} {
      error $opts(-errmsg)
    }

    set obj(blockPreset) 0

    set obj(encoding) utf-8
    dom setResultEncoding $obj(encoding)
    set obj(indent) none

    set obj(creator) $opts(creator)
    set obj(created) [clock format [clock seconds] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]

    set obj(sheets) 0
    array set sheets {}

    set obj(sharedStrings) 0
    set sharedStrings {}

    set obj(numFmts) $::ooxml::defaults(numFmts,start)
    array set numFmts {}

    set obj(borders) 1
    set borders(0) {left {style {} color {}} right {style {} color {}} top {style {} color {}} bottom {style {} color {}} diagonal {style {} color {} direction {}}}

    set obj(fills) 2
    set fills(0) {patterntype none fgcolor {} bgcolor {}}
    set fills(1) {patterntype gray125 fgcolor {} bgcolor {}}

    set obj(fonts) 1
    set fonts(0) {name Calibri family 2 size 12 color {theme 1} scheme minor bold 0 italic 0 underline 0 color {}}

    set obj(styles) 1
    set styles(0) {numfmt 0 font 0 fill 0 border 0 xf 0 horizontal {} vertical {} rotate {}}

    set obj(cols) 0
    array set cols {}

    set obj(calcChain) 0

    set obj(defaultdatestyle) 0

    array set cells {}

    return 0
  }

  destructor {
    return 0
  }

  method numberformat { args } {
    my variable obj
    my variable numFmts

    if {[::ooxml::Getopt opts {list format.arg {} general date time datetime iso8601 number decimal red separator fraction scientific percent string text} $args]} {
      error $opts(-errmsg)
    }

    if {$opts(list)} {
      array set tmp [array get ::ooxml::predefNumFmts]
      array set tmp [array get numFmts]
      return [array get tmp]
    }

    set obj(blockPreset) 1

    if {$opts(general)} {
      return 0
    }
    if {$opts(date)} {
      return 14
    }
    if {$opts(time)} {
      return 20
    }
    if {$opts(number)} {
      if {$opts(separator)} {
	if {$opts(red)} {
	  return 38
	} else {
	  return 3
	}
      } else {
	if {$opts(red)} {
	  return -1
	} else {
	  return 1
	}
      }
    }
    if {$opts(decimal)} {
      if {$opts(percent)} {
        return 10
      }
      if {$opts(separator)} {
	if {$opts(red)} {
	  return 40
	} else {
	  return 4
	}
      } else {
	if {$opts(red)} {
	  return -1
	} else {
	  return 2
	}
      }
    }
    if {$opts(fraction)} {
      return 12
    }
    if {$opts(scientific)} {
      return 11
    }
    if {$opts(percent)} {
      return 9
    }
    if {$opts(text) || $opts(string)} {
      return 49
    }

    if {$opts(datetime)} {
      set opts(format) {dd/mm/yyyy\ hh:mm;@}
    }
    if {$opts(iso8601)} {
      set opts(format) {yyyy\-mm\-dd\ hh:mm:ss;@}
    }

    foreach idx [array names ::ooxml::predefNumFmts] {
      if {[dict get $::ooxml::predefNumFmts($idx) fmt] eq $opts(format)} {
        return $idx
      }
    }
    foreach idx [array names numFmts] {
      if {$numFmts($idx) eq $opts(format)} {
        return $idx
      }
    }

    if {$opts(format) eq {}} {
      return -1
    }

    set idx $obj(numFmts)
    set numFmts($idx) $opts(format)
    incr obj(numFmts)

    return $idx
  }

  method defaultdatestyle { style } {
    my variable obj

    set obj(defaultdatestyle) $style
  }

  method font { args } {
    my variable obj
    my variable fonts

    array set a $fonts(0)
    if {[::ooxml::Getopt opts [list list name.arg $a(name) family.arg $a(family) size.arg $a(size) color $a(color) scheme $a(scheme) bold italic underline color.arg {}] $args]} {
      error $opts(-errmsg)
    }

    if {$opts(list)} {
      return [array get fonts]
    }

    set obj(blockPreset) 1

    if {$opts(name) eq {}} {
      set opts(name) $a(name)
    }
    if {![string is integer -strict $opts(family)] || $opts(family) < 0} {
      set opts(family) $a(family)
    }
    if {![string is integer -strict $opts(size)] || $opts(size) < 0} {
      set opts(size) $a(size)
    }
    if {$opts(scheme) ni {major minor none}} {
      set opts(scheme) $a(scheme)
    }
    set opts(color) [::ooxml::Color $opts(color)]

    foreach idx [lsort -integer [array names fonts]] {
      array set a $fonts($idx)
      set found 1
      foreach name [array names a] {
        if {$a($name) ne $opts($name)} {
	  set found 0
	  break
	}
      }
      if {$found} {
        return $idx
      }
    }

    set fonts($obj(fonts)) {}
    foreach item {name family size bold italic underline color scheme} {
      lappend fonts($obj(fonts)) $item $opts($item)
    }
    set idx $obj(fonts)
    incr obj(fonts)
    return $idx
  }

  method fill { args } {
    my variable obj
    my variable fills

    if {[::ooxml::Getopt opts {list patterntype.arg none fgcolor.arg {} bgcolor.arg {}} $args]} {
      error $opts(-errmsg)
    }

    if {$opts(list)} {
      return [array get fills]
    }

    set obj(blockPreset) 1

    if {$opts(patterntype) ni $::ooxml::predefPatternType} {
      set opts(patterntype) none
    }
    set opts(fgcolor) [::ooxml::Color $opts(fgcolor)]
    set opts(bgcolor) [::ooxml::Color $opts(bgcolor)]

    foreach idx [lsort -integer [array names fills]] {
      array set a $fills($idx)
      set found 1
      foreach name [array names a] {
        if {$a($name) ne $opts($name)} {
	  set found 0
	  break
	}
      }
      if {$found} {
        return $idx
      }
    }

    set fills($obj(fills)) {}
    foreach item {patterntype fgcolor bgcolor} {
      lappend fills($obj(fills)) $item $opts($item)
    }
    set idx $obj(fills)
    incr obj(fills)
    
    return $idx
  }

  method border { args } {
    my variable obj
    my variable borders

    if {[::ooxml::Getopt opts {list leftstyle.arg {} leftcolor.arg {} rightstyle.arg {} rightcolor.arg {} topstyle.arg {} topcolor.arg {} bottomstyle.arg {} bottomcolor.arg {} diagonalstyle.arg {} diagonalcolor.arg {} diagonaldirection.arg {}} $args]} {
      error $opts(-errmsg)
    }

    if {$opts(list)} {
      return [array get borders]
    }

    set obj(blockPreset) 1

    if {$opts(leftstyle) ni $::ooxml::predefBorderLineStyles || $opts(leftstyle) eq {none}} {
      set opts(leftstyle) {}
    }
    set opts(leftcolor) [::ooxml::Color $opts(leftcolor)]
    if {$opts(rightstyle) ni $::ooxml::predefBorderLineStyles || $opts(rightstyle) eq {none}} {
      set opts(rightstyle) {}
    }
    set opts(rightcolor) [::ooxml::Color $opts(rightcolor)]
    if {$opts(topstyle) ni $::ooxml::predefBorderLineStyles || $opts(topstyle) eq {none}} {
      set opts(topstyle) {}
    }
    set opts(topcolor) [::ooxml::Color $opts(topcolor)]
    if {$opts(bottomstyle) ni $::ooxml::predefBorderLineStyles || $opts(bottomstyle) eq {none}} {
      set opts(bottomstyle) {}
    }
    set opts(bottomcolor) [::ooxml::Color $opts(bottomcolor)]
    if {$opts(diagonalstyle) ni $::ooxml::predefBorderLineStyles || $opts(diagonalstyle) eq {none}} {
      set opts(diagonalstyle) {}
    }
    set opts(diagonalcolor) [::ooxml::Color $opts(diagonalcolor)]
    if {$opts(diagonaldirection) ni {up down}} {
      set opts(diagonaldirection) {}
    }
    switch -- $opts(diagonaldirection) {
      up {
	set opts(diagonaldirection) diagonalUp
      }
      down {
	set opts(diagonaldirection) diagonalDown
      }
      default {
	set opts(diagonaldirection) {}
      }
    }

    dict set tmp left style $opts(leftstyle)
    dict set tmp left color $opts(leftcolor)
    dict set tmp right style $opts(rightstyle)
    dict set tmp right color $opts(rightcolor)
    dict set tmp top style $opts(topstyle)
    dict set tmp top color $opts(topcolor)
    dict set tmp bottom style $opts(bottomstyle)
    dict set tmp bottom color $opts(bottomcolor)
    dict set tmp diagonal style $opts(diagonalstyle)
    dict set tmp diagonal color $opts(diagonalcolor)
    dict set tmp diagonal direction $opts(diagonaldirection)

    foreach idx [lsort -integer [array names borders]] {
      set found 1
      foreach key [dict keys $tmp] {
	foreach subkey [dict keys [dict get $tmp $key]] {
	  if {[dict get $borders($idx) $key $subkey] ne [dict get $tmp $key $subkey]} {
	    set found 0
	    break
	  }
	}
      }
      if {$found} {
        return $idx
      }
    }

    set borders($obj(borders)) $tmp
    set idx $obj(borders)
    incr obj(borders)
    
    return $idx
  }

  method style { args } {
    my variable obj
    my variable styles

    if {[::ooxml::Getopt opts {list numfmt.arg 0 font.arg 0 fill.arg 0 border.arg 0 xf.arg 0 horizontal.arg {} vertical.arg {} rotate.arg {}} $args]} {
      error $opts(-errmsg)
    }

    if {$opts(list)} {
      return [array get styles]
    }

    set obj(blockPreset) 1
    
    if {![string is integer -strict $opts(numfmt)] || $opts(numfmt) < 0} {
      set opts(numfmt) 0
    }
    if {![string is integer -strict $opts(font)] || $opts(font) < 0} {
      set opts(font) 0
    }
    if {![string is integer -strict $opts(fill)] || $opts(fill) < 0} {
      set opts(fill) 0
    }
    if {![string is integer -strict $opts(border)] || $opts(border) < 0} {
      set opts(border) 0
    }
    if {![string is integer -strict $opts(xf)] || $opts(xf) < 0} {
      set opts(xf) 0
    }
    if {$opts(horizontal) ni {right center left}} {
      set opts(horizontal) {}
    }
    if {$opts(vertical) ni {top center bottom}} {
      set opts(vertical) {}
    }
    if {![string is integer -strict $opts(rotate)] || $opts(rotate) < 0 || $opts(rotate) > 360} {
      set opts(rotate) {}
    }

    foreach idx [lsort -integer [array names styles]] {
      array set a $styles($idx)
      set found 1
      foreach name [array names a] {
        if {$a($name) ne $opts($name)} {
	  set found 0
	  break
	}
      }
      if {$found} {
        return $idx
      }
    }

    set styles($obj(styles)) {}
    foreach item {numfmt font fill border xf horizontal vertical rotate} {
      lappend styles($obj(styles)) $item $opts($item)
    }
    set idx $obj(styles)
    incr obj(styles)
    return $idx
  }

  method worksheet { name } {
    my variable obj

    incr obj(sheets)
    set obj(callRow,$obj(sheets)) 0
    set obj(sheet,$obj(sheets)) $name
    set obj(gCol,$obj(sheets)) -1
    set obj(row,$obj(sheets)) -1
    set obj(col,$obj(sheets)) -1
    set obj(dminrow,$obj(sheets)) 4294967295
    set obj(dmaxrow,$obj(sheets)) 0
    set obj(dmincol,$obj(sheets)) 4294967295
    set obj(dmaxcol,$obj(sheets)) 0
    set obj(autofilter,$obj(sheets)) {}
    set obj(freeze,$obj(sheets)) {}
    set obj(merge,$obj(sheets)) {}
    set obj(rowHeight,$obj(sheets)) {}

    return $obj(sheets)
  }

  method column { sheet args } {
    my variable obj
    my variable cols

    if {[::ooxml::Getopt opts {index.arg {} to.arg {} width.arg {} style.arg {} bestfit customwidth string nozero calcfit} $args]} {
      error $opts(-errmsg)
    }

    lassign [split $opts(index) ,] row col
    if {[string is integer -strict $opts(index)] && $opts(index) > -1} {
      set obj(gCol,$sheet) $opts(index)
    } elseif {[string is integer -strict $col] && $col > -1} {
      set obj(gCol,$sheet) $col
    } elseif {[string trim $opts(index)] eq {}} {
      incr obj(gCol,$sheet)
    }
    set opts(index) $obj(gCol,$sheet)

    if {$opts(to) eq {}} {
      set opts(to) $opts(index)
    } else {
      set opts(to) [::ooxml::IndexToString $opts(to)]
    }
    set opts(index) [lindex [split $opts(index) ,] end]
    set opts(to) [lindex [split $opts(to) ,] end]
    if {[string is integer -strict $opts(index)]} {
      set obj(gCol,$sheet) $opts(index)
    }
    
    if {![string is double -strict $opts(width)] || $opts(width) < 0} {
      set opts(width) {}
    }
    if {![string is integer -strict $opts(style)] || $opts(style) < 0} {
      set opts(style) {}
    }

    if {$opts(width) ne {} || ([string is integer -strict $opts(style)] && $opts(style) > 0) || $opts(bestfit) > 0} {
      if {$opts(width) eq {}} {
        set opts(width) $::ooxml::defaults(cols,width)
      }
      set cols($sheet,$opts(index)) [list min $opts(index) max $opts(to) width $opts(width) style $opts(style) bestfit $opts(bestfit) customwidth $opts(customwidth) string $opts(string) nozero $opts(nozero)]
    }
    set obj($sheet,cols) [llength [array names cols $sheet,*]]

    return $obj(gCol,$sheet)
  }

  method row { sheet args } {
    my variable obj

    if {[::ooxml::Getopt opts {index.arg {} height.arg {}} $args]} {
      error $opts(-errmsg)
    }

    if {![string is integer -strict $opts(height)] || $opts(height) < 1 || $opts(height) > 1024} {
      set opts(height) {}
    }
    if {[string is integer -strict $opts(index)] && $opts(index) > -1} {
      set obj(callRow,$obj(sheets)) 1
      set obj(col,$obj(sheets)) -1
      set obj(row,$sheet) $opts(index)
      if {$opts(height) ne {}} {
        dict set obj(rowHeight,$sheet) $obj(row,$sheet) $opts(height)
      }
      return $obj(row,$sheet)
    }
    if {[string trim $opts(index)] eq {}} {
      set obj(callRow,$obj(sheets)) 1
      set obj(col,$obj(sheets)) -1
      incr obj(row,$sheet)
      if {$opts(height) ne {}} {
        dict set obj(rowHeight,$sheet) $obj(row,$sheet) $opts(height)
      }
      return $obj(row,$sheet)
    }
    return -1
  }

  method rowheight { sheet row height } {
    my variable obj

    if {![string is integer -strict $row] || ![string is integer -strict $height] || $height < 1 || $height > 1024} {
      return -1
    }

    dict set obj(rowHeight,$sheet) $row $height
    return $row
  }

  method cell { sheet {data {}} args } {
    my variable obj
    my variable cells
    my variable cols

    if {[::ooxml::Getopt opts {index.arg {} style.arg 0 formula.arg {} string nozero globalstyle height.arg {}} $args]} {
      error $opts(-errmsg)
    }

    if {!$obj(callRow,$obj(sheets))} {
      set obj(callRow,$obj(sheets)) 1
      incr obj(row,$sheet)
    }

    lassign [split $opts(index) ,] row col
    if {[string is integer -strict $opts(index)] && $opts(index) > -1} {
      set obj(col,$sheet) $opts(index)
    } elseif {[string is integer -strict $row] && [string is integer -strict $col] && $row > -1 && $col > -1} {
      set obj(row,$sheet) $row
      set obj(col,$sheet) $col
    } elseif {[string trim $opts(index)] eq {}} {
      incr obj(col,$sheet)
    }
    if {$obj(row,$sheet) < 0 || $obj(col,$sheet) < 0} {
      return -1
    }

    if {$opts(globalstyle) && [string is integer -strict $opts(style)] && $opts(style) < 1} {
      if {[info exists cols($sheet,$obj(col,$sheet))] && [dict get $cols($sheet,$obj(col,$sheet)) style] > 0} {
        set opts(style) [dict get $cols($sheet,$obj(col,$sheet)) style]
      }
    }
    if {$opts(string) == 0 && [info exists cols($sheet,$obj(col,$sheet))] && [dict get $cols($sheet,$obj(col,$sheet)) string] == 1} {
      set opts(string) 1
    }
    if {$opts(nozero) == 0 && [info exists cols($sheet,$obj(col,$sheet))] && [dict get $cols($sheet,$obj(col,$sheet)) nozero] == 1} {
      set opts(nozero) 1
    }

    set cell ${sheet},$obj(row,$sheet),$obj(col,$sheet)
    set cells($cell) {}

    if {[string is integer -strict $opts(height)] && $opts(height) > 0 && $opts(height) < 1024} {
      dict set obj(rowHeight,$sheet) $obj(row,$sheet) $opts(height)
    }

    set data [string trimright $data]
    if {$opts(nozero) && [string is double -strict $data] && $data == 0} {
      set data {}
    }

    if {$opts(string)} {
      set type s
    } elseif {[set datetime [::ooxml::ScanDateTime $data]] ne {}} {
      set type n
      set data $datetime
      if {[string is integer -strict $opts(style)] && $opts(style) < 1} {
        set opts(style) $obj(defaultdatestyle)
      }
    } elseif {[string is double -strict $data]} {
      set type n
      set data [string trim $data]
      if {$data in {Inf infinity NaN -NaN} || $opts(string)} {
	set type s
      }
    } else {
      set type s
    }

    if {[string is integer -strict $opts(style)] && $opts(style) > 0} {
      lappend cells($cell) s $opts(style)
    }
    if {[string trim $opts(formula)] ne {}} {
      lappend cells($cell) t $type
      lappend cells($cell) f $opts(formula)
    } else {
      lappend cells($cell) v $data t $type
    }

    if {[string trim $data] eq {} && [string trim $opts(formula)] eq {} && ![string is integer -strict $opts(style)] && $opts(style) < 1} {
      unset -nocomplain cells($cell)
    } else {
      if {$obj(row,$sheet) < $obj(dminrow,$sheet)} {
	set obj(dminrow,$sheet) $obj(row,$sheet)
      }
      if {$obj(row,$sheet) > $obj(dmaxrow,$sheet)} {
	set obj(dmaxrow,$sheet) $obj(row,$sheet)
      }
      if {$obj(col,$sheet) < $obj(dmincol,$sheet)} {
	set obj(dmincol,$sheet) $obj(col,$sheet)
      }
      if {$obj(col,$sheet) > $obj(dmaxcol,$sheet)} {
	set obj(dmaxcol,$sheet) $obj(col,$sheet)
      }
    }
    
    return $obj(row,$sheet),$obj(col,$sheet)
  }

  method autofilter { sheet indexFrom indexTo } {
    my variable obj

    set indexFrom [::ooxml::IndexToString $indexFrom]
    set indexTo [::ooxml::IndexToString $indexTo]
    if {$indexFrom ne {} && $indexTo ne {}} {
      set obj(autofilter,$sheet) $indexFrom:$indexTo
      return 0
    }
    return 1
  }

  method freeze { sheet index } {
    my variable obj

    set index [::ooxml::IndexToString $index]
    if {$index ne {}} {
      set obj(freeze,$sheet) $index
      return 0
    }
    return 1
  }

  method merge { sheet indexFrom indexTo } {
    my variable obj

    set indexFrom [::ooxml::IndexToString $indexFrom]
    set indexTo [::ooxml::IndexToString $indexTo]
    if {$indexFrom ne {} && $indexTo ne {} && "$indexFrom:$indexTo" ni $obj(merge,$sheet)} {
      lappend obj(merge,$sheet) $indexFrom:$indexTo
      return 0
    }
    return 1
  }

  method presetstyles { valName args } {
    my variable obj
    my variable cols
    my variable fonts
    my variable numFmts
    my variable styles
    my variable fills
    my variable borders

    if {$obj(blockPreset)} {
      return 1
    }

    upvar $valName a

    if {[info exists a(s,@)]} {
      set obj(blockPreset) 1

      if {[dict exists $a(s,@) numFmtId]} {
	set obj(numFmts) [dict get $a(s,@) numFmtId]
	foreach idx $a(s,numFmtsIds) {
	  if {[info exists a(s,numFmts,$idx)]} {
	    set numFmts($idx) $a(s,numFmts,$idx)
	  }
	}
      }

      foreach item {fonts fills borders styles} {
        if {[dict exists $a(s,@) $item]} {
	  upvar 0 $item ad
	  for {set idx 0} {$idx < [dict get $a(s,@) $item]} {incr idx} {
	    if {[info exists a(s,$item,$idx)]} {
	      set ad($idx) $a(s,$item,$idx)
	    }
	  }
	  set obj($item) [dict get $a(s,@) $item]
	}
      }
    }

    foreach sheet $a(sheets) {
      for {set idx 0} {$idx < $a($sheet,cols)} {incr idx} {
	if {[info exists a($sheet,col,$idx)]} {
	  set cols([expr {$sheet + 1}],$idx) $a($sheet,col,$idx)
	}
      }
      set obj([expr {$sheet + 1}],cols) [llength [array names cols]]
    }

    return 0
  }

  method presetsheets { valName args } {
    upvar $valName a

    # [self object] -> my
    if {[info exists a(sheets)]} {
      foreach sheet $a(sheets) {
	if {[set currentSheet [my worksheet $a($sheet,n)]] > -1} {
	  dict set a(sheetmap) $sheet $currentSheet
	  foreach item [lsort -dictionary [array names a $sheet,v,*]] {
	    lassign [split $item ,] sheet tag row col
	    set options [list -index $row,$col]
	    if {[info exists a($sheet,f,$row,$col)]} {
	      lappend options -formula $a($sheet,f,$row,$col)
	    }
	    if {[info exists a($sheet,t,$row,$col)] && $a($sheet,t,$row,$col) eq {s}} {
	      lappend options -string
	    }
	    if {[info exists a($sheet,s,$row,$col)]} {
	      lappend options -style $a($sheet,s,$row,$col)
	    }
	    my cell $currentSheet $a($item) {*}$options
	  }
	  if {[info exists a($sheet,rowheight)]} {
	    foreach {row height} $a($sheet,rowheight) {
	      my rowheight $currentSheet $row $height
	    }
	  }
	  if {[info exists a($sheet,freeze)]} {
	    my freeze $currentSheet $a($sheet,freeze)
	  }
	  if {[info exists a($sheet,filter)]} {
	    foreach item $a($sheet,filter) {
	      my autofilter $currentSheet {*}[split $item :]
	    }
	  }
	  if {[info exists a($sheet,merge)]} {
	    foreach item $a($sheet,merge) {
	      my merge $currentSheet {*}[split $item :]
	    }
	  }
	}
      }
    }
  }

  method debug { args } {
    foreach item $args {
      catch {
	my variable $item
	parray $item
      }
    }
  }

  method write { file args } {
    my variable obj
    my variable cells
    my variable sharedStrings
    my variable fonts
    my variable numFmts
    my variable styles
    my variable fills
    my variable borders
    my variable cols

    if {[::ooxml::Getopt opts {holdcontainerdirectory} $args]} {
      error $opts(-errmsg)
    }

    foreach {n v} [array get cells] {
      if {[dict exists $v t] && [dict get $v t] eq {s} && [dict exists $v v] && [dict get $v v] ne {}} {
	if {[set pos [lsearch -exact $sharedStrings [dict get $v v]]] == -1} {
	  lappend sharedStrings [dict get $v v]
	  set pos [lsearch -exact $sharedStrings [dict get $v v]]
	}
	set obj(sharedStrings) 1
	dict set cells($n) v $pos
      }
    }
    unset -nocomplain n v

    # _rels/.rels
    set doc [set obj(doc,_rels/.rels) [dom createDocument Relationships]]
    set root [$doc documentElement]
    $root setAttribute xmlns http://schemas.openxmlformats.org/package/2006/relationships

    set rId 0

    $root appendChild [set node0 [$doc createElement Relationship]]
      $node0 setAttribute Id rId[incr rId]
      $node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
      $node0 setAttribute Target xl/workbook.xml

    $root appendChild [set node0 [$doc createElement Relationship]]
      $node0 setAttribute Id rId[incr rId]
      $node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties
      $node0 setAttribute Target docProps/app.xml

    $root appendChild [set node0 [$doc createElement Relationship]]
      $node0 setAttribute Id rId[incr rId]
      $node0 setAttribute Type http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
      $node0 setAttribute Target docProps/core.xml


    # [Content_Types].xml
    set doc [set obj(doc,\[Content_Types\].xml) [dom createDocument Types]]
    set root [$doc documentElement]
    $root setAttribute xmlns http://schemas.openxmlformats.org/package/2006/content-types

    $root appendChild [set node0 [$doc createElement Default]]
      $node0 setAttribute Extension xml
      $node0 setAttribute ContentType application/xml

    $root appendChild [set node0 [$doc createElement Default]]
      $node0 setAttribute Extension rels
      $node0 setAttribute ContentType application/vnd.openxmlformats-package.relationships+xml

    $root appendChild [set node0 [$doc createElement Override]]
      $node0 setAttribute PartName /xl/workbook.xml
      $node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml

    for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
      $root appendChild [set node0 [$doc createElement Override]]
	$node0 setAttribute PartName /xl/worksheets/sheet${ws}.xml
	$node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml
    }

    $root appendChild [set node0 [$doc createElement Override]]
      $node0 setAttribute PartName /xl/theme/theme1.xml
      $node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.theme+xml

    $root appendChild [set node0 [$doc createElement Override]]
      $node0 setAttribute PartName /xl/styles.xml
      $node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml

    if {$obj(sharedStrings) > 0} {
      $root appendChild [set node0 [$doc createElement Override]]
	$node0 setAttribute PartName /xl/sharedStrings.xml
	$node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml
    }

    if {$obj(calcChain)} {
      $root appendChild [set node0 [$doc createElement Override]]
	$node0 setAttribute PartName /xl/calcChain.xml
	$node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml
    }

    $root appendChild [set node0 [$doc createElement Override]]
      $node0 setAttribute PartName /docProps/core.xml
      $node0 setAttribute ContentType application/vnd.openxmlformats-package.core-properties+xml

    $root appendChild [set node0 [$doc createElement Override]]
      $node0 setAttribute PartName /docProps/app.xml
      $node0 setAttribute ContentType application/vnd.openxmlformats-officedocument.extended-properties+xml


    # docProps/app.xml
    set doc [set obj(doc,docProps/app.xml) [dom createDocument Properties]]
    set root [$doc documentElement]
    $root setAttribute xmlns http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
    $root setAttribute xmlns:vt http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes

    $root appendChild [set node0 [$doc createElement Application]]
      $node0 appendChild [$doc createTextNode {Tcl - Office Open XML - Spreadsheet}]

    $root appendChild [set node0 [$doc createElement DocSecurity]]
      $node0 appendChild [$doc createTextNode 0]

    $root appendChild [set node0 [$doc createElement ScaleCrop]]
      $node0 appendChild [$doc createTextNode false]

    $root appendChild [set node0 [$doc createElement HeadingPairs]]
      set node1 [$node0 appendChild [$doc createElement vt:vector]]
	$node1 setAttribute size 2 baseType variant
	set node2 [$node1 appendChild [$doc createElement vt:variant]]
	  set node3 [$node2 appendChild [$doc createElement vt:lpstr]]
	    $node3 appendChild [$doc createTextNode [msgcat::mc Worksheets]]
	set node2 [$node1 appendChild [$doc createElement vt:variant]]
	  set node3 [$node2 appendChild [$doc createElement vt:i4]]
	    $node3 appendChild [$doc createTextNode 3]

    $root appendChild [set node0 [$doc createElement TitlesOfParts]]
      set node1 [$node0 appendChild [$doc createElement vt:vector]]
	$node1 setAttribute size $obj(sheets) baseType lpstr
	for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
	  set node2 [$node1 appendChild [$doc createElement vt:lpstr]]
	    $node2 appendChild [$doc createTextNode [msgcat::mc Sheet]$ws]
	}

    $root appendChild [set node0 [$doc createElement Company]]

    $root appendChild [set node0 [$doc createElement LinksUpToDate]]
      $node0 appendChild [$doc createTextNode false]

    $root appendChild [set node0 [$doc createElement SharedDoc]]
      $node0 appendChild [$doc createTextNode false]

    $root appendChild [set node0 [$doc createElement HyperlinksChanged]]
      $node0 appendChild [$doc createTextNode false]

    $root appendChild [set node0 [$doc createElement AppVersion]]
      $node0 appendChild [$doc createTextNode 1.0]


    # docProps/core.xml
    set doc [set obj(doc,docProps/core.xml) [dom createDocument cp:coreProperties]]
    set root [$doc documentElement]
    $root setAttribute xmlns:cp http://schemas.openxmlformats.org/package/2006/metadata/core-properties
    $root setAttribute xmlns:dc http://purl.org/dc/elements/1.1/
    $root setAttribute xmlns:dcterms http://purl.org/dc/terms/
    $root setAttribute xmlns:dcmitype http://purl.org/dc/dcmitype/
    $root setAttribute xmlns:xsi http://www.w3.org/2001/XMLSchema-instance

    $root appendChild [set node0 [$doc createElement dc:creator]]
      $node0 appendChild [$doc createTextNode $obj(creator)]

    $root appendChild [set node0 [$doc createElement cp:lastModifiedBy]]
      $node0 appendChild [$doc createTextNode $obj(creator)]

    $root appendChild [set node0 [$doc createElement dcterms:created]]
      $node0 setAttribute xsi:type dcterms:W3CDTF
      $node0 appendChild [$doc createTextNode $obj(created)]

    $root appendChild [set node0 [$doc createElement dcterms:modified]]
      $node0 setAttribute xsi:type dcterms:W3CDTF
      $node0 appendChild [$doc createTextNode $obj(created)]


    # xl/_rels/workbook.xml.rels
    set doc [set obj(doc,xl/_rels/workbook.xml.rels) [dom createDocument Relationships]]
    set root [$doc documentElement]
    $root setAttribute xmlns http://schemas.openxmlformats.org/package/2006/relationships

    for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
      $root appendChild [set node0 [$doc createElement Relationship]]
	$node0 setAttribute Id rId$ws
	$node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet
	$node0 setAttribute Target worksheets/sheet${ws}.xml
    }
    set rId [incr ws -1]

    $root appendChild [set node0 [$doc createElement Relationship]]
      $node0 setAttribute Id rId[incr rId]
      $node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme
      $node0 setAttribute Target theme/theme1.xml

    $root appendChild [set node0 [$doc createElement Relationship]]
      $node0 setAttribute Id rId[incr rId]
      $node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles
      $node0 setAttribute Target styles.xml

    if {$obj(sharedStrings) > 0} {
      $root appendChild [set node0 [$doc createElement Relationship]]
	$node0 setAttribute Id rId[incr rId]
	$node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings
	$node0 setAttribute Target sharedStrings.xml
    }

    if $obj(calcChain) {
      $root appendChild [set node0 [$doc createElement Relationship]]
	$node0 setAttribute Id rId[incr rId]
	$node0 setAttribute Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain
      $node0 setAttribute Target calcChain.xml
    }


    # xl/sharedStrings.xml
    if {$obj(sharedStrings) > 0} {
      set doc [set obj(doc,xl/sharedStrings.xml) [dom createDocument sst]]
      set root [$doc documentElement]
      $root setAttribute xmlns http://schemas.openxmlformats.org/spreadsheetml/2006/main
      $root setAttribute count [llength $sharedStrings] uniqueCount [llength $sharedStrings]

      foreach string $sharedStrings {
	$root appendChild [set node0 [$doc createElement si]]
	  $node0 appendChild [set node1 [$doc createElement t]]
	    $node1 appendChild [$doc createTextNode $string]
      }
    }


    # xl/calcChain.xml
    if $obj(calcChain) {
      set doc [set obj(doc,xl/calcChain.xml) [dom createDocument calcChain]]
      set root [$doc documentElement]
      $root setAttribute xmlns http://schemas.openxmlformats.org/spreadsheetml/2006/main

      $root appendChild [set node0 [$doc createElement c]]
	$node0 setAttribute r C1 i 3 l 1

      $root appendChild [set node0 [$doc createElement c]]
	$node0 setAttribute r A3 i 2
    }


    # xl/styles.xml
    set doc [set obj(doc,xl/styles.xml) [dom createDocument styleSheet]]
    set root [$doc documentElement]
    $root setAttribute xmlns http://schemas.openxmlformats.org/spreadsheetml/2006/main
    $root setAttribute xmlns:mc http://schemas.openxmlformats.org/markup-compatibility/2006
    $root setAttribute xmlns:x14ac http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac
    $root setAttribute mc:Ignorable x14ac

    if {$obj(numFmts) > $::ooxml::defaults(numFmts,start)} {
      $root appendChild [set node0 [$doc createElement numFmts]]
	$node0 setAttribute count [llength [array names numFmts]]
	foreach idx [lsort -integer [array names numFmts]] {
	  $node0 appendChild [set node1 [$doc createElement numFmt]]
	    $node1 setAttribute numFmtId $idx formatCode $numFmts($idx)
	}
    }
    $root appendChild [set node0 [$doc createElement fonts]]
      $node0 setAttribute count $obj(fonts) x14ac:knownFonts 1
      foreach idx [lsort -integer [array names fonts]] {
	$node0 appendChild [set node1 [$doc createElement font]]
	  if {[dict get $fonts($idx) bold] == 1} {
	    $node1 appendChild [set node2 [$doc createElement b]]
	  }
	  if {[dict get $fonts($idx) italic] == 1} {
	    $node1 appendChild [set node2 [$doc createElement i]]
	  }
	  if {[dict get $fonts($idx) underline] == 1} {
	    $node1 appendChild [set node2 [$doc createElement u]]
	  }
	  $node1 appendChild [set node2 [$doc createElement sz]]
	    $node2 setAttribute val [dict get $fonts($idx) size]
	  if {[dict get $fonts($idx) color] ne {}} {
	    $node1 appendChild [set node2 [$doc createElement color]]
	      $node2 setAttribute [lindex [dict get $fonts($idx) color] 0]  [lindex [dict get $fonts($idx) color] 1]
	  }
	  $node1 appendChild [set node2 [$doc createElement name]]
	    $node2 setAttribute val [dict get $fonts($idx) name]
	  $node1 appendChild [set node2 [$doc createElement family]]
	    $node2 setAttribute val [dict get $fonts($idx) family]
	  $node1 appendChild [set node2 [$doc createElement scheme]]
	    $node2 setAttribute val [dict get $fonts($idx) scheme]
      }

    if {$obj(fills) > 0} {
      $root appendChild [set node0 [$doc createElement fills]]
	$node0 setAttribute count $obj(fills)
	foreach idx [lsort -integer [array names fills]] {
	  $node0 appendChild [set node1 [$doc createElement fill]]
	    $node1 appendChild [set node2 [$doc createElement patternFill]]
	      $node2 setAttribute patternType [dict get $fills($idx) patterntype]
	      foreach tag {fgColor bgColor} {
	        set key [string tolower $tag]
		if {[dict get $fills($idx) $key] ne {}} {
		  $node2 appendChild [set node3 [$doc createElement $tag]]
		    $node3 setAttribute [lindex [dict get $fills($idx) $key] 0] [lindex [dict get $fills($idx) $key] 1]
		}
	      }
	}
    }

    if {$obj(borders) > 0} {
      $root appendChild [set node0 [$doc createElement borders]]
	$node0 setAttribute count $obj(borders)
	foreach idx [lsort -integer [array names borders]] {
	  $node0 appendChild [set node1 [$doc createElement border]]
	    if {[dict exists $borders($idx) diagonal direction] && [dict get $borders($idx) diagonal direction] ne {}} {
	      $node1 setAttribute [string map {up diagonalUp down diagonalDown} [dict get $borders($idx) diagonal direction]] 1
	    }
	    foreach item {left right top bottom diagonal} {
	      $node1 appendChild [set node2 [$doc createElement $item]]
	      if {[dict exists $borders($idx) $item style] && [dict get $borders($idx) $item style] ne {}} {
		$node2 setAttribute style [dict get $borders($idx) $item style]
	      }
	      if {[dict exists $borders($idx) $item color] && [dict get $borders($idx) $item color] ne {}} {
		$node2 appendChild [set node3 [$doc createElement color]]
		  $node3 setAttribute [lindex [dict get $borders($idx) $item color] 0] [lindex [dict get $borders($idx) $item color] 1]
	      }
	    }
	}
    }

    $root appendChild [set node0 [$doc createElement cellStyleXfs]]
      $node0 setAttribute count 1
      $node0 appendChild [set node1 [$doc createElement xf]]
	$node1 setAttribute numFmtId 0
	$node1 setAttribute fontId 0
	$node1 setAttribute fillId 0
	$node1 setAttribute borderId 0

    $root appendChild [set node0 [$doc createElement cellXfs]]
      $node0 setAttribute count $obj(styles)
      foreach idx [lsort -integer [array names styles]] {
	$node0 appendChild [set node1 [$doc createElement xf]]
	  $node1 setAttribute numFmtId [dict get $styles($idx) numfmt]
	  $node1 setAttribute fontId [dict get $styles($idx) font]
	  $node1 setAttribute fillId [dict get $styles($idx) fill]
	  $node1 setAttribute borderId [dict get $styles($idx) border]
	  $node1 setAttribute xfId [dict get $styles($idx) xf]
	  if {[dict get $styles($idx) numfmt] > 0} {
	    $node1 setAttribute applyNumberFormat 1
	  }
	  if {[dict get $styles($idx) font] > 0} {
	    $node1 setAttribute applyFont 1
	  }
	  if {[dict get $styles($idx) fill] > 0} {
	    $node1 setAttribute applyFill 1
	  }
	  if {[dict get $styles($idx) border] > 0} {
	    $node1 setAttribute applyBorder 1
	  }
	  if {[dict get $styles($idx) horizontal] ne {} || [dict get $styles($idx) vertical] ne {} || [dict get $styles($idx) rotate] ne {}} {
	    $node1 setAttribute applyAlignment 1
	    $node1 appendChild [set node2 [$doc createElement alignment]]
	      if {[dict get $styles($idx) horizontal] ne {}} {
		$node2 setAttribute horizontal [dict get $styles($idx) horizontal]
	      }
	      if {[dict get $styles($idx) vertical] ne {}} {
		$node2 setAttribute vertical [dict get $styles($idx) vertical]
	      }
	      if {[dict get $styles($idx) rotate] ne {}} {
		$node2 setAttribute textRotation [dict get $styles($idx) rotate]
	      }
	  }
	  # $node1 setAttribute applyProtection 1 quotePrefix 1
      }

    $root appendChild [set node0 [$doc createElement cellStyles]]
      $node0 setAttribute count 1
      $node0 appendChild [set node1 [$doc createElement cellStyle]]
	$node1 setAttribute name Standard
	$node1 setAttribute xfId 0
	$node1 setAttribute builtinId 0

    $root appendChild [set node0 [$doc createElement dxfs]]
      $node0 setAttribute count 0

    $root appendChild [set node0 [$doc createElement tableStyles]]
    $node0 setAttribute count 0


    # xl/theme/theme1.xml
    set doc [set obj(doc,xl/theme/theme1.xml) [dom createDocument a:theme]]
    set root [$doc documentElement]
    $root setAttribute xmlns:a http://schemas.openxmlformats.org/drawingml/2006/main
    $root setAttribute name Office-Design

    $root appendChild [set node0 [$doc createElement a:themeElements]]
      $node0 appendChild [set node1 [$doc createElement a:clrScheme]]
	$node1 setAttribute name Office
	$node1 appendChild [set node2 [$doc createElement a:dk1]]
	  $node2 appendChild [set node3 [$doc createElement a:sysClr]]
	    $node3 setAttribute val windowText lastClr 000000
	$node1 appendChild [set node2 [$doc createElement a:lt1]]
	  $node2 appendChild [set node3 [$doc createElement a:sysClr]]
	    $node3 setAttribute val window lastClr FFFFFF
	$node1 appendChild [set node2 [$doc createElement a:dk2]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 1F497D
	$node1 appendChild [set node2 [$doc createElement a:lt2]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val EEECE1
	$node1 appendChild [set node2 [$doc createElement a:accent1]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 4F81BD
	$node1 appendChild [set node2 [$doc createElement a:accent2]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val C0504D
	$node1 appendChild [set node2 [$doc createElement a:accent3]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 9BBB59
	$node1 appendChild [set node2 [$doc createElement a:accent4]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 8064A2
	$node1 appendChild [set node2 [$doc createElement a:accent5]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 4BACC6
	$node1 appendChild [set node2 [$doc createElement a:accent6]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val F79646
	$node1 appendChild [set node2 [$doc createElement a:hlink]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 0000FF
	$node1 appendChild [set node2 [$doc createElement a:folHlink]]
	  $node2 appendChild [set node3 [$doc createElement a:srgbClr]]
	    $node3 setAttribute val 800080
      $node0 appendChild [set node1 [$doc createElement a:fontScheme]]
	$node1 setAttribute name Office
	$node1 appendChild [set node2 [$doc createElement a:majorFont]]
	  $node2 appendChild [set node3 [$doc createElement a:latin]]
	    $node3 setAttribute typeface Cambria
	  $node2 appendChild [set node3 [$doc createElement a:ea]]
	    $node3 setAttribute typeface {}
	  $node2 appendChild [set node3 [$doc createElement a:cs]]
	    $node3 setAttribute typeface {}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Jpan typeface \uFF2D\uFF33\u0020\uFF30\u30B4\u30B7\u30C3\u30AF
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hang typeface \uB9D1\uC740\u0020\uACE0\uB515
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hans typeface \u5B8B\u4F53
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hant typeface \u65B0\u7D30\u660E\u9AD4
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Arab typeface {Times New Roman}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hebr typeface {Times New Roman}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Thai typeface Tahoma
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Ethi typeface Nyala
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Beng typeface Vrinda
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Gujr typeface Shruti
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Khmr typeface MoolBoran
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Knda typeface Tunga
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Guru typeface Raavi
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Cans typeface Euphemia
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Cher typeface {Plantagenet Cherokee}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Yiii typeface {Microsoft Yi Baiti}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Tibt typeface {Microsoft Himalaya}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Thaa typeface {MV Boli}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Deva typeface Mangal
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Telu typeface Gautami
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Taml typeface Latha
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Syrc typeface {Estrangelo Edessa}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Orya typeface Kalinga
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Mlym typeface Kartika
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Laoo typeface DokChampa
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Sinh typeface {Iskoola Pota}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Mong typeface {Mongolian Baiti}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Viet typeface {Times New Roman}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Uigh typeface {Microsoft Uighur}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Geor typeface Sylfaen
	$node1 appendChild [set node2 [$doc createElement a:minorFont]]
	  $node2 appendChild [set node3 [$doc createElement a:latin]]
	    $node3 setAttribute typeface Calibri
	  $node2 appendChild [set node3 [$doc createElement a:ea]]
	    $node3 setAttribute typeface {}
	  $node2 appendChild [set node3 [$doc createElement a:cs]]
	    $node3 setAttribute typeface {}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Jpan typeface \uFF2D\uFF33\u0020\uFF30\u30B4\u30B7\u30C3\u30AF
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hang typeface \uB9D1\uC740\u0020\uACE0\uB515
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hans typeface \u5B8B\u4F53
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hant typeface \u65B0\u7D30\u660E\u9AD4
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Arab typeface Arial
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Hebr typeface Arial
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Thai typeface Tahoma
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Ethi typeface Nyala
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Beng typeface Vrinda
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Gujr typeface Shruti
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Khmr typeface DaunPenh
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Knda typeface Tunga
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Guru typeface Raavi
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Cans typeface Euphemia
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Cher typeface {Plantagenet Cherokee}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Yiii typeface {Microsoft Yi Baiti}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Tibt typeface {Microsoft Himalaya}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Thaa typeface {MV Boli}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Deva typeface Mangal
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Telu typeface Gautami
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Taml typeface Latha
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Syrc typeface {Estrangelo Edessa}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Orya typeface Kalinga
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Mlym typeface Kartika
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Laoo typeface DokChampa
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Sinh typeface {Iskoola Pota}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Mong typeface {Mongolian Baiti}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Viet typeface Arial
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Uigh typeface {Microsoft Uighur}
	  $node2 appendChild [set node3 [$doc createElement a:font]]
	    $node3 setAttribute script Geor typeface Sylfaen
      $node0 appendChild [set node1 [$doc createElement a:fmtScheme]]
	$node1 setAttribute name Office
	$node1 appendChild [set node2 [$doc createElement a:fillStyleLst]]
	  $node2 appendChild [set node3 [$doc createElement a:solidFill]]
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val phClr
	  $node2 appendChild [set node3 [$doc createElement a:gradFill]]
	    $node3 setAttribute rotWithShape 1
	    $node3 appendChild [set node4 [$doc createElement a:gsLst]]
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 0
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 50000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 300000
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 35000
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 37000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 300000
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 100000
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 15000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 350000
	    $node3 appendChild [set node4 [$doc createElement a:lin]]
	      $node4 setAttribute ang 16200000 scaled 1
	  $node2 appendChild [set node3 [$doc createElement a:gradFill]]
	    $node3 setAttribute rotWithShape 1
	    $node3 appendChild [set node4 [$doc createElement a:gsLst]]
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 0
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 100000
		  $node6 appendChild [set node7 [$doc createElement a:shade]]
		    $node7 setAttribute val 100000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 130000
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 100000
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 50000
		  $node6 appendChild [set node7 [$doc createElement a:shade]]
		    $node7 setAttribute val 100000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 350000
	    $node3 appendChild [set node4 [$doc createElement a:lin]]
	      $node4 setAttribute ang 16200000 scaled 0
	$node1 appendChild [set node2 [$doc createElement a:lnStyleLst]]
	  $node2 appendChild [set node3 [$doc createElement a:ln]]
	    $node3 setAttribute w 9525 cap flat cmpd sng algn ctr
	    $node3 appendChild [set node4 [$doc createElement a:solidFill]]
	      $node4 appendChild [set node5 [$doc createElement a:schemeClr]]
		$node5 setAttribute val phClr
		$node5 appendChild [set node6 [$doc createElement a:shade]]
		  $node6 setAttribute val 95000
		$node5 appendChild [set node6 [$doc createElement a:satMod]]
		  $node6 setAttribute val 105000
	    $node3 appendChild [set node4 [$doc createElement a:prstDash]]
	      $node4 setAttribute val solid
	  $node2 appendChild [set node3 [$doc createElement a:ln]]
	    $node3 setAttribute w 25400 cap flat cmpd sng algn ctr
	    $node3 appendChild [set node4 [$doc createElement a:solidFill]]
	      $node4 appendChild [set node5 [$doc createElement a:schemeClr]]
		$node5 setAttribute val phClr
	    $node3 appendChild [set node4 [$doc createElement a:prstDash]]
		  $node4 setAttribute val solid
	  $node2 appendChild [set node3 [$doc createElement a:ln]]
	    $node3 setAttribute w 38100 cap flat cmpd sng algn ctr
	    $node3 appendChild [set node4 [$doc createElement a:solidFill]]
	      $node4 appendChild [set node5 [$doc createElement a:schemeClr]]
		$node5 setAttribute val phClr
	    $node3 appendChild [set node4 [$doc createElement a:prstDash]]
	      $node4 setAttribute val solid
	$node1 appendChild [set node2 [$doc createElement a:effectStyleLst]]
	  $node2 appendChild [set node3 [$doc createElement a:effectStyle]]
	    $node3 appendChild [set node4 [$doc createElement a:effectLst]]
	      $node4 appendChild [set node5 [$doc createElement a:outerShdw]]
		$node5 setAttribute blurRad 40000 dist 20000 dir 5400000 rotWithShape 0
		$node5 appendChild [set node6 [$doc createElement a:srgbClr]]
		  $node6 setAttribute val 000000
		  $node6 appendChild [set node7 [$doc createElement a:alpha]]
		    $node7 setAttribute val 38000
	  $node2 appendChild [set node3 [$doc createElement a:effectStyle]]
	    $node3 appendChild [set node4 [$doc createElement a:effectLst]]
	      $node4 appendChild [set node5 [$doc createElement a:outerShdw]]
		$node5 setAttribute blurRad 40000 dist 23000 dir 5400000 rotWithShape 0
		$node5 appendChild [set node6 [$doc createElement a:srgbClr]]
		  $node6 setAttribute val 000000
		  $node6 appendChild [set node7 [$doc createElement a:alpha]]
		    $node7 setAttribute val 35000
	  $node2 appendChild [set node3 [$doc createElement a:effectStyle]]
	    $node3 appendChild [set node4 [$doc createElement a:effectLst]]
	      $node4 appendChild [set node5 [$doc createElement a:outerShdw]]
		$node5 setAttribute blurRad 40000 dist 23000 dir 5400000 rotWithShape 0
		$node5 appendChild [set node6 [$doc createElement a:srgbClr]]
		  $node6 setAttribute val 000000
		  $node6 appendChild [set node7 [$doc createElement a:alpha]]
		    $node7 setAttribute val 35000
	    $node3 appendChild [set node4 [$doc createElement a:scene3d]]
	      $node4 appendChild [set node5 [$doc createElement a:camera]]
		$node5 setAttribute prst orthographicFront
		$node5 appendChild [set node6 [$doc createElement a:rot]]
		  $node6 setAttribute lat 0 lon 0 rev 0
	      $node4 appendChild [set node5 [$doc createElement a:lightRig]]
		$node5 setAttribute rig threePt dir t
		$node5 appendChild [set node6 [$doc createElement a:rot]]
		  $node6 setAttribute lat 0 lon 0 rev 1200000
	    $node3 appendChild [set node4 [$doc createElement a:sp3d]]
	      $node4 appendChild [set node5 [$doc createElement a:bevelT]]
		$node5 setAttribute w 63500 h 25400
	$node1 appendChild [set node2 [$doc createElement a:bgFillStyleLst]]
	  $node2 appendChild [set node3 [$doc createElement a:solidFill]]
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val phClr
	  $node2 appendChild [set node3 [$doc createElement a:gradFill]]
	    $node3 setAttribute rotWithShape 1
	    $node3 appendChild [set node4 [$doc createElement a:gsLst]]
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 0
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 40000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 350000
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 40000
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 45000
		  $node6 appendChild [set node7 [$doc createElement a:shade]]
		    $node7 setAttribute val 99000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 350000
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 100000
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:shade]]
		    $node7 setAttribute val 20000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 255000
	    $node3 appendChild [set node4 [$doc createElement a:path]]
	      $node4 setAttribute path circle
	      $node4 appendChild [set node5 [$doc createElement a:fillToRect]]
		$node5 setAttribute l 50000 t -80000 r 50000 b 180000
	  $node2 appendChild [set node3 [$doc createElement a:gradFill]]
	    $node3 setAttribute rotWithShape 1
	    $node3 appendChild [set node4 [$doc createElement a:gsLst]]
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 0
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:tint]]
		    $node7 setAttribute val 80000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 300000
	      $node4 appendChild [set node5 [$doc createElement a:gs]]
		$node5 setAttribute pos 100000
		$node5 appendChild [set node6 [$doc createElement a:schemeClr]]
		  $node6 setAttribute val phClr
		  $node6 appendChild [set node7 [$doc createElement a:shade]]
		    $node7 setAttribute val 30000
		  $node6 appendChild [set node7 [$doc createElement a:satMod]]
		    $node7 setAttribute val 200000
	    $node3 appendChild [set node4 [$doc createElement a:path]]
	      $node4 setAttribute path circle
	      $node4 appendChild [set node5 [$doc createElement a:fillToRect]]
		$node5 setAttribute l 50000 t 50000 r 50000 b 50000

    $root appendChild [set node0 [$doc createElement a:objectDefaults]]
      $node0 appendChild [set node1 [$doc createElement a:spDef]]
	$node1 appendChild [set node2 [$doc createElement a:spPr]]
	$node1 appendChild [set node2 [$doc createElement a:bodyPr]]
	$node1 appendChild [set node2 [$doc createElement a:lstStyle]]
	$node1 appendChild [set node2 [$doc createElement a:style]]
	  $node2 appendChild [set node3 [$doc createElement a:lnRef]]
	    $node3 setAttribute idx 1
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val accent1
	  $node2 appendChild [set node3 [$doc createElement a:fillRef]]
	    $node3 setAttribute idx 3
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val accent1
	  $node2 appendChild [set node3 [$doc createElement a:effectRef]]
	    $node3 setAttribute idx 2
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val accent1
	  $node2 appendChild [set node3 [$doc createElement a:fontRef]]
	    $node3 setAttribute idx minor
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val lt1
      $node0 appendChild [set node1 [$doc createElement a:lnDef]]
	$node1 appendChild [set node2 [$doc createElement a:spPr]]
	$node1 appendChild [set node2 [$doc createElement a:bodyPr]]
	$node1 appendChild [set node2 [$doc createElement a:lstStyle]]
	$node1 appendChild [set node2 [$doc createElement a:style]]
	  $node2 appendChild [set node3 [$doc createElement a:lnRef]]
	    $node3 setAttribute idx 2
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	    $node4 setAttribute val accent1
	  $node2 appendChild [set node3 [$doc createElement a:fillRef]]
	    $node3 setAttribute idx 0
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val accent1
	  $node2 appendChild [set node3 [$doc createElement a:effectRef]]
	    $node3 setAttribute idx 1
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val accent1
	  $node2 appendChild [set node3 [$doc createElement a:fontRef]]
	    $node3 setAttribute idx minor
	    $node3 appendChild [set node4 [$doc createElement a:schemeClr]]
	      $node4 setAttribute val tx1

    $root appendChild [set node0 [$doc createElement a:extraClrSchemeLst]]


    # xl/workbook.xml
    set doc [set obj(doc,xl/workbook.xml) [dom createDocument workbook]]
    set root [$doc documentElement]
    $root setAttribute xmlns http://schemas.openxmlformats.org/spreadsheetml/2006/main
    $root setAttribute xmlns:r http://schemas.openxmlformats.org/officeDocument/2006/relationships

    $root appendChild [set node0 [$doc createElement fileVersion]]
      $node0 setAttribute appName xl lastEdited 5 lowestEdited 5 rupBuild 5000

    $root appendChild [set node0 [$doc createElement workbookPr]]
      $node0 setAttribute showInkAnnotation 0 autoCompressPictures 0

    $root appendChild [set node0 [$doc createElement bookViews]]
      $node0 appendChild [set node1 [$doc createElement workbookView]]
	$node1 setAttribute activeTab 1

    $root appendChild [set node0 [$doc createElement sheets]]
      for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
	$node0 appendChild [set node1 [$doc createElement sheet]]
	  $node1 setAttribute name $obj(sheet,$ws)
	  $node1 setAttribute sheetId $ws
	  $node1 setAttribute r:id rId$ws
      }

    if 0 {
    $root appendChild [set node0 [$doc createElement definedNames]]
      $node0 appendChild [set node1 [$doc createElement definedName]]
	$node1 setAttribute name _xlnm._FilterDatabase localSheetId 0 hidden 1
	$node1 appendChild [$doc createTextNode {Blatt1!$A$1:$C$1}]
    }

    $root appendChild [set node0 [$doc createElement calcPr]]
      $node0 setAttribute calcId 140000 concurrentCalc 0
      # fullCalcOnLoad 1


    # xl/worksheets/sheet1.xml SHEET
    for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
      set doc [set obj(doc,xl/worksheets/sheet$ws.xml) [dom createDocument worksheet]]
      set root [$doc documentElement]
      $root setAttribute xmlns http://schemas.openxmlformats.org/spreadsheetml/2006/main
      $root setAttribute xmlns:r http://schemas.openxmlformats.org/officeDocument/2006/relationships
      $root setAttribute xmlns:mc http://schemas.openxmlformats.org/markup-compatibility/2006
      $root setAttribute xmlns:x14ac http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac
      $root setAttribute mc:Ignorable x14ac

      $root appendChild [set node0 [$doc createElement dimension]]
 	$node0 setAttribute ref [::ooxml::RowColumnToString $obj(dminrow,$ws),$obj(dmincol,$ws)]:[::ooxml::RowColumnToString $obj(dmaxrow,$ws),$obj(dmaxcol,$ws)]

      $root appendChild [set node0 [$doc createElement sheetViews]]
	$node0 appendChild [set node1 [$doc createElement sheetView]]
	  $node1 setAttribute workbookViewId 0
	  if {$obj(freeze,$ws) ne {}} {
	    lassign [split [::ooxml::StringToRowColumn $obj(freeze,$ws)] ,] row col
	    $node1 appendChild [set node2 [$doc createElement pane]]
	      $node2 setAttribute xSplit $col ySplit $row topLeftCell $obj(freeze,$ws) state frozen
	  }

      $root appendChild [set node0 [$doc createElement sheetFormatPr]]
	$node0 setAttribute baseColWidth 10 defaultRowHeight 16 x14ac:dyDescent 0.2

      if {[info exists obj($ws,cols)] && $obj($ws,cols) > 0} {
	$root appendChild [set node0 [$doc createElement cols]]
	  set colsNode $node0
      }

      $root appendChild [set node0 [$doc createElement sheetData]]

      set lastRow -1
      foreach idx [lsort -dictionary [array names cells $ws,*,*]] {
	lassign [split $idx ,] sheet row col
	set maxCol $col
	if {$row != $lastRow} {
	  set lastRow $row
	  set minCol $col
	  $node0 appendChild [set node1 [$doc createElement row]]
	    $node1 setAttribute r [expr {$row + 1}]
	    if {[dict exists $obj(rowHeight,$ws) $row]} {
	      $node1 setAttribute ht [dict get $obj(rowHeight,$ws) $row] customHeight 1
	    }
        }
	if {([dict exists $cells($idx) v] && [string trim [dict get $cells($idx) v]] ne {}) || ([dict exists $cells($idx) f] && [string trim [dict get $cells($idx) f]] ne {})} {
	  #$node1 setAttribute spans [expr {$minCol + 1}]:[expr {$maxCol + 1}]
	  $node1 appendChild [set node2 [$doc createElement c]]
	    $node2 setAttribute r [::ooxml::RowColumnToString $row,$col]
	    if {[dict exists $cells($idx) v] && [dict get $cells($idx) v] ne {}} {
	      if {[dict exists $cells($idx) s] && [dict get $cells($idx) s] > 0} {
		$node2 setAttribute s [dict get $cells($idx) s]
	      }
	      if {[dict exists $cells($idx) t] && [dict get $cells($idx) t] ne {n}} {
		$node2 setAttribute t [dict get $cells($idx) t]
	      }
	      $node2 appendChild [set node3 [$doc createElement v]]
		$node3 appendChild [$doc createTextNode [dict get $cells($idx) v]]
	    }
	    if {[dict exists $cells($idx) f] && [dict get $cells($idx) f] ne {}} {
	      $node2 appendChild [set node3 [$doc createElement f]]
		$node3 appendChild [$doc createTextNode [dict get $cells($idx) f]]
	    }
	} elseif {[dict exists $cells($idx) s] && [string is integer -strict [dict get $cells($idx) s]] && [dict get $cells($idx) s] > 0} {
	  $node1 appendChild [set node2 [$doc createElement c]]
	    $node2 setAttribute r [::ooxml::RowColumnToString $row,$col]
	    $node2 setAttribute s [dict get $cells($idx) s]
	}
      }

      if {$obj(autofilter,$ws) ne {}} {
	$root appendChild [set node0 [$doc createElement autoFilter]]
	  $node0 setAttribute ref $obj(autofilter,$ws)
      }

      if {[info exists obj(merge,$ws)] && $obj(merge,$ws) ne {}} {
	$root appendChild [set node0 [$doc createElement mergeCells]]
	  $node0 setAttribute count [llength $obj(merge,$ws)]
	  foreach item $obj(merge,$ws) {
	    $node0 appendChild [set node1 [$doc createElement mergeCell]]
	      $node1 setAttribute ref $item
	  }
      }

      $root appendChild [set node0 [$doc createElement pageMargins]]
	$node0 setAttribute left 0.75 right 0.75 top 1 bottom 1 header 0.5 footer 0.5

      if {[info exists obj($ws,cols)] && $obj($ws,cols) > 0} {
	set node0 $colsNode
	  foreach idx [lsort -dictionary [array names cols $ws,*]] {
	    $node0 appendChild [set node1 [$doc createElement col]]
	      $node1 setAttribute min [expr {[dict get $cols($idx) min] + 1}] max [expr {[dict get $cols($idx) max] + 1}]
	      if {[dict get $cols($idx) width] ne {}} {
		$node1 setAttribute width [dict get $cols($idx) width]
		if {[dict get $cols($idx) width] != $::ooxml::defaults(cols,width)} {
		  dict set $cols($idx) customwidth 1
		}
	      }
	      if {[dict get $cols($idx) style] ne {} && [dict get $cols($idx) style] > 0} {
		$node1 setAttribute style [dict get $cols($idx) style]
	      }
	      if {[dict get $cols($idx) bestfit] == 1} {
		$node1 setAttribute bestFit [dict get $cols($idx) bestfit]
	      }
	      if {[dict get $cols($idx) customwidth] == 1} {
		$node1 setAttribute customWidth [dict get $cols($idx) customwidth]
	      }
	  }
      }
    }

    # Content-Type application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
    set file [string trim $file]
    if {$file eq {}} {
      set file {spreadsheetml.xlsx}
    }
    if {[file extension $file] ne {.xlsx}} {
      append file {.xlsx}
    }
    set path [file dirname $file]
    set uid [format xl_%X [clock microseconds]]
    set filesToZip {}
    foreach {tag doc} [array get obj doc,*] {
      lappend filesToZip [set docname [lindex [split $tag ,] 1]]
      set xmlfile [file join $path $uid $docname]
      file mkdir [file dirname $xmlfile]
      if {![catch {open $xmlfile w} fd]} {
	fconfigure $fd -encoding utf-8
	#puts $fd "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	puts $fd [[$doc documentElement] asXML -indent $obj(indent) -xmlDeclaration 1 -encString [string toupper $obj(encoding)]]
	close $fd
	$doc delete
      }
    }
    set pwd [pwd]
    cd [file join $path $uid]
    if {$path eq {.}} {
      set file [file join .. $file]
    }
    ::ooxml::Zip $file . $filesToZip
    cd $pwd
    if {!$opts(holdcontainerdirectory)} {
      file delete -force [file join $path $uid]
    }
    return 0
  }
}


#
# ooxml::tablelist_to_xl
#

proc ::ooxml::tablelist_to_xl { lb args } {
  variable defaults

  if {![winfo exists $lb]} {
    tk_messageBox -message {Die Tablelist existiert nicht!}
    return
  }

  if {[::ooxml::Getopt opts [list callback.arg {::ooxml::tablelist_to_xl_callback} path.arg $defaults(path) file.arg {tablelist.xlsx} creator.arg {unknown} name.arg {Tablelist1} rootonly addtimestamp] $args]} {
    error $opts(-errmsg)
  }
  if {[string trim $opts(path)] eq {}} {
    set opts(path) {.}
  }
  if {[string trim $opts(file)] eq {}} {
    set opts(file) {tablelist.xlsx}
  }
  if {[file extension $opts(file)] eq {.xlsx}} {
    set opts(file) [file tail [file rootname $opts(file)]]
  }
  if {$opts(addtimestamp)} {
    append opts(file) _[clock format [clock seconds] -format %Y%m%dT%H%M%S]
  }
  append opts(file) {.xlsx}

  set file [tk_getSaveFile -confirmoverwrite 1 -filetypes {{{Excel Office Open XML} {.xlsx}}} -initialdir $opts(path) -initialfile $opts(file) -parent . -title "Excel Office Open XML"]
  if {$file eq {}} {
    tk_messageBox -message {Keine Datei ausgewählt!}
    return
  }

  set spreadsheet [::ooxml::xl_write new -creator $opts(creator)]
  if {[set sheet [$spreadsheet worksheet $opts(name)]] > -1} {
    set columncount [expr {[$lb columncount] - 1}]
    if {$columncount > 0} {
      $spreadsheet autofilter $sheet 0,0 0,$columncount
    }
    set titlecolumns [$lb cget -titlecolumns]
    if {$titlecolumns > 0} {
      $spreadsheet freeze $sheet 1,$titlecolumns
    }

    set col -1
    set title SETUP
    set width 0
    set align {}
    set sortmode {}
    set hide 1
    $opts(callback) $spreadsheet $sheet $columncount $col $title $width $align $sortmode $hide

    $spreadsheet row $sheet
    for {set col 0} {$col < $columncount} {incr col} {
      set title [$lb columncget $col -title]
      set width [$lb columncget $col -width]
      set align [$lb columncget $col -align]
      set sortmode [$lb columncget $col -sortmode]
      set hide [$lb columncget $col -hide]

      if {[info commands $opts(callback)] eq $opts(callback)} {
        $opts(callback) $spreadsheet $sheet $columncount $col $title $width $align $sortmode $hide
      }

      $spreadsheet cell $sheet $title
    }

    if {$opts(rootonly)} {
      foreach row [$lb get [$lb childkeys root]] {
	$spreadsheet row $sheet
	set idx 0
	foreach col $row {
	  if {[string trim $col] ne {}} {
	    $spreadsheet cell $sheet $col -index $idx
	  }
	  incr idx
	}
      }
    } else {
      foreach row [$lb get 0 end] {
	$spreadsheet row $sheet
	set idx 0
	foreach col $row {
	  if {[string trim $col] ne {}} {
	    $spreadsheet cell $sheet $col -index $idx
	  }
	  incr idx
	}
      }
    }

    $spreadsheet write $file
  }
}

proc ::ooxml::tablelist_to_xl_callback { spreadsheet sheet maxcol column title width align sortmode hide } {
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

package provide ooxml 1.0
