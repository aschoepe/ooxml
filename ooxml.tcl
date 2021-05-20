#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2018-2021 Alexander Schoepe, Bochum, DE, <alx.tcl@sowaswie.de>
#  Copyright (C) 2019 Rolf Ade, DE
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
#     -created UTC-TIMESTAMP
#     -modifiedby NAME
#     -modified UTC-TIMESTAMP
#     -application NAME
#     return class
# 
#   method numberformat args
#     -format FORMAT -general -date -time -datetime -iso8601 -number -decimal -red -separator -fraction -scientific -percent -text -string
#     -tag NAME
#     return NUMFMTID
#
#   method defaultdatestyle STYLEID
# 
#   method font args
#     -list -name NAME -family FAMILY -size SIZE -color COLOR -scheme SCHEME -bold -italic -underline -color COLOR -tag NAME
#     return FONTID
# 
#   method fill args
#     -list -patterntype PATTERNTYPE -fgcolor COLOR -bgcolor COLOR -tag NAME
#     return FILLID
# 
#   method border args
#     -list -leftstyle BORDERLINESTYLE -leftcolor COLOR -rightstyle BORDERLINESTYLE -rightcolor COLOR -topstyle BORDERLINESTYLE -topcolor COLOR
#     -bottomstyle BORDERLINESTYLE -bottomcolor COLOR -diagonalstyle BORDERLINESTYLE -diagonalcolor COLOR -diagonaldirection DIAGONALDIRECTION
#     -tag NAME
#     return BORDERID
# 
#   method style args
#     -list -numfmt NUMFMTID -font FONTID -fill FILLID -border BORDERID -xf XFID -horizontal HORIZONTAL -vertical VERTICAL -rotate DEGREE -wrap
#     -tag NAME
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
#     -index INDEX -style STYLEID -formula FORMULA -formulaidx SHARE -formularef INDEX:INDEX -string -nozero -height HEIGHT
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
#   method view args
#     -avtivetab TAB -x TWIPS -y TWIPS -height TWIPS -width TWIPS -list
#
#   method write filename
# 
#
# ::ooxml::tablelist_to_xl lb args
#   -callback CALLBACK -path PATH -file FILENAME -creator CREATOR -name NAME -rootonly -addtimestamp
#   Callback arguments
#     spreadsheet sheet maxcol column title width align sortmode hide
#


package require Tcl 8.6.7-
package require tdom 0.9.0-
package require msgcat


namespace eval ::ooxml {
  namespace export xl_sheets xl_read xl_write

  variable defaults
  variable initNodeCmds
  variable predefNumFmts
  variable predefColors
  variable predefColorsName
  variable predefColorsARBG
  variable predefBorderLineStyles
  variable predefPatternType
  variable xmlns

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

  array set xmlns {
    M http://schemas.openxmlformats.org/spreadsheetml/2006/main
    CT http://schemas.openxmlformats.org/package/2006/content-types
    EP http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
    PR http://schemas.openxmlformats.org/package/2006/relationships
    a http://schemas.openxmlformats.org/drawingml/2006/main
    cp http://schemas.openxmlformats.org/package/2006/metadata/core-properties
    dc http://purl.org/dc/elements/1.1/
    dcmitype http://purl.org/dc/dcmitype/
    dcterms http://purl.org/dc/terms/
    mc http://schemas.openxmlformats.org/markup-compatibility/2006
    r http://schemas.openxmlformats.org/officeDocument/2006/relationships
    vt http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes
    x14ac http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac
    x16r2 http://schemas.microsoft.com/office/spreadsheetml/2015/02/main
    xsi http://www.w3.org/2001/XMLSchema-instance
  }

  # ar - Arabic - العربية 
  msgcat::mcset ar LANGUAGE \u0627\u0644\u0639\u0631\u0628\u064a\u0629
  msgcat::mcset ar Book \u0627\u0644\u0643\u062a\u0627\u0628
  msgcat::mcset ar Worksheets "\u0623\u0648\u0631\u0627\u0642 \u0627\u0644\u0639\u0645\u0644"
  msgcat::mcset ar Sheet \u0627\u0644\u0648\u0631\u0642\u0629
  # cs - Czech - čeština, český jazyk
  msgcat::mcset cs LANGUAGE \u010de\u0161tina
  msgcat::mcset cs Book Ses\u030cit
  msgcat::mcset cs Worksheets Listy
  msgcat::mcset cs Sheet List
  # da - Danish - dansk
  msgcat::mcset da LANGUAGE dansk
  msgcat::mcset da Book Mappe
  msgcat::mcset da Worksheets Regneark
  msgcat::mcset da Sheet Ark
  # de - German - Deutsch
  msgcat::mcset de LANGUAGE Deutsch
  msgcat::mcset de Book Mappe
  msgcat::mcset de Worksheets Arbeitsbl\u00e4tter
  msgcat::mcset de Sheet Blatt
  msgcat::mcset de {Tablelist does not exists!} {Die Tablelist existiert nicht!}
  msgcat::mcset de {No file selected!} "Keine Datei ausgew\u00e4hlt!"
  # el - Greek - ελληνικά
  msgcat::mcset el LANGUAGE \u03b5\u03bb\u03bb\u03b7\u03bd\u03b9\u03ba\u03ac
  msgcat::mcset el Book \u0392\u03b9\u03b2\u03bb\u03b9\u0301\u03bf
  msgcat::mcset el Worksheets "\u03a6\u03cd\u03bb\u03bb\u03b1 \u03b5\u03c1\u03b3\u03b1\u03c3\u03af\u03b1\u03c2"
  msgcat::mcset el Sheet \u03a6\u03cd\u03bb\u03bb\u03bf
  # en - English - English
  msgcat::mcset en LANGUAGE English
  msgcat::mcset en Book Book
  msgcat::mcset en Worksheets Worksheets
  msgcat::mcset en Sheet Sheet
  msgcat::mcset en {Tablelist does not exists!} {Tablelist does not exists!}
  msgcat::mcset en {No file selected!} {No file selected!}
  # es - Spanish - Español
  msgcat::mcset es LANGUAGE Espa\u00f1ol
  msgcat::mcset es Book Libro
  msgcat::mcset es Worksheets "Hojas de c\u00e1lculo"
  msgcat::mcset es Sheet Hoja
  msgcat::mcset es {Tablelist does not exists!} "\u00a1Tablelist no existe!"
  msgcat::mcset es {No file selected!} "\u00a1Ning\u00fan archivo seleccionado!"
  # fi - Finnish - suomi, suomen kieli
  msgcat::mcset fi LANGUAGE suomi
  msgcat::mcset fi Book Tyo\u0308kirja
  msgcat::mcset fi Worksheets Laskentataulukot
  msgcat::mcset fi Sheet Taulukko
  # fr - French - français, langue française
  msgcat::mcset fr LANGUAGE fran\u00e7ais
  msgcat::mcset fr Book Classeur
  msgcat::mcset fr Worksheets "Feuilles de calcul"
  msgcat::mcset fr Sheet Feuil
  msgcat::mcset fr {Tablelist does not exists!} {Tablelist n'existe pas!}
  msgcat::mcset fr {No file selected!} "Aucun fichier s\u00e9lectionn\u00e9!"
  # he - Hebrew - עברית
  msgcat::mcset he LANGUAGE \u05e2\u05d1\u05e8\u05d9\u05ea
  msgcat::mcset he Book \u05d7\u05d5\u05d1\u05e8\u05ea
  msgcat::mcset he Worksheets "\u05d2\u05dc\u05d9\u05d5\u05e0\u05d5\u05ea \u05e2\u05d1\u05d5\u05d3\u05d4"
  msgcat::mcset he Sheet \u05d2\u05d9\u05dc\u05d9\u05d5\u05df
  # hu - Hungarian - magyar
  msgcat::mcset hu LANGUAGE magyar
  msgcat::mcset hu Book Munkafu\u0308zet
  msgcat::mcset hu Worksheets Munkalapok
  msgcat::mcset hu Sheet Munkalap
  # it - italian - Italiano
  msgcat::mcset it LANGUAGE Italiano
  msgcat::mcset it Book Cartel
  msgcat::mcset it Worksheets "Fogli di lavoro"
  msgcat::mcset it Sheet Foglio
  msgcat::mcset it {Tablelist does not exists!} {Tablelist non esiste!}
  msgcat::mcset it {No file selected!} {Nessun file selezionato!}
  # ja - Japanese - 日本語 (にほんご)
  msgcat::mcset ja LANGUAGE "\u65e5\u672c\u8a9e (\u306b\u307b\u3093\u3054)"
  msgcat::mcset ja Book Book
  msgcat::mcset ja Worksheets \u30ef\u30fc\u30af\u30b7\u30fc\u30c8
  msgcat::mcset ja Sheet Sheet
  # ko - Korean - 한국어
  msgcat::mcset ko LANGUAGE "\ud55c\uad6d\uc5b4"
  msgcat::mcset ko Book "\u1110\u1169\u11bc\u1112\u1161\u11b8 \u1106\u116e\u11ab\u1109\u1165"
  msgcat::mcset ko Worksheets \uc6cc\ud06c\uc2dc\ud2b8
  msgcat::mcset ko Sheet \uc2dc\ud2b8
  # nl - Dutch, Flemish - Nederlands, Vlaams
  msgcat::mcset nl LANGUAGE Nederlands
  msgcat::mcset nl Book Map
  msgcat::mcset nl Worksheets Werkbladen
  msgcat::mcset nl Sheet Blad
  msgcat::mcset nl {Tablelist does not exists!} {Tablelist bestaat niet!}
  msgcat::mcset nl {No file selected!} {Geen bestand geselecteerd!}
  # no - Norwegian - Norsk
  msgcat::mcset no LANGUAGE Norsk
  msgcat::mcset no Book Bok
  msgcat::mcset no Worksheets Regneark
  msgcat::mcset no Sheet Ark
  # pl - Polish - język polski, polszczyzna
  msgcat::mcset pl LANGUAGE polszczyzna
  msgcat::mcset pl Book Skoroszyt
  msgcat::mcset pl Worksheets Arkusze
  msgcat::mcset pl Sheet Arkusz
  # pt - Portuguese - Português
  msgcat::mcset pt LANGUAGE Portugu\u00eas
  msgcat::mcset pt Book Livro
  msgcat::mcset pt Worksheets "Folhas de C\u00e1lculo"
  msgcat::mcset pt Sheet Folha
  # ru - Russian - русский
  msgcat::mcset ru LANGUAGE \u0440\u0443\u0441\u0441\u043a\u0438\u0439
  msgcat::mcset ru Book \u041a\u043d\u0438\u0433\u0430
  msgcat::mcset ru Worksheets \u041b\u0438\u0441\u0442\u044b
  msgcat::mcset ru Sheet \u041b\u0438\u0441\u0442
  # sl - Slovenian - Slovenski Jezik, Slovenščina
  msgcat::mcset sl LANGUAGE Sloven\u0161\u010dina
  msgcat::mcset sl Book Zos\u030cit
  msgcat::mcset sl Worksheets H\u00e1rky
  msgcat::mcset sl Sheet H\u00e1rok
  # sv - Swedish - Svenska
  msgcat::mcset sv LANGUAGE Svenska
  msgcat::mcset sv Book Bok
  msgcat::mcset sv Worksheets Kalkylblad
  msgcat::mcset sv Sheet Blad
  # th - Thai - ไทย
  msgcat::mcset th LANGUAGE \u0e44\u0e17\u0e22
  msgcat::mcset th Book \u0e2a\u0e21\u0e38\u0e14\u0e07\u0e32\u0e19
  msgcat::mcset th Worksheets \u0e40\u0e27\u0e34\u0e23\u0e4c\u0e01\u0e0a\u0e35\u0e15
  msgcat::mcset th Sheet \u0e41\u0e1c\u0e48\u0e19\u0e07\u0e32\u0e19
  # tr - Turkish - Türkçe
  msgcat::mcset tr LANGUAGE T\u00fcrk\u00e7e
  msgcat::mcset tr Book Kitap
  msgcat::mcset tr Worksheets "\u00c7al\u0131\u015fma Sayfalar\u0131"
  msgcat::mcset tr Sheet Sayfa
  # zh - Chinese - 中文 (Zhōngwén), 汉语, 漢語
  msgcat::mcset zh LANGUAGE \u4e2d\u6587
  msgcat::mcset zh Book \u5de5\u4f5c\u7c3f
  msgcat::mcset zh Worksheets \u5de5\u4f5c\u8868
  msgcat::mcset zh Sheet \u5de5\u4f5c\u8868
}

# ooxml::timet_to_dos
#
#        Convert a unix timestamp into a DOS timestamp for ZIP times.
#
#   DOS timestamps are 32 bits split into bit regions as follows:
#                  24                16                 8                 0
#   +-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+
#   |Y|Y|Y|Y|Y|Y|Y|m| |m|m|m|d|d|d|d|d| |h|h|h|h|h|m|m|m| |m|m|m|s|s|s|s|s|
#   +-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+
#
# From tcllib / zipfile::mkzip      

proc ::ooxml::timet_to_dos {time_t} {
  set s [clock format $time_t -format {%Y %m %e %k %M %S}]
  scan $s {%d %d %d %d %d %d} year month day hour min sec
  expr {(($year-1980) << 25) | ($month << 21) | ($day << 16) 
        | ($hour << 11) | ($min << 5) | ($sec >> 1)}
}


# ooxml::add_str_to_archive --
#
#        Add a string as a single file with string as content with
#        argument path to a zip archive. The zipchan channel must
#        already be open and binary. The return value is the central
#        directory record that will need to be used when finalizing
#        the zip archive.
#
# Derived from tcllib / zipfile::mkzip::add_file_to_archive

proc ::ooxml::add_str_to_archive {zipchan path data {comment {}}} {
  set mtime [timet_to_dos [clock seconds]]
  set utfpath [encoding convertto utf-8 $path]
  set utfcomment [encoding convertto utf-8 $comment]
  set flags [expr {(1<<11)}] ;# utf-8 comment and path
  set method 0               ;# store 0, deflate 8
  set attr 0                 ;# text or binary (default binary)
  set version 20             ;# minumum version req'd to extract
  set extra {}
  set crc 0
  set size 0
  set csize 0
  set seekable [expr {[tell $zipchan] != -1}]
  set attrex 0x81b60020  ;# 0o100666 (-rw-rw-rw-)
  
  set utfdata [encoding convertto utf-8 $data]
  set size [string length $utfdata]
  
  set offset [tell $zipchan]
  set local [binary format a4sssiiiiss PK\03\04 $version $flags $method $mtime $crc $csize $size [string length $utfpath] [string length $extra]]
  append local $utfpath $extra
  puts -nonewline $zipchan $local
  
  set crc [::zlib crc32 $utfdata]
  set cdata [::zlib deflate $utfdata]
  if {[string length $cdata] < $size} {
    set method 8
    set utfdata $cdata
  }
  set csize [string length $utfdata]
  puts -nonewline $zipchan $utfdata

  # update the header
  set local [binary format a4sssiiii PK\03\04 $version $flags $method $mtime $crc $csize $size]
  set current [tell $zipchan]
  seek $zipchan $offset
  puts -nonewline $zipchan $local
  seek $zipchan $current
  
  set hdr [binary format a4ssssiiiisssssii PK\01\02 0x0317 $version $flags $method $mtime $crc $csize $size [string length $utfpath] [string length $extra] [string length $utfcomment] 0 $attr $attrex $offset]
  append hdr $utfpath $extra $utfcomment
  return $hdr
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

proc ::ooxml::Column { col } {
  set name {}
  while {$col >= 0} {
    set char [binary format c [expr {($col % 26) + 65}]]
    set name $char$name
    set col [expr {$col / 26 -1}]
  }
  return $name
}


proc ::ooxml::RowColumnToString { rowcol } {
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
  variable xmlns

  package require vfs::zip

  set sheets {}

  set mnt [vfs::zip::Mount $file xlsx]

  set rels 0
  if {![catch {open xlsx/xl/_rels/workbook.xml.rels r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} rdoc]} {
      set rels 1
      $rdoc documentElement relsRoot
      $rdoc selectNodesNamespaces [list PR $xmlns(PR)]
    }
    close $fd
  }

  if {![catch {open xlsx/xl/workbook.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      $doc documentElement root
      $doc selectNodesNamespaces [list M $xmlns(M) r $xmlns(r)]
      set idx -1
      foreach node [$root selectNodes /M:workbook/M:sheets/M:sheet] {
	if {[$node hasAttribute sheetId] && [$node hasAttribute name]} {
	  set sheetId [$node @sheetId]
	  set name [$node @name]
	  set rid [$node getAttributeNS $xmlns(r) id]
	  foreach node [$relsRoot selectNodes {/PR:Relationships/PR:Relationship[@Id=$rid]}] {
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
  variable xmlns

  variable predefNumFmts

  package require vfs::zip

  array set cellXfs {}
  array set numFmts [array get predefNumFmts]
  array set sharedStrings {}
  set sheets {}

  array set opts {
    valuesonly 0
    keylist 0
    sheets {}
    sheetnames {}
    datefmt {%Y-%m-%d %H:%M:%S}
    as array
  }

  set len [llength $args]
  set idx 0
  for {set idx 0} {$idx < $len} {incr idx} {
    switch -- [set opt [lindex $args $idx]] {
      -sheets - -sheetnames - -datefmt - -as {
	incr idx
	if {$idx < $len} {
	  set opts([string range $opt 1 end]) [lindex $args $idx]
	} else {
	  error "option '$opt': missing argument"
	}            
      }
      -valuesonly - -keylist {
	set opts([string range $opt 1 end]) 1
      }
      default {
	error "unknown option \"$opt\", should be: -sheets, -sheetnames, -datefmt, -as, -valuesonly or -keylist"
      }
    }
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
      $rdoc documentElement relsRoot
      $rdoc selectNodesNamespaces [list PR $xmlns(PR)]
    }
    close $fd
  }

  if {![catch {open xlsx/xl/workbook.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      $doc documentElement root
      $doc selectNodesNamespaces [list M $xmlns(M) r $xmlns(r)]
      set idx -1
      foreach node [$root selectNodes /M:workbook/M:sheets/M:sheet] {
	if {[$node hasAttribute sheetId] && [$node hasAttribute name]} {
	  set sheetId [$node @sheetId]
	  set name [$node @name]
	  set rid [$node getAttributeNS $xmlns(r) id]
	  foreach node [$relsRoot selectNodes {/PR:Relationships/PR:Relationship[@Id=$rid]}] {
	    if {[$node hasAttribute Target]} {
	      lappend sheets [incr idx] $sheetId $name $rid [$node @Target]
	    }
	  }
	}
      }
      foreach node [$root selectNodes /M:workbook/M:bookViews/M:workbookView] {
        if {[$node hasAttribute activeTab]} {
	  lappend wb(view) activetab [$node @activeTab]
	}
        if {[$node hasAttribute xWindow]} {
	  lappend wb(view) x [$node @xWindow]
	}
        if {[$node hasAttribute yWindow]} {
	  lappend wb(view) y [$node @yWindow]
	}
        if {[$node hasAttribute windowHeight]} {
	  lappend wb(view) height [$node @windowHeight]
	}
        if {[$node hasAttribute windowWidth]} {
	  lappend wb(view) width [$node @windowWidth]
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
      $doc documentElement root
      $doc selectNodesNamespaces [list M $xmlns(M)]
      set idx -1
      foreach shared [$root selectNodes /M:sst/M:si] {
	incr idx
	foreach node [$shared selectNodes M:t/text()] {
	  append sharedStrings($idx) [$node nodeValue]
	}
	foreach node [$shared selectNodes */M:t/text()] {
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
      $doc documentElement root
      $doc selectNodesNamespaces [list M $xmlns(M) mc $xmlns(mc) x14ac $xmlns(x14ac) x16r2 $xmlns(x16r2)]
      set idx -1
      foreach node [$root selectNodes /M:styleSheet/M:numFmts/M:numFmt] {
        incr idx
	if {[$node hasAttribute numFmtId] && [$node hasAttribute formatCode]} {
	  set numFmtId [$node @numFmtId]
	  set formatCode [$node @formatCode]
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
      foreach node [$root selectNodes /M:styleSheet/M:cellXfs/M:xf] {
        incr idx
	if {[$node hasAttribute numFmtId]} {
	  set numFmtId [$node @numFmtId]
	  if {[$node hasAttribute applyNumberFormat]} {
	    set applyNumberFormat [$node @applyNumberFormat]
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
      foreach node [$root selectNodes /M:styleSheet/M:numFmts/M:numFmt] {
        if {[$node hasAttribute numFmtId] && [$node hasAttribute formatCode]} {
	  set wb(s,numFmts,[set idx [$node @numFmtId]]) [$node @formatCode]
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
      foreach node [$root selectNodes /M:styleSheet/M:fonts/M:font] {
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
		set a(size) [$node1 @val]
	      }
	    }
	    color {
	      if {[$node1 hasAttribute auto]} {
		set a(color) [list auto [$node1 @auto]]
	      } elseif {[$node1 hasAttribute rgb]} {
		set a(color) [list rgb [$node1 @rgb]]
	      } elseif {[$node1 hasAttribute indexed]} {
		set a(color) [list indexed [$node1 @indexed]]
	      } elseif {[$node1 hasAttribute theme]} {
		set a(color) [list theme [$node1 @theme]]
	      }
	    }
	    name {
	      if {[$node1 hasAttribute val]} {
		set a(name) [$node1 @val]
	      }
	    }
	    family {
	      if {[$node1 hasAttribute val]} {
		set a(family) [$node1 @val]
	      }
	    }
	    scheme {
	      if {[$node1 hasAttribute val]} {
		set a(scheme) [$node1 @val]
	      }
	    }
	  }
	}
	set wb(s,fonts,$idx) [array get a]
      }
      lappend wb(s,@) fonts [incr idx]


      set idx -1
      array unset a *
      foreach node [$root selectNodes /M:styleSheet/M:fills/M:fill] {
	incr idx
	array set a {patterntype {} fgcolor {} bgcolor {}}
	foreach node1 [$node childNodes] {
	  switch -- [$node1 nodeName] {
	    patternFill {
	      if {[$node1 hasAttribute patternType]} {
		set a(patterntype) [$node1 @patternType]
	      }
	      foreach node2 [$node1 childNodes] {
		if {[$node2 nodeName] in { fgColor bgColor}} {
		  if {[$node2 hasAttribute auto]} {
		    set a([string tolower [$node2 nodeName]]) [list auto [$node2 @auto]]
		  } elseif {[$node2 hasAttribute rgb]} {
		    set a([string tolower [$node2 nodeName]]) [list rgb [$node2 @rgb]]
		  } elseif {[$node2 hasAttribute indexed]} {
		    set a([string tolower [$node2 nodeName]]) [list indexed [$node2 @indexed]]
		  } elseif {[$node2 hasAttribute theme]} {
		    set a([string tolower [$node2 nodeName]]) [list theme [$node2 @theme]]
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
      foreach node [$root selectNodes /M:styleSheet/M:borders/M:border] {
	incr idx
	set d {left {style {} color {}} right {style {} color {}} top {style {} color {}} bottom {style {} color {}} diagonal {style {} color {} direction {}}}
	foreach node1 [$node childNodes] {
	  if {[$node1 hasAttribute style]} {
	    set style [$node1 @style]
	  } else {
	    set style {}
	  }
	  set color {}
	  foreach node2 [$node1 childNodes] {
	    if {[$node2 nodeName] eq {color}} {
	      if {[$node2 hasAttribute auto]} {
		set color [list auto [$node2 @auto]]
	      } elseif {[$node2 hasAttribute rgb]} {
		set color [list rgb [$node2 @rgb]]
	      } elseif {[$node2 hasAttribute indexed]} {
		set color [list indexed [$node2 @indexed]]
	      } elseif {[$node2 hasAttribute theme]} {
		set color [list theme [$node2 @theme]]
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
	  dict set d diagonal direction diagonalUp
	} elseif {[$node hasAttribute diagonalDown]} {
	  dict set d diagonal direction diagonalDown
	}
	set wb(s,borders,$idx) $d
      }
      lappend wb(s,@) borders [incr idx]


      set idx -1
      array unset a *
      foreach node [$root selectNodes /M:styleSheet/M:cellXfs/M:xf] {
	incr idx
	array set a {numfmt 0 font 0 fill 0 border 0 xf 0 horizontal {} vertical {} rotate {} wrap {}}
        if {[$node hasAttribute numFmtId]} {
	  set a(numfmt) [$node @numFmtId]
	}
        if {[$node hasAttribute fontId]} {
	  set a(font) [$node @fontId]
	}
        if {[$node hasAttribute fillId]} {
	  set a(fill) [$node @fillId]
	}
        if {[$node hasAttribute borderId]} {
	  set a(border) [$node @borderId]
	}
        if {[$node hasAttribute xfId]} {
	  set a(xf) [$node @xfId]
	}
	foreach node1 [$node childNodes] {
	  switch -- [$node1 nodeName] {
	    alignment {
	      if {[$node1 hasAttribute horizontal]} {
		set a(horizontal) [$node1 @horizontal]
	      }
	      if {[$node1 hasAttribute vertical]} {
		set a(vertical) [$node1 @vertical]
	      }
	      if {[$node1 hasAttribute textRotation]} {
		set a(rotate) [$node1 @textRotation]
	      }
	      if {[$node1 hasAttribute wrapText]} {
		set a(wrap) [$node1 @wrapText]
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
	$doc documentElement root
        $doc selectNodesNamespaces [list M $xmlns(M) r $xmlns(r) mc $xmlns(mc) x14ac $xmlns(x14ac)]
	set idx -1
	foreach col [$root selectNodes /M:worksheet/M:cols/M:col] {
	  incr idx
	  set cols {}
	  foreach item {min max width style bestFit customWidth} {
	    if {[$col hasAttribute $item]} {
	      switch -- $item {
	        min - max {
		  lappend cols [string tolower $item] [expr {[$col @$item] - 1}]
		}
		default {
		  lappend cols [string tolower $item] [$col @$item]
		}
	      }
	    } else {
	      lappend cols [string tolower $item] 0
	    }
	  }
	  lappend cols string 0 nozero 0 calcfit 0
	  set wb($sheet,col,[dict get $cols min]) $cols
	}
	set wb($sheet,cols) [incr idx]
	foreach cell [$root selectNodes /M:worksheet/M:sheetData/M:row/M:c] {
	  if {[$cell hasAttribute t]} {
	    set type [$cell @t]
	  } else {
	    set type n
	  }
	  set value {}
	  set datetime {}

	  if {[$cell hasAttribute r]} {
	    ## FORMULA ##
	    if {!$opts(valuesonly) && [set node [$cell selectNodes M:f]] ne {}} {
	      set rowcol [StringToRowColumn [$cell @r]]
	      set wb($sheet,f,$rowcol) {}
	      if {[set formula [$cell selectNodes M:f/text()]] ne {}} {
		lappend wb($sheet,f,$rowcol) f [$formula nodeValue]
	      }
	      if {[$node hasAttribute t] && [$node @t] eq {shared}} {
	        if {[$node hasAttribute si]} {
		  lappend wb($sheet,f,$rowcol) i [$node @si]
		}
	        if {[$node hasAttribute ref]} {
		  lappend wb($sheet,f,$rowcol) r [$node @ref]
		}
	      }
	    }
	  }

	  switch -- $type {
	    n - b - d - str {
	      # number (default), boolean, iso-date, formula string
	      if {[set node [$cell selectNodes M:v/text()]] ne {}} {
		set value [$node nodeValue]
		if {$type eq {n} && [$cell hasAttribute s] && [string is double -strict $value]} {
		  set idx [$cell @s]
		  if {[info exists cellXfs($idx)] && [dict exists $cellXfs($idx) nfi]} {
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
	      if {[set node [$cell selectNodes M:v/text()]] ne {}} {
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
	      if {[set string [$cell selectNodes M:is]] ne {}} {
		foreach node [$string selectNodes M:t/text()] {
		  append value [$node nodeValue]
		}
		foreach node [$string selectNodes */M:t/text()] {
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
	      set wb($sheet,c,[StringToRowColumn [$cell @r]]) [$cell @r]
	    }
	    if {!$opts(valuesonly)} {
	      if {[$cell hasAttribute s]} {
		set wb($sheet,s,[StringToRowColumn [$cell @r]]) [$cell @s]
	      }
	    }
	    if {!$opts(valuesonly)} {
	      if {[$cell hasAttribute t]} {
		set wb($sheet,t,[StringToRowColumn [$cell @r]]) [$cell @t]
	      }
	    }
	    set wb($sheet,v,[StringToRowColumn [$cell @r]]) $value
	    if {!$opts(valuesonly) && $datetime ne {}} {
	      set wb($sheet,d,[StringToRowColumn [$cell @r]]) $datetime
	    }
	  }
	}

	if {!$opts(valuesonly)} {
	  foreach row [$root selectNodes /M:worksheet/M:sheetData/M:row] {
	    if {[$row hasAttribute r] && [$row hasAttribute ht] && [$row hasAttribute customHeight] && [$row @customHeight] == 1} {
	      dict set wb($sheet,rowheight) [expr {[$row @r] - 1}] [$row @ht]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach freeze [$root selectNodes /M:worksheet/M:sheetViews/M:sheetView/M:pane] {
	    if {[$freeze hasAttribute topLeftCell] && [$freeze hasAttribute state] && [$freeze @state] eq {frozen}} {
	      set wb($sheet,freeze) [$freeze @topLeftCell]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach filter [$root selectNodes /M:worksheet/M:autoFilter] {
	    if {[$filter hasAttribute ref]} {
	      lappend wb($sheet,filter) [$filter @ref]
	    }
	  }
	}
	if {!$opts(valuesonly)} {
	  foreach merge [$root selectNodes /M:worksheet/M:mergeCells/M:mergeCell] {
	    if {[$merge hasAttribute ref]} {
	      lappend wb($sheet,merge) [$merge @ref]
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

# Internal helper
proc ooxml::Dom2zip {zf node path cd count} {
  upvar $cd mycd
  upvar $count mycount
  append mycd [::ooxml::add_str_to_archive $zf $path [$node asXML -indent none -xmlDeclaration 1 -encString "UTF-8"]]
  incr mycount
}


#
# ooxml::InitNodeCommands
#


proc ooxml::InitNodeCommands {} {
  variable initNodeCmds
  variable xmlns

  if {[info exists initNodeCmds] && $initNodeCmds} return

  set elementNodes {
    AppVersion Application
    Company
    Default DocSecurity
    HeadingPairs HyperlinksChanged
    LinksUpToDate
    Override
    Relationship
    ScaleCrop SharedDoc
    TitlesOfParts
    a:accent1 a:accent2 a:accent3 a:accent4 a:accent5 a:accent6 a:alpha a:bevelT a:bgFillStyleLst a:bodyPr a:camera
    a:clrScheme a:cs a:dk1 a:dk2 a:ea a:effectLst a:effectRef a:effectStyle a:effectStyleLst a:extraClrSchemeLst
    a:fillRef a:fillStyleLst a:fillToRect a:fmtScheme a:folHlink a:font a:fontRef a:fontScheme a:gradFill a:gs
    a:gsLst a:hlink a:latin a:lightRig a:lin a:ln a:lnDef a:lnRef a:lnStyleLst a:lstStyle a:lt1 a:lt2 a:majorFont
    a:minorFont a:objectDefaults a:outerShdw a:path a:prstDash a:rot a:satMod a:scene3d a:schemeClr a:shade
    a:solidFill a:sp3d a:spDef a:spPr a:srgbClr a:style a:sysClr a:themeElements a:tint
    alignment autoFilter
    b bgColor bookViews border borders bottom
    c calcPr cellStyle cellStyleXfs cellStyles cellXfs col color cols
    cp:lastModifiedBy
    dc:creator
    dcterms:created dcterms:modified
    definedName definedNames diagonal dimension dxfs
    f family fgColor fileVersion fill fills font fonts
    i
    left
    mergeCell mergeCells
    name numFmt numFmts
    pageMargins pane patternFill
    right row
    scheme sheet sheetData sheetFormatPr sheetView sheetViews sheets si sz
    t tableStyles top
    u
    v
    vt:i4 vt:lpstr vt:variant vt:vector
    workbookPr workbookView
    xf
  }

  namespace eval ::ooxml "dom createNodeCmd textNode Text; namespace export Text"

  # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
  # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
  # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
  #
  # foreach tag $elementNodes {
  #   switch -glob -- $tag {
  #     a:* {
  #       set ns $xmlns(a)
  #     }
  #     cp:* {
  #       set ns $xmlns(cp)
  #     }
  #     dc:* {
  #       set ns $xmlns(dc)
  #     }
  #     dcterms:* {
  #       set ns $xmlns(dcterms)
  #     }
  #     vt:* {
  #       set ns $xmlns(vt)
  #     }
  #     AppVersion - Application - Company - DocSecurity - HeadingPairs - HyperlinksChanged - LinksUpToDate - ScaleCrop - SharedDoc - TitlesOfParts {
  #       set ns $xmlns(EP)
  #     }
  #     Default - Override {
  #       set ns $xmlns(CT)
  #     }
  #     Relationship {
  #       set ns $xmlns(PR)
  #     }
  #     default {
  #       set ns $xmlns(M)
  #     }
  #   }
  #   namespace eval ::ooxml "dom createNodeCmd -tagName $tag -namespace $ns elementNode Tag_$tag; namespace export Tag_$tag"
  # }

  foreach tag $elementNodes {
    namespace eval ::ooxml "dom createNodeCmd -tagName $tag elementNode Tag_$tag; namespace export Tag_$tag"
  }
  
  set initNodeCmds 1
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
    my variable view
    my variable tags

    array set opts {
      creator {unknown}
      created {}
      modifiedby {}
      modified {}
      application {}
    }

    array set tags {}

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -creator - -created - -modifiedby - -modified - -application {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -creator, -created, -modifiedby, -modified or -application"
        }
      }
    }

    set obj(blockPreset) 0

    set obj(encoding) utf-8
    set obj(indent) none

    if {[string trim $opts(creator)] eq {}} {
      set obj(creator) {unknown}
    } else {
      set obj(creator) $opts(creator)
    }
    if {[string trim $opts(created)] eq {} || [catch {clock scan $opts(created)}]} {
      set obj(created) [clock format [clock seconds] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    } elseif {[string is integer -strict $opts(created)] && $opts(created) > 0} {
      set obj(created) [clock format $opts(created) -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    } else {
      set obj(created) [clock format [clock scan $opts(created)] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    }
    if {[string trim $opts(modifiedby)] eq {}} {
      set obj(lastModifiedBy) $opts(creator)
    } else {
      set obj(lastModifiedBy) $opts(modifiedby)
    }
    if {[string trim $opts(modified)] eq {} || [catch {clock scan $opts(modified)}]} {
      set obj(modified) [clock format [clock seconds] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    } elseif {[string is integer -strict $opts(modified)] && $opts(modified) > 0} {
      set obj(modified) [clock format $opts(modified) -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    } else {
      set obj(modified) [clock format [clock scan $opts(modified)] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    }
    if {[string trim $opts(application)] eq {}} {
      set obj(application) {Tcl - Office Open XML - Spreadsheet}
    } else {
      set obj(application) $opts(application)
    }

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
    set styles(0) {numfmt 0 font 0 fill 0 border 0 xf 0 horizontal {} vertical {} rotate {} wrap {}}

    set obj(cols) 0
    array set cols {}

    set obj(calcChain) 0

    set obj(defaultdatestyle) 0

    array set cells {}

    array set view {activetab 0}

    return 0
  }

  destructor {
    return 0
  }

  method numberformat { args } {
    my variable obj
    my variable numFmts
    my variable tags

    array set opts {
      list 0
      format {}
      general 0
      date 0
      time 0
      datetime 0
      iso8601 0
      number 0
      decimal 0
      red 0
      separator 0
      fraction 0
      scientific 0
      percent 0
      string 0
      text 0
      tag {}
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -format {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -list - -general - -date - -time - -datetime - -iso8601 - -number - -decimal - -red - -separator - -fraction - -scientific - -percent - -string - -text {
	  set opts([string range $opt 1 end]) 1
        }
        -tag {
	  incr idx
          if {$idx < $len} {
	    if {[string is integer -strict [set tag [lindex $args $idx]]]} {
	      error "option '$opt': should not be an integer value"
	    } else {
	      set opts([string range $opt 1 end]) $tag
	    } 
	    unset tag
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -format, -list, -general, -date, -time, -datetime, -iso8601, -number, -decimal, -red, -separator, -fraction, -scientific, -percent, -string, -text or tag"
        }
      }
    }

    if {$opts(list)} {
      array set tmp [array get ::ooxml::predefNumFmts]
      array set tmp [array get numFmts]
      return [array get tmp]
    }

    set obj(blockPreset) 1

    if {$opts(general)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 0
      }
      return 0
    }
    if {$opts(date)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 14
      }
      return 14
    }
    if {$opts(time)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 20
      }
      return 20
    }
    if {$opts(number)} {
      if {$opts(separator)} {
	if {$opts(red)} {
	  if {$opts(tag) ne {}} {
	    set tags(numFmts,$opts(tag)) 38
	  }
	  return 38
	} else {
	  if {$opts(tag) ne {}} {
	    set tags(numFmts,$opts(tag)) 3
	  }
	  return 3
	}
      } else {
	if {$opts(red)} {
	  return -1
	} else {
	  if {$opts(tag) ne {}} {
	    set tags(numFmts,$opts(tag)) 1
	  }
	  return 1
	}
      }
    }
    if {$opts(decimal)} {
      if {$opts(percent)} {
	if {$opts(tag) ne {}} {
	  set tags(numFmts,$opts(tag)) 10
	}
        return 10
      }
      if {$opts(separator)} {
	if {$opts(red)} {
	  if {$opts(tag) ne {}} {
	    set tags(numFmts,$opts(tag)) 40
	  }
	  return 40
	} else {
	  if {$opts(tag) ne {}} {
	    set tags(numFmts,$opts(tag)) 4
	  }
	  return 4
	}
      } else {
	if {$opts(red)} {
	  return -1
	} else {
	  if {$opts(tag) ne {}} {
	    set tags(numFmts,$opts(tag)) 2
	  }
	  return 2
	}
      }
    }
    if {$opts(fraction)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 12
      }
      return 12
    }
    if {$opts(scientific)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 11
      }
      return 11
    }
    if {$opts(percent)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 9
      }
      return 9
    }
    if {$opts(text) || $opts(string)} {
      if {$opts(tag) ne {}} {
	set tags(numFmts,$opts(tag)) 49
      }
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
	if {$opts(tag) ne {}} {
	  set tags(numFmts,$opts(tag)) $idx
	}
        return $idx
      }
    }
    foreach idx [array names numFmts] {
      if {$numFmts($idx) eq $opts(format)} {
	if {$opts(tag) ne {}} {
	  set tags(numFmts,$opts(tag)) $idx
	}
        return $idx
      }
    }

    if {$opts(format) eq {}} {
      return -1
    }

    set idx $obj(numFmts)
    set numFmts($idx) $opts(format)
    incr obj(numFmts)

    if {$opts(tag) ne {}} {
      set tags(numFmts,$opts(tag)) $idx
    }
    return $idx
  }

  method defaultdatestyle { style } {
    my variable obj
    my variable tags

    if {![string is integer -strict $style] && [info exists tags(styles,$style)]} {
      set style $tags(styles,$style)]
    }
    set obj(defaultdatestyle) $style
  }

  method font { args } {
    my variable obj
    my variable fonts
    my variable tags

    array set a $fonts(0)

    array set opts "
      list 0
      name [list $a(name)]
      family [list $a(family)]
      size [list $a(size)]
      color [list $a(color)]
      scheme [list $a(scheme)]
      bold 0
      italic 0
      underline 0
      tag {}
    "

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -name - -family - -size - -color - -scheme {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -list - -bold - -italic - -underline {
	  set opts([string range $opt 1 end]) 1
        }
        -tag {
	  incr idx
          if {$idx < $len} {
	    if {[string is integer -strict [set tag [lindex $args $idx]]]} {
	      error "option '$opt': should not be an integer value"
	    } else {
	      set opts([string range $opt 1 end]) $tag
	    } 
	    unset tag
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -name, -family, -size, -color, -scheme, -list, -bold, -italic, -underline or -tag"
        }
      }
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
	if {$opts(tag) ne {}} {
	  set tags(fonts,$opts(tag)) $idx
	}
        return $idx
      }
    }

    set fonts($obj(fonts)) {}
    foreach item {name family size bold italic underline color scheme} {
      lappend fonts($obj(fonts)) $item $opts($item)
    }
    set idx $obj(fonts)
    incr obj(fonts)
    if {$opts(tag) ne {}} {
      set tags(fonts,$opts(tag)) $idx
    }
    return $idx
  }

  method fill { args } {
    my variable obj
    my variable fills
    my variable tags

    array set opts {
      list 0
      patterntype none
      fgcolor {}
      bgcolor {}
      tag {}
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -patterntype - -fgcolor - -bgcolor {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -list {
	  set opts([string range $opt 1 end]) 1
        }
        -tag {
	  incr idx
          if {$idx < $len} {
	    if {[string is integer -strict [set tag [lindex $args $idx]]]} {
	      error "option '$opt': should not be an integer value"
	    } else {
	      set opts([string range $opt 1 end]) $tag
	    } 
	    unset tag
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -patterntype, -fgcolor, -bgcolor, -list or -tag"
        }
      }
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
	if {$opts(tag) ne {}} {
	  set tags(fills,$opts(tag)) $idx
	}
        return $idx
      }
    }

    set fills($obj(fills)) {}
    foreach item {patterntype fgcolor bgcolor} {
      lappend fills($obj(fills)) $item $opts($item)
    }
    set idx $obj(fills)
    incr obj(fills)
    
    if {$opts(tag) ne {}} {
      set tags(fills,$opts(tag)) $idx
    }
    return $idx
  }

  method border { args } {
    my variable obj
    my variable borders
    my variable tags

    array set opts {
      list 0
      leftstyle {}
      leftcolor {}
      rightstyle {}
      rightcolor {}
      topstyle {}
      topcolor {}
      bottomstyle {}
      bottomcolor {}
      diagonalstyle {}
      diagonalcolor {}
      diagonaldirection {}
      tag {}
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
	-leftstyle - -leftcolor - -rightstyle - -rightcolor - -topstyle - -topcolor - -bottomstyle - -bottomcolor - -diagonalstyle - -diagonalcolor - -diagonaldirection {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -list {
	  set opts([string range $opt 1 end]) 1
        }
        -tag {
	  incr idx
          if {$idx < $len} {
	    if {[string is integer -strict [set tag [lindex $args $idx]]]} {
	      error "option '$opt': should not be an integer value"
	    } else {
	      set opts([string range $opt 1 end]) $tag
	    } 
	    unset tag
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -leftstyle, -leftcolor, -rightstyle, -rightcolor, -topstyle, -topcolor, -bottomstyle, -bottomcolor, -diagonalstyle, -diagonalcolor, -diagonaldirection, -list or -tag"
        }
      }
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
	if {$opts(tag) ne {}} {
	  set tags(borders,$opts(tag)) $idx
	}
        return $idx
      }
    }

    set borders($obj(borders)) $tmp
    set idx $obj(borders)
    incr obj(borders)
    
    if {$opts(tag) ne {}} {
      set tags(borders,$opts(tag)) $idx
    }
    return $idx
  }

  method style { args } {
    my variable obj
    my variable styles
    my variable tags

    array set opts {
      list 0
      numfmt 0
      font 0
      fill 0
      border 0
      xf 0
      horizontal {}
      vertical {}
      rotate {}
      wrap 0
      tag {}
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -numfmt - -font - -fill - -border - -xf - -horizontal - -vertical - -rotate {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -list - -wrap {
	  set opts([string range $opt 1 end]) 1
        }
        -tag {
	  incr idx
          if {$idx < $len} {
	    if {[string is integer -strict [set tag [lindex $args $idx]]]} {
	      error "option '$opt': should not be an integer value"
	    } else {
	      set opts([string range $opt 1 end]) $tag
	    } 
	    unset tag
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -numfmt, -font, -fill, -border, -xf, -horizontal, -vertical, -rotate, -list, -wrap or -tag"
        }
      }
    }

    if {$opts(list)} {
      return [array get styles]
    }

    if {![string is integer -strict $opts(numfmt)] && [info exists tags(numFmts,$opts(numfmt))]} {
      set opts(numfmt) $tags(numFmts,$opts(numfmt))
    }
    if {![string is integer -strict $opts(font)] && [info exists tags(fonts,$opts(font))]} {
      set opts(font) $tags(fonts,$opts(font))
    }
    if {![string is integer -strict $opts(fill)] && [info exists tags(fills,$opts(fill))]} {
      set opts(fill) $tags(fills,$opts(fill))
    }
    if {![string is integer -strict $opts(border)] && [info exists tags(borders,$opts(border))]} {
      set opts(border) $tags(borders,$opts(border))
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
    if {$opts(wrap) ne {1}} {
      set opts(wrap) {}
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
	if {$opts(tag) ne {}} {
	  set tags(styles,$opts(tag)) $idx
	}
        return $idx
      }
    }

    set styles($obj(styles)) {}
    foreach item {numfmt font fill border xf horizontal vertical rotate wrap} {
      lappend styles($obj(styles)) $item $opts($item)
    }
    set idx $obj(styles)
    incr obj(styles)
    if {$opts(tag) ne {}} {
      set tags(styles,$opts(tag)) $idx
    }
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
    my variable tags

    array set opts {
      index {}
      to {}
      width {}
      style 0
      bestfit 0
      customwidth 0
      string 0
      nozero 0
      calcfit 0
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -index - -to - -width - -style  {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -bestfit - -customwidth - -string - -nozero - -calcfit {
	  set opts([string range $opt 1 end]) 1
        }
        default {
          error "unknown option \"$opt\", should be: -index, -to, -width, -style, -bestfit, -customwidth, -string, -nozero or -calcfit"
        }
      }
    }

    if {![string is integer -strict $opts(style)] && [info exists tags(styles,$opts(style))]} {
      set opts(style) $tags(styles,$opts(style))
    }

    if {[regexp {^\d+$} $opts(index)] && $opts(index) > -1} {
      set obj(gCol,$sheet) $opts(index)
    } elseif {[regexp {^[A-Z]+$} $opts(index)]} {
      set obj(gCol,$sheet) [lindex [split [::ooxml::StringToRowColumn $opts(index)] ,] end]
    } elseif {[string trim $opts(index)] eq {}} {
      incr obj(gCol,$sheet)
    }
    set opts(index) $obj(gCol,$sheet)

    if {[regexp {^[A-Z]+$} $opts(to)]} {
      set opts(to) [lindex [split [::ooxml::StringToRowColumn $opts(to)] ,] end]
    } elseif {[string trim $opts(to)] eq {} || ![string is integer -strict $opts(to)] || $opts(to) <= $opts(index)} {
      set opts(to) $opts(index)
    }

    if {![string is double -strict $opts(width)] || $opts(width) < 0} {
      set opts(width) {}
    }
    if {![string is integer -strict $opts(style)] || $opts(style) < 0} {
      set opts(style) {}
    }

    if {$opts(width) ne {} || ([string is integer -strict $opts(style)] && $opts(style) >= 0) || $opts(bestfit) > 0} {
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

    array set opts {
      index {}
      height {}
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -index - -height  {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        default {
          error "unknown option \"$opt\", should be: -index or -height"
        }
      }
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
    my variable tags

    array set opts {
      index {}
      style -1
      formula {}
      formulaidx {}
      formularef {}
      string -1
      nostring -1
      zero -1
      nozero -1
      globalstyle {}
      height {}
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -index - -style - -formula - -formulaidx - -formularef - -height  {
	  incr idx
          if {$idx < $len} {
            set opts([string range $opt 1 end]) [lindex $args $idx]
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -string - -nostring - -zero - -nozero - -globalstyle {
	  set opts([string range $opt 1 end]) 1
        }
        default {
          error "unknown option \"$opt\", should be: -index, -style, -formula, -formulaidx, -formularef, -height, -string, nostring, -zero or -nozero"
        }
      }
    }

    if {![string is integer -strict $opts(style)] && [info exists tags(styles,$opts(style))]} {
      set opts(style) $tags(styles,$opts(style))
    }

    if {$opts(nostring) == 1} {
      set opts(string) 0
    }
    if {$opts(zero) == 1} {
      set opts(nozero) 0
    }

    if {!$obj(callRow,$obj(sheets))} {
      set obj(callRow,$obj(sheets)) 1
      incr obj(row,$sheet)
    }

    if {[regexp {^\d+$} $opts(index)] && $opts(index) > -1} {
      set obj(col,$sheet) $opts(index)
    } elseif {[regexp {^[A-Z]+$} $opts(index)]} {
      set obj(col,$sheet) [lindex [split [::ooxml::StringToRowColumn $opts(index)] ,] end]
    } elseif {[regexp {^(\d+),(\d+)$} $opts(index)]} {
      lassign [split $opts(index) ,] obj(row,$sheet) obj(col,$sheet)
    } elseif {[regexp {^[A-Z]+\d+$} $opts(index)]} {
      lassign [split [::ooxml::StringToRowColumn $opts(index)] ,] obj(row,$sheet) obj(col,$sheet)
    } elseif {[string trim $opts(index)] eq {}} {
      incr obj(col,$sheet)
    }
    if {$obj(row,$sheet) < 0 || $obj(col,$sheet) < 0} {
      return -1
    }

    if {[string is integer -strict $opts(style)] && $opts(style) == -1} {
      if {[info exists cols($sheet,$obj(col,$sheet))] && [dict get $cols($sheet,$obj(col,$sheet)) style] >= 0} {
        set opts(style) [dict get $cols($sheet,$obj(col,$sheet)) style]
      } elseif {$opts(style) == -1} {
	set opts(style) 0
      }
    }
    if {[string is integer -strict $opts(string)] && $opts(string) == -1} {
      if {[info exists cols($sheet,$obj(col,$sheet))] && [dict get $cols($sheet,$obj(col,$sheet)) string] == 1} {
	set opts(string) 1
      } elseif {$opts(string) == -1} {
	set opts(string) 0
      }
    }
    if {[string is integer -strict $opts(nozero)] && $opts(nozero) == -1} {
      if {[info exists cols($sheet,$obj(col,$sheet))] && [dict get $cols($sheet,$obj(col,$sheet)) nozero] == 1} {
	set opts(nozero) 1
      } elseif {$opts(nozero) == -1} {
	set opts(nozero) 0
      }
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
    ## FORMULA ##
    if {[string trim $opts(formula)] ne {} || [string is integer -strict $opts(formulaidx)]} {
      lappend cells($cell) t $type
      lappend cells($cell) f $opts(formula)
      if {[string trim $opts(formulaidx)] ne {}} {
	lappend cells($cell) fsi $opts(formulaidx)
      }
      if {[string trim $opts(formularef)] ne {}} {
        set lref [split [string trim $opts(formularef)] :]
	if {[llength $lref] == 2} {
	  set fref {}
	  if {[regexp {^(\d+),(\d+)$} [lindex $lref 0]]} {
	    append fref [::ooxml::IndexToString [lindex $lref 0]]
	  } else {
	    append fref [lindex $lref 0]
	  }
	  append fref :
	  if {[regexp {^(\d+),(\d+)$} [lindex $lref 1]]} {
	    append fref [::ooxml::IndexToString [lindex $lref 1]]
	  } else {
	    append fref [lindex $lref 1]
	  }
	  lappend cells($cell) fsr $fref
	}
      }
    } else {
      lappend cells($cell) v $data t $type
    }

    if {[string trim $data] eq {} && [string trim $opts(formula)] eq {} && [string trim $opts(formulaidx)] eq {} && ![string is integer -strict $opts(style)] && $opts(style) < 1} {
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
      foreach item [array names a $sheet,col,*] {
	set idx [lindex [split $item ,] end]
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
	      ## FORMULA ##
	      if {[dict exists $a($sheet,f,$row,$col) f]} {
		lappend options -formula [dict get $a($sheet,f,$row,$col) f]
		if {[dict exists $a($sheet,f,$row,$col) r]} {
		  lappend options -formularef [dict get $a($sheet,f,$row,$col) r]
		}
	      }
	      if {[dict exists $a($sheet,f,$row,$col) i]} {
		lappend options -formulaidx [dict get $a($sheet,f,$row,$col) i]
	      }
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

    if {[info exists a(view)]} {
      foreach item {activetab} {
        if {[dict exists $a(view) $item]} {
	  my view -$item [dict get $a(view) $item]
	}
      }
    }
  }

  method view { args } {
    my variable view

    array set opts {
      list 0
    }

    set len [llength $args]
    set idx 0
    for {set idx 0} {$idx < $len} {incr idx} {
      switch -- [set opt [lindex $args $idx]] {
        -activetab - -x - -y - -height - -width {
	  incr idx
          if {$idx < $len} {
	    if {[string is integer -strict [lindex $args $idx]] && [lindex $args $idx] > -1} {
	      set view([string range $opt 1 end]) [lindex $args $idx]
	    } else {
	      error "option '$opt': must be a positive integer"
	    }
          } else {
            error "option '$opt': missing argument"
          }            
        }
        -list {
	  set opts([string range $opt 1 end]) 1
        }
        default {
          error "unknown option \"$opt\", should be: -activetab, -x, -y, -height, -width or -list"
        }
      }
    }

    if {$opts(list)} {
      return [array get view]
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

  method write { file } {
    my variable obj
    my variable cells
    my variable sharedStrings
    my variable fonts
    my variable numFmts
    my variable styles
    my variable fills
    my variable borders
    my variable cols
    my variable view

    upvar #0 ::ooxml::xmlns xmlns

    ooxml::InitNodeCommands
    namespace import ::ooxml::Tag_* ::ooxml::Text

    # Initialize zip file
    set file [string trim $file]
    if {$file eq {}} {
      set file {spreadsheetml.xlsx}
    }
    if {[file extension $file] ne {.xlsx}} {
      append file {.xlsx}
    }
    if {[catch {open $file w} zf]} {
      error "cannot open file $file for writing"
    }
    fconfigure $zf -encoding binary -translation binary -eofchar {}
    set count 0
    set cd {}
    
    foreach {n v} [array get cells] {
      if {[dict exists $v t] && [dict get $v t] eq {s} && [dict exists $v v] && [dict get $v v] ne {}} {
        set thisv [dict get $v v]
        if {[info exists lookup($thisv)]} {
          set pos $lookup($thisv)
        } else {
          set pos [llength $sharedStrings]
	  lappend sharedStrings $thisv
          set lookup($thisv) $pos
	}
        set obj(sharedStrings) 1
	dict set cells($n) v $pos
      }
    }
    unset -nocomplain n v
    array unset lookup
    
    # ------------------------------ _rels/.rels ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(PR) Relationships]
    # set root [$doc documentElement]

    dom createDocument Relationships doc
    $doc documentElement root

    $root setAttribute xmlns $xmlns(PR)

    set rId 0

    $root appendFromScript {
      Tag_Relationship Id rId1 Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument Target xl/workbook.xml {}
      Tag_Relationship Id rId2 Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties Target docProps/app.xml {}
      Tag_Relationship Id rId3 Type http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties Target docProps/core.xml {}
    }
    ::ooxml::Dom2zip $zf $root "_rels/.rels" cd count
    $doc delete

    # ------------------------------ [Content_Types].xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(CT) Types]
    # set root [$doc documentElement]

    dom createDocument Types doc
    $doc documentElement root

    $root setAttribute xmlns $xmlns(CT)

    $root appendFromScript {
      Tag_Default Extension xml ContentType application/xml {}
      Tag_Default Extension rels ContentType application/vnd.openxmlformats-package.relationships+xml {}
      Tag_Override PartName /xl/workbook.xml ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml {}
      Tag_Override PartName /xl/worksheets/sheet1.xml ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml {}
      for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
	Tag_Override PartName /xl/theme/theme${ws}.xml ContentType application/vnd.openxmlformats-officedocument.theme+xml {}
      }
      Tag_Override PartName /xl/styles.xml ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml {}
      if {$obj(sharedStrings) > 0} {
	Tag_Override PartName /xl/sharedStrings.xml ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml {}
      }
      if {$obj(calcChain)} {
	Tag_Override PartName /xl/calcChain.xml ContentType application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml {}
      }
      Tag_Override PartName /docProps/core.xml ContentType application/vnd.openxmlformats-package.core-properties+xml {}
      Tag_Override PartName /docProps/app.xml ContentType application/vnd.openxmlformats-officedocument.extended-properties+xml {}
    }
    ::ooxml::Dom2zip $zf $root "\[Content_Types\].xml" cd count
    $doc delete

    # ------------------------------ docProps/app.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(EP) Properties]
    # set root [$doc documentElement]
    # $root setAttributeNS {} xmlns:vt $xmlns(vt)

    dom createDocument Properties doc
    $doc documentElement root

    $root setAttribute xmlns $xmlns(EP)
    $root setAttribute xmlns:vt $xmlns(vt)

    $root appendFromScript {
      Tag_Application { Text $obj(application) }
      Tag_DocSecurity { Text 0 }
      Tag_ScaleCrop { Text false }
      Tag_HeadingPairs {
	Tag_vt:vector size 2 baseType variant {
	  Tag_vt:variant {
	    Tag_vt:lpstr { Text [msgcat::mc Worksheets] }
	  }
	  Tag_vt:variant {
	    Tag_vt:i4 { Text $obj(sheets) }
	  }
	}
      }
      Tag_TitlesOfParts {
	Tag_vt:vector size $obj(sheets) baseType lpstr {
	  for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
	    Tag_vt:lpstr {
	      Text [msgcat::mc Sheet]$ws
	    }
	  }
	}
      }
      Tag_Company {}
      Tag_LinksUpToDate { Text false }
      Tag_SharedDoc { Text false }
      Tag_HyperlinksChanged { Text false }
      Tag_AppVersion { Text 1.0 }
    }
    ::ooxml::Dom2zip $zf $root "docProps/app.xml" cd count
    $doc delete

    # ------------------------------ docProps/core.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(cp) cp:coreProperties]
    # set root [$doc documentElement]
    # $root setAttributeNS {} xmlns:dc $xmlns(dc)
    # $root setAttributeNS {} xmlns:dcterms $xmlns(dcterms)
    # $root setAttributeNS {} xmlns:dcmitype $xmlns(dcmitype)
    # $root setAttributeNS {} xmlns:xsi $xmlns(xsi)

    dom createDocument cp:coreProperties doc
    $doc documentElement root

    $root setAttribute xmlns:cp $xmlns(cp)
    $root setAttribute xmlns:dc $xmlns(dc)
    $root setAttribute xmlns:dcterms $xmlns(dcterms)
    $root setAttribute xmlns:dcmitype $xmlns(dcmitype)
    $root setAttribute xmlns:xsi $xmlns(xsi)

    $root appendFromScript {
      Tag_dc:creator { Text $obj(creator) }
      Tag_cp:lastModifiedBy { Text $obj(lastModifiedBy) }
      Tag_dcterms:created xsi:type dcterms:W3CDTF { Text $obj(created) }
      Tag_dcterms:modified xsi:type dcterms:W3CDTF { Text $obj(modified) }
    }
    ::ooxml::Dom2zip $zf $root "docProps/core.xml" cd count
    $doc delete

    # ------------------------------ xl/_rels/workbook.xml.rels ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(PR) Relationships]
    # set root [$doc documentElement]

    dom createDocument Relationships doc
    $doc documentElement root

    $root setAttribute xmlns $xmlns(PR)

    $root appendFromScript {
      for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
	Tag_Relationship Id rId$ws Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet Target worksheets/sheet${ws}.xml {}
      }
      set rId [incr ws -1]
      Tag_Relationship Id rId[incr rId] Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme Target theme/theme1.xml {}
      Tag_Relationship Id rId[incr rId] Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles Target styles.xml {}
      if {$obj(sharedStrings) > 0} {
	Tag_Relationship Id rId[incr rId] Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings Target sharedStrings.xml {}
      }
      if {$obj(calcChain)} {
	Tag_Relationship Id rId[incr rId] Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain Target calcChain.xml {}
      }
    }
    ::ooxml::Dom2zip $zf $root "xl/_rels/workbook.xml.rels" cd count
    $doc delete


    # ------------------------------ xl/sharedStrings.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(M) sst]
    # set root [$doc documentElement]

    if {$obj(sharedStrings) > 0} {
      dom createDocument sst doc
      $doc documentElement root

      $root setAttribute xmlns $xmlns(M)

      $root setAttribute count [llength $sharedStrings]
      $root setAttribute uniqueCount [llength $sharedStrings]

      $root appendFromScript {
	foreach string $sharedStrings {
	  Tag_si {
	    Tag_t { Text $string }
	  }
	}
	# garbage collection
	set sharedStrings {}
      }
      ::ooxml::Dom2zip $zf $root "xl/sharedStrings.xml" cd count
      $doc delete
    }


    # ------------------------------ xl/calcChain.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(M) calcChain]
    # set root [$doc documentElement]

    if {$obj(calcChain)} {
      dom createDocument calcChain doc
      $doc documentElement root

      $root setAttribute xmlns $xmlns(M)

      $root appendFromScript {
	Tag_c r C1 i 3 l 1 {}
	Tag_c r A3 i 2 {}
      }
      ::ooxml::Dom2zip $zf $root "xl/calcChain.xml" cd count
      $doc delete
    }


    # ------------------------------ xl/styles.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(M) styleSheet]
    # set root [$doc documentElement]
    # $root setAttributeNS {} xmlns:mc $xmlns(mc)
    # $root setAttributeNS {} xmlns:x14ac $xmlns(x14ac)

    dom createDocument styleSheet doc
    $doc documentElement root

    $root setAttribute xmlns $xmlns(M)
    $root setAttribute xmlns:mc $xmlns(mc)
    $root setAttribute xmlns:x14ac $xmlns(x14ac)

    $root setAttribute mc:Ignorable x14ac

    $root appendFromScript {
      if {$obj(numFmts) > $::ooxml::defaults(numFmts,start)} {
	Tag_numFmts count [llength [array names numFmts]] {
	  foreach idx [lsort -integer [array names numFmts]] {
	    Tag_numFmt numFmtId $idx formatCode $numFmts($idx) {}
	  }
	}
      }
      Tag_fonts count $obj(fonts) x14ac:knownFonts 1 {
	foreach idx [lsort -integer [array names fonts]] {
	  Tag_font {
	    if {[dict get $fonts($idx) color] ne {}} {
	      Tag_color [lindex [dict get $fonts($idx) color] 0] [lindex [dict get $fonts($idx) color] 1]
	    }
	    if {[dict get $fonts($idx) bold] == 1} {
	      Tag_b {}
	    }
	    if {[dict get $fonts($idx) italic] == 1} {
	      Tag_i {}
	    }
	    if {[dict get $fonts($idx) underline] == 1} {
	      Tag_u {}
	    }
	    Tag_sz val [dict get $fonts($idx) size] {}
	    Tag_name val [dict get $fonts($idx) name] {}
	    Tag_family val [dict get $fonts($idx) family] {}
	    Tag_scheme val [dict get $fonts($idx) scheme] {}
	  }
	}
      }
      if {$obj(fills) > 0} {
	Tag_fills count $obj(fills) {
	  foreach idx [lsort -integer [array names fills]] {
	    Tag_fill {
	      Tag_patternFill patternType [dict get $fills($idx) patterntype] {
		foreach tag {fgColor bgColor} {
		  set key [string tolower $tag]
		  if {[dict get $fills($idx) $key] ne {}} {
		    Tag_$tag [lindex [dict get $fills($idx) $key] 0] [lindex [dict get $fills($idx) $key] 1] {}
		  }
		}
	      }
	    }
	  }
	}
      }
      if {$obj(borders) > 0} {
	Tag_borders count $obj(borders) {
	  foreach idx [lsort -integer [array names borders]] {
	    set attr {}
	    if {[dict exists $borders($idx) diagonal direction] && [dict get $borders($idx) diagonal direction] ne {}} {
	      lappend attr [string map {up diagonalUp down diagonalDown} [dict get $borders($idx) diagonal direction]] 1
	    }
	    Tag_border {*}$attr {
	      foreach item {left right top bottom diagonal} {
		set attr {}
		if {[dict exists $borders($idx) $item style] && [dict get $borders($idx) $item style] ne {}} {
		  lappend attr style [dict get $borders($idx) $item style]
		}
	        Tag_$item {*}$attr {
		  if {[dict exists $borders($idx) $item color] && [dict get $borders($idx) $item color] ne {}} {
		    Tag_color [lindex [dict get $borders($idx) $item color] 0] [lindex [dict get $borders($idx) $item color] 1] {}
		  }
		}
	      }
	    }
	  }
	}
      }
      Tag_cellStyleXfs count 1 {
	Tag_xf numFmtId 0 fontId 0 fillId 0 borderId 0 {}
      }
      Tag_cellXfs count $obj(styles) {
	foreach idx [lsort -integer [array names styles]] {
	  set attr {}
	  lappend attr numFmtId [dict get $styles($idx) numfmt]
	  lappend attr fontId [dict get $styles($idx) font]
	  lappend attr fillId [dict get $styles($idx) fill]
	  lappend attr borderId [dict get $styles($idx) border]
	  lappend attr xfId [dict get $styles($idx) xf]
	  if {[dict get $styles($idx) numfmt] > 0} {
	    lappend attr applyNumberFormat 1
	  }
	  if {[dict get $styles($idx) font] > 0} {
	    lappend attr applyFont 1
	  }
	  if {[dict get $styles($idx) fill] > 0} {
	    lappend attr applyFill 1
	  }
	  if {[dict get $styles($idx) border] > 0} {
	    lappend attr applyBorder 1
	  }
	  # lappend attr applyProtection 1 quotePrefix 1
	  if {[dict get $styles($idx) horizontal] ne {} || [dict get $styles($idx) vertical] ne {} || [dict get $styles($idx) rotate] ne {} || [dict get $styles($idx) wrap] ne {}} {
	    lappend attr applyAlignment 1
	    set alignment 1
	  } else {
	    set alignment 0
	  }
	  Tag_xf {*}$attr {
	    set attr {}
	    if {$alignment} {
	      if {[dict get $styles($idx) horizontal] ne {}} {
		lappend attr horizontal [dict get $styles($idx) horizontal]
	      }
	      if {[dict get $styles($idx) vertical] ne {}} {
		lappend attr vertical [dict get $styles($idx) vertical]
	      }
	      if {[dict get $styles($idx) rotate] ne {}} {
		lappend attr textRotation [dict get $styles($idx) rotate]
	      }
	      if {[dict get $styles($idx) wrap] ne {}} {
		lappend attr wrapText [dict get $styles($idx) wrap]
	      }
	      Tag_alignment {*}$attr {}
	    }
	  }
	}
      }
      Tag_cellStyles count 1 {
	Tag_cellStyle name Standard xfId 0 builtinId 0 {}
      }
      Tag_dxfs count 0 {}
      Tag_tableStyles count 0 {}
    }
    ::ooxml::Dom2zip $zf $root "xl/styles.xml" cd count
    $doc delete


    # ------------------------------ xl/theme/theme1.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    set doc [dom createDocumentNS $xmlns(a) a:theme]
    set root [$doc documentElement]

    dom createDocument a:theme doc
    $doc documentElement root

    $root setAttribute xmlns:a $xmlns(a)

    $root setAttribute name Office-Design

    $root appendFromScript {
      Tag_a:themeElements {
	Tag_a:clrScheme name Office {
	  Tag_a:dk1 {
	    Tag_a:sysClr val windowText lastClr 000000 {}
	  }
	  Tag_a:lt1 {
	    Tag_a:sysClr val window lastClr FFFFFF {}
	  }
	  Tag_a:dk2 {
	    Tag_a:srgbClr val 1F497D {}
	  }
	  Tag_a:lt2 {
	    Tag_a:srgbClr val EEECE1 {}
	  }
	  Tag_a:accent1 {
	    Tag_a:srgbClr val 4F81BD {}
	  }
	  Tag_a:accent2 {
	    Tag_a:srgbClr val C0504D {}
	  }
	  Tag_a:accent3 {
	    Tag_a:srgbClr val 9BBB59 {}
	  }
	  Tag_a:accent4 {
	    Tag_a:srgbClr val 8064A2 {}
	  }
	  Tag_a:accent5 {
	    Tag_a:srgbClr val 4BACC6 {}
	  }
	  Tag_a:accent6 {
	    Tag_a:srgbClr val F79646 {}
	  }
	  Tag_a:hlink {
	    Tag_a:srgbClr val 0000FF {}
	  }
	  Tag_a:folHlink {
	    Tag_a:srgbClr val 800080 {}
	  }
	}
	Tag_a:fontScheme name Office {
	  Tag_a:majorFont {
	    Tag_a:latin typeface Cambria {}
	    Tag_a:ea typeface {} {}
	    Tag_a:cs typeface {} {}
	    Tag_a:font script Jpan typeface \uFF2D\uFF33\u0020\uFF30\u30B4\u30B7\u30C3\u30AF {}
	    Tag_a:font script Hang typeface \uB9D1\uC740\u0020\uACE0\uB515 {}
	    Tag_a:font script Hans typeface \u5B8B\u4F53 {}
	    Tag_a:font script Hant typeface \u65B0\u7D30\u660E\u9AD4 {}
	    Tag_a:font script Arab typeface {Times New Roman} {}
	    Tag_a:font script Hebr typeface {Times New Roman} {}
	    Tag_a:font script Thai typeface Tahoma {}
	    Tag_a:font script Ethi typeface Nyala {}
	    Tag_a:font script Beng typeface Vrinda {}
	    Tag_a:font script Gujr typeface Shruti {}
	    Tag_a:font script Khmr typeface MoolBoran {}
	    Tag_a:font script Knda typeface Tunga {}
	    Tag_a:font script Guru typeface Raavi {}
	    Tag_a:font script Cans typeface Euphemia {}
	    Tag_a:font script Cher typeface {Plantagenet Cherokee} {}
	    Tag_a:font script Yiii typeface {Microsoft Yi Baiti} {}
	    Tag_a:font script Tibt typeface {Microsoft Himalaya} {}
	    Tag_a:font script Thaa typeface {MV Boli} {}
	    Tag_a:font script Deva typeface Mangal {}
	    Tag_a:font script Telu typeface Gautami {}
	    Tag_a:font script Taml typeface Latha {}
	    Tag_a:font script Syrc typeface {Estrangelo Edessa} {}
	    Tag_a:font script Orya typeface Kalinga {}
	    Tag_a:font script Mlym typeface Kartika {}
	    Tag_a:font script Laoo typeface DokChampa {}
	    Tag_a:font script Sinh typeface {Iskoola Pota} {}
	    Tag_a:font script Mong typeface {Mongolian Baiti} {}
	    Tag_a:font script Viet typeface {Times New Roman} {}
	    Tag_a:font script Uigh typeface {Microsoft Uighur} {}
	    Tag_a:font script Geor typeface Sylfaen {}
	  }
	  Tag_a:minorFont {
	    Tag_a:latin typeface Calibri {}
	    Tag_a:ea typeface {} {}
	    Tag_a:cs typeface {} {}
	    Tag_a:font script Jpan typeface \uFF2D\uFF33\u0020\uFF30\u30B4\u30B7\u30C3\u30AF {}
	    Tag_a:font script Hang typeface \uB9D1\uC740\u0020\uACE0\uB515 {}
	    Tag_a:font script Hans typeface \u5B8B\u4F53 {}
	    Tag_a:font script Hant typeface \u65B0\u7D30\u660E\u9AD4 {}
	    Tag_a:font script Arab typeface Arial {}
	    Tag_a:font script Hebr typeface Arial {}
	    Tag_a:font script Thai typeface Tahoma {}
	    Tag_a:font script Ethi typeface Nyala {}
	    Tag_a:font script Beng typeface Vrinda {}
	    Tag_a:font script Gujr typeface Shruti {}
	    Tag_a:font script Khmr typeface DaunPenh {}
	    Tag_a:font script Knda typeface Tunga {}
	    Tag_a:font script Guru typeface Raavi {}
	    Tag_a:font script Cans typeface Euphemia {}
	    Tag_a:font script Cher typeface {Plantagenet Cherokee} {}
	    Tag_a:font script Yiii typeface {Microsoft Yi Baiti} {}
	    Tag_a:font script Tibt typeface {Microsoft Himalaya} {}
	    Tag_a:font script Thaa typeface {MV Boli} {}
	    Tag_a:font script Deva typeface Mangal {}
	    Tag_a:font script Telu typeface Gautami {}
	    Tag_a:font script Taml typeface Latha {}
	    Tag_a:font script Syrc typeface {Estrangelo Edessa} {}
	    Tag_a:font script Orya typeface Kalinga {}
	    Tag_a:font script Mlym typeface Kartika {}
	    Tag_a:font script Laoo typeface DokChampa {}
	    Tag_a:font script Sinh typeface {Iskoola Pota} {}
	    Tag_a:font script Mong typeface {Mongolian Baiti} {}
	    Tag_a:font script Viet typeface Arial {}
	    Tag_a:font script Uigh typeface {Microsoft Uighur} {}
	    Tag_a:font script Geor typeface Sylfaen {}
	  }
	}
	Tag_a:fmtScheme name Office {
	  Tag_a:fillStyleLst {
	    Tag_a:solidFill {
	      Tag_a:schemeClr val phClr {}
	    }
	    Tag_a:gradFill rotWithShape 1 {
	      Tag_a:gsLst {
		Tag_a:gs pos 0 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 50000 {}
		    Tag_a:satMod val 300000 {}
		  }
		}
		Tag_a:gs pos 35000 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 37000 {}
		    Tag_a:satMod val 300000 {}
		  }
		}
		Tag_a:gs pos 100000 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 15000 {}
		    Tag_a:satMod val 350000 {}
		  }
		}
	      }
	      Tag_a:lin ang 16200000 scaled 1 {}
	    }
	    Tag_a:gradFill rotWithShape 1 {
	      Tag_a:gsLst {
		Tag_a:gs pos 0 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 100000 {}
		    Tag_a:shade val 100000 {}
		    Tag_a:satMod val 130000 {}
		  }
		}
		Tag_a:gs pos 100000 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 50000 {}
		    Tag_a:shade val 100000 {}
		    Tag_a:satMod val 350000 {}
		  }
		}
	      }
	      Tag_a:lin ang 16200000 scaled 0 {}
	    }
	  }
	  Tag_a:lnStyleLst {
	    Tag_a:ln w 9525 cap flat cmpd sng algn ctr {
	      Tag_a:solidFill {
		Tag_a:schemeClr val phClr {
		  Tag_a:shade val 95000 {
		  }
		  Tag_a:satMod val 105000 {
		  }
		}
	      }
	      Tag_a:prstDash val solid {}
	    }
	    Tag_a:ln w 25400 cap flat cmpd sng algn ctr {
	      Tag_a:solidFill {
		Tag_a:schemeClr val phClr {}
	      }
	      Tag_a:prstDash val solid {}
	    }
	    Tag_a:ln w 38100 cap flat cmpd sng algn ctr {
	      Tag_a:solidFill {
		Tag_a:schemeClr val phClr {}
	      }
	      Tag_a:prstDash val solid {}
	    }
	  }
	  Tag_a:effectStyleLst {
	    Tag_a:effectStyle {
	      Tag_a:effectLst {
		Tag_a:outerShdw blurRad 40000 dist 20000 dir 5400000 rotWithShape 0 {
		  Tag_a:srgbClr val 000000 {
		    Tag_a:alpha val 38000 {}
		  }
		}
	      }
	    }
	    Tag_a:effectStyle {
	      Tag_a:effectLst {
		Tag_a:outerShdw blurRad 40000 dist 23000 dir 5400000 rotWithShape 0 {
		  Tag_a:srgbClr val 000000 {
		    Tag_a:alpha val 35000 {}
		  }
		}
	      }
	    }
	    Tag_a:effectStyle {
	      Tag_a:effectLst {
		Tag_a:outerShdw blurRad 40000 dist 23000 dir 5400000 rotWithShape 0 {
		  Tag_a:srgbClr val 000000 {
		    Tag_a:alpha val 35000 {}
		  }
		}
	      }
	      Tag_a:scene3d {
		Tag_a:camera prst orthographicFront {
		  Tag_a:rot lat 0 lon 0 rev 0 {
		  }
		}
		Tag_a:lightRig rig threePt dir t {
		  Tag_a:rot lat 0 lon 0 rev 1200000 {
		  }
		}
	      }
	      Tag_a:sp3d {
		Tag_a:bevelT w 63500 h 25400 {}
	      }
	    }
	  }
	  Tag_a:bgFillStyleLst {
	    Tag_a:solidFill {
	      Tag_a:schemeClr val phClr {}
	    }
	    Tag_a:gradFill rotWithShape 1 {
	      Tag_a:gsLst {
		Tag_a:gs pos 0 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 40000 {}
		    Tag_a:satMod val 350000 {}
		  }
		}
		Tag_a:gs pos 40000 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 45000 {}
		    Tag_a:shade val 99000 {}
		    Tag_a:satMod val 350000 {}
		  }
		}
		Tag_a:gs pos 100000 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:shade val 20000 {}
		    Tag_a:satMod val 255000 {}
		  }
		}
	      }
	      Tag_a:path path circle {
		Tag_a:fillToRect l 50000 t -80000 r 50000 b 180000 {}
	      }
	    }
	    Tag_a:gradFill rotWithShape 1 {
	      Tag_a:gsLst {
		Tag_a:gs pos 0 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:tint val 80000 {}
		    Tag_a:satMod val 300000 {}
		  }
		}
		Tag_a:gs pos 100000 {
		  Tag_a:schemeClr val phClr {
		    Tag_a:shade val 30000 {}
		    Tag_a:satMod val 200000 {}
		  }
		}
	      }
	      Tag_a:path path circle {
		Tag_a:fillToRect l 50000 t 50000 r 50000 b 50000 {}
	      }
	    }
	  }
	}
      }
      Tag_a:objectDefaults {
	Tag_a:spDef {
	  Tag_a:spPr {}
	  Tag_a:bodyPr {}
	  Tag_a:lstStyle {}
	  Tag_a:style {
	    Tag_a:lnRef idx 1 {
	      Tag_a:schemeClr val accent1 {}
	    }
	    Tag_a:fillRef idx 3 {
	      Tag_a:schemeClr val accent1 {}
	    }
	    Tag_a:effectRef idx 2 {
	      Tag_a:schemeClr val accent1 {}
	    }
	    Tag_a:fontRef idx minor {
	      Tag_a:schemeClr val lt1 {}
	    }
	  }
	}
	Tag_a:lnDef {
	  Tag_a:spPr {}
	  Tag_a:bodyPr {}
	  Tag_a:lstStyle {}
	  Tag_a:style {
	    Tag_a:lnRef idx 2 {
	      Tag_a:schemeClr val accent1 {}
	    }
	    Tag_a:fillRef idx 0 {
	      Tag_a:schemeClr val accent1 {}
	    }
	    Tag_a:effectRef idx 1 {
	      Tag_a:schemeClr val accent1 {}
	    }
	    Tag_a:fontRef idx minor {
	      Tag_a:schemeClr val tx1 {}
	    }
	  }
	}
      }
      Tag_a:extraClrSchemeLst {}
    }
    ::ooxml::Dom2zip $zf $root "xl/theme/theme1.xml" cd count
    $doc delete


    # ------------------------------ xl/workbook.xml ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(M) workbook]
    # set root [$doc documentElement]
    # $root setAttributeNS {} xmlns:r $xmlns(r)

    dom createDocument workbook doc
    $doc documentElement root

    $root setAttribute xmlns $xmlns(M)
    $root setAttribute xmlns:r $xmlns(r)

    $root appendFromScript {
      Tag_fileVersion appName xl lastEdited 5 lowestEdited 5 rupBuild 5000 {}
      Tag_workbookPr showInkAnnotation 0 autoCompressPictures 0 {}
      Tag_bookViews {
	set attr {}
	foreach {n v} [array get view] {
	  switch -- $n {
	    activetab {
	      if {$v ni $obj(sheets)} {
	        set v 0
	      }
	      lappend attr activeTab $v
	    }
	    x {
	      lappend attr xWindow $v
	    }
	    y {
	      lappend attr yWindow $v
	    }
	    height {
	      lappend attr windowHeight $v
	    }
	    width {
	      lappend attr windowWidth $v
	    }
	    default {
	    }
	  }
	}
	Tag_workbookView {*}$attr {}
      }
      Tag_sheets {
	for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
	  Tag_sheet name $obj(sheet,$ws) sheetId $ws r:id rId$ws {}
	}
      }
      if {0} {
	Tag_definedNames {
	  Tag_definedName name _xlnm._FilterDatabase localSheetId 0 hidden 1 { Text Blatt1!$A$1:$C$1 }
	}
      }
      Tag_calcPr calcId 140000 concurrentCalc 0 {}
      # fullCalcOnLoad 1
    }
    ::ooxml::Dom2zip $zf $root "xl/workbook.xml" cd count
    $doc delete


    # ------------------------------ xl/worksheets/sheet1.xml SHEET ------------------------------

    # In the commented out part, serialization is done with namespaces set correctly, but this takes a few milliseconds more time.
    # Since we may process large amounts of data and do not use the same tag names in different namespaces and do not search in
    # the DOM during serialization, we only set the namespaces as attributes, which leads to the same result in the serialized XML.
    # set doc [dom createDocumentNS $xmlns(M) worksheet]
    # set root [$doc documentElement]
    # $root setAttributeNS {} xmlns:r $xmlns(r)
    # $root setAttributeNS {} xmlns:mc $xmlns(mc)
    # $root setAttributeNS {} xmlns:x14ac $xmlns(x14ac)
    # $doc selectNodesNamespaces [list M $xmlns(M) r $xmlns(r) mc $xmlns(mc) ac $xmlns(x14ac)]

    for {set ws 1} {$ws <= $obj(sheets)} {incr ws} {
      dom createDocument worksheet doc
      $doc documentElement root

      $root setAttribute xmlns $xmlns(M)
      $root setAttribute xmlns:r $xmlns(r)
      $root setAttribute xmlns:mc $xmlns(mc)
      $root setAttribute xmlns:x14ac $xmlns(x14ac)

      $root setAttribute mc:Ignorable x14ac

      $root appendFromScript {
	Tag_dimension ref [::ooxml::RowColumnToString $obj(dminrow,$ws),$obj(dmincol,$ws)]:[::ooxml::RowColumnToString $obj(dmaxrow,$ws),$obj(dmaxcol,$ws)] {}
	Tag_sheetViews {
	  Tag_sheetView workbookViewId 0 {
	    if {$obj(freeze,$ws) ne {}} {
	      lassign [split [::ooxml::StringToRowColumn $obj(freeze,$ws)] ,] row col
	      Tag_pane xSplit $col ySplit $row topLeftCell $obj(freeze,$ws) state frozen {}
	    }
	  }
	}
	Tag_sheetFormatPr baseColWidth 10 defaultRowHeight 16 x14ac:dyDescent 0.2 {}
	if {[info exists obj($ws,cols)] && $obj($ws,cols) > 0} {
	  Tag_cols {}
	}
	Tag_sheetData {
	  array unset rows
	  foreach idx [lsort -dictionary [array names cells $ws,*,*]] {
	    lassign [split $idx ,] sheet row col
            lappend rows($row) $col
	  }
	  foreach row [lsort -integer [array names rows]] {
	    set attr {}
	    if {[dict exists $obj(rowHeight,$ws) $row]} {
	      lappend attr ht [dict get $obj(rowHeight,$ws) $row] customHeight 1
	    }
	    # lappend attr spans [expr {$minCol + 1}]:[expr {$maxCol + 1}]
	    Tag_row r [expr {$row + 1}] {*}$attr {
              foreach col $rows($row) {
                set idx "$ws,$row,$col"
		if {([dict exists $cells($idx) v] && [string trim [dict get $cells($idx) v]] ne {}) || ([dict exists $cells($idx) f] && [string trim [dict get $cells($idx) f]] ne {}) || [dict exists $cells($idx) fsi]} {
		  set attr {}
		  if {[dict exists $cells($idx) s] && [dict get $cells($idx) s] > 0} {
		    lappend attr s [dict get $cells($idx) s]
		  }
		  if {[dict exists $cells($idx) t] && [dict get $cells($idx) t] ne {n}} {
		    lappend attr t [dict get $cells($idx) t]
		  }
		  Tag_c r [::ooxml::RowColumnToString $row,$col] {*}$attr {
		    if {[dict exists $cells($idx) v] && [dict get $cells($idx) v] ne {}} {
		      Tag_v { Text [dict get $cells($idx) v] }
		    }
		    if {([dict exists $cells($idx) f] && [dict get $cells($idx) f] ne {}) || [dict exists $cells($idx) fsi]} {
		      ## FORMULA ##
		      set attr {}
		      if {[dict exists $cells($idx) fsi] && [dict get $cells($idx) fsi] > -1} {
			lappend attr t shared
		      }
		      if {[dict exists $cells($idx) fsr] && [dict get $cells($idx) fsr] ne {}} {
			lappend attr ref [dict get $cells($idx) fsr]
		      }
		      if {[dict exists $cells($idx) fsi] && [dict get $cells($idx) fsi] > -1} {
			lappend attr si [dict get $cells($idx) fsi]
		      }
		      Tag_f {*}$attr { 
			if {[dict exists $cells($idx) f] && [dict get $cells($idx) f] ne {}} {
			  Text [dict get $cells($idx) f]
			}
		      }
		    }
		  }
		} elseif {[dict exists $cells($idx) s] && [string is integer -strict [dict get $cells($idx) s]] && [dict get $cells($idx) s] > 0} {
		  Tag_c r [::ooxml::RowColumnToString $row,$col] s [dict get $cells($idx) s] {}
		}
		# garbage collection
	        unset -nocomplain cells($idx)
	      }
	    }
	  }
	}
	if {$obj(autofilter,$ws) ne {}} {
	  Tag_autoFilter ref $obj(autofilter,$ws) {}
	}
	if {[info exists obj(merge,$ws)] && $obj(merge,$ws) ne {}} {
	  Tag_mergeCells count [llength $obj(merge,$ws)] {
	    foreach item $obj(merge,$ws) {
	      Tag_mergeCell ref $item {}
	    }
	  }
	}
	Tag_pageMargins left 0.75 right 0.75 top 1 bottom 1 header 0.5 footer 0.5 {}
      }

      if {[set colsNode [$root selectNodes /worksheet/cols]] ne {}} {
	if {[info exists obj($ws,cols)] && $obj($ws,cols) > 0} {
	  $colsNode appendFromScript {
	    foreach idx [lsort -dictionary [array names cols $ws,*]] {
	      set attr {}
	      lappend attr min [expr {[dict get $cols($idx) min] + 1}] max [expr {[dict get $cols($idx) max] + 1}]
	      if {[dict get $cols($idx) width] ne {}} {
		lappend attr width [dict get $cols($idx) width]
		if {[dict get $cols($idx) width] != $::ooxml::defaults(cols,width)} {
		  dict set $cols($idx) customwidth 1
		}
	      }
	      if {[dict get $cols($idx) style] ne {} && [dict get $cols($idx) style] > 0} {
		lappend attr style [dict get $cols($idx) style]
	      }
	      if {[dict get $cols($idx) bestfit] == 1} {
		lappend attr bestFit [dict get $cols($idx) bestfit]
	      }
	      if {[dict get $cols($idx) customwidth] == 1} {
		lappend attr customWidth [dict get $cols($idx) customwidth]
	      }
	      Tag_col {*}$attr {}
	    }
	  }
	}
      }
      ::ooxml::Dom2zip $zf $root "xl/worksheets/sheet$ws.xml" cd count
      $doc delete
    }

    # Finalize zip.
    set cdoffset [tell $zf]
    set endrec [binary format a4ssssiis PK\05\06 0 0 $count $count [string length $cd] $cdoffset 0]
    puts -nonewline $zf $cd
    puts -nonewline $zf $endrec
    close $zf
    return 0
  }
}


#
# ooxml::tablelist_to_xl
#

proc ::ooxml::tablelist_to_xl { lb args } {
  variable defaults

  if {![winfo exists $lb]} {
    tk_messageBox -message [msgcat::mc {Tablelist does not exists!}]
    return
  }

  array set opts "
    callback ::ooxml::tablelist_to_xl_callback
    path [list $defaults(path)]
    file tablelist.xlsx
    creator unknown
    name Tablelist1
    rootonly 0
    addtimestamp 0
    globalstyle 0
  "

  set len [llength $args]
  set idx 0
  for {set idx 0} {$idx < $len} {incr idx} {
    switch -- [set opt [lindex $args $idx]] {
      -callback - -path - -file - -creator - -name {
	incr idx
	if {$idx < $len} {
	  set opts([string range $opt 1 end]) [lindex $args $idx]
	} else {
	  error "option '$opt': missing argument"
	}            
      }
      -rootonly - -addtimestamp - -globalstyle {
	set opts([string range $opt 1 end]) 1
      }
      default {
	error "unknown option \"$opt\", should be: -callback, -path, -file, -creator, -name, -rootonly, -addtimestamp or -globalstyle"
      }
    }
  }
  if {$opts(callback) eq {} || ([info commands $opts(callback)] eq {} && [info commands ::$opts(callback)] eq {})} {
    set opts(callback) ::ooxml::tablelist_to_xl_callback
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
  if {$opts(globalstyle)} {
    set globalstyle {-globalstyle}
  } else {
    set globalstyle {}
  }

  set file [tk_getSaveFile -confirmoverwrite 1 -filetypes {{{Excel Office Open XML} {.xlsx}}} -initialdir $opts(path) -initialfile $opts(file) -parent . -title "Excel Office Open XML"]
  if {$file eq {}} {
    tk_messageBox -message [msgcat::mc {No file selected!}]
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
    for {set col 0} {$col <= $columncount} {incr col} {
      set title [$lb columncget $col -title]
      set width [$lb columncget $col -width]
      set align [$lb columncget $col -align]
      set sortmode [$lb columncget $col -sortmode]
      set hide [$lb columncget $col -hide]

      $opts(callback) $spreadsheet $sheet $columncount $col $title $width $align $sortmode $hide

      $spreadsheet cell $sheet $title
    }

    if {$opts(rootonly)} {
      foreach row [$lb get [$lb childkeys root]] {
	$spreadsheet row $sheet
	set idx 0
	foreach col $row {
	  if {[string trim $col] ne {}} {
	    $spreadsheet cell $sheet $col -index $idx {*}$globalstyle
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
	    $spreadsheet cell $sheet $col -index $idx {*}$globalstyle
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

package provide ooxml 1.6

# Local Variables:
# tcl-indent-level: 2
# End:
