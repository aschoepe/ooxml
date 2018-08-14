#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

package require vfs::zip
package require tdom 0.9.0-

proc UnicodeCharsToHex { txt } {
  set new {}
  foreach {a b c d} [split [binary encode hex [encoding convertto unicode $txt]] {}] {
    if {[expr 0x${c}${d}${a}${b}] <= 0x7f} {
      append new [binary decode hex ${c}${d}${a}${b}]
    } else {
      append new \\u $c $d $a $b
    }
  }
  return [string map {\0 {}} $new]
}

foreach file [glob -directory languages *.xlsx] {
  set lang [lindex [split $file -.] 2]
  set files($lang) $file
}

set maptable {}

foreach lang [lsort [array names files]] {
  set file $files($lang)
  set mnt [vfs::zip::Mount $file xlsx]
  set name [lindex [split [UnicodeCharsToHex $file] -.] 0]

  set translation [string trim [file tail $name] 1]
  if {[regexp -all {\s+} $translation]} {
    set translation \"$translation\"
  }
  puts "msgcat::mcset $lang [string trim Book1 1] $translation"

  if {![catch {open xlsx/docProps/app.xml r} fd]} {
    fconfigure $fd -encoding utf-8
    if {![catch {dom parse [read $fd]} doc]} {
      set root [$doc documentElement]
      set idx -1
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI] vt http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes] {/X:Properties/X:HeadingPairs/vt:vector/vt:variant/vt:lpstr/text()}] {
	set translation [UnicodeCharsToHex [$node nodeValue]]
	if {[regexp -all {\s+} $translation]} {
	  set translation \"$translation\"
	}
	puts "msgcat::mcset $lang Worksheets $translation"
      }
      foreach node [$root selectNodes -namespaces [list X [$root namespaceURI] vt http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes] {/X:Properties/X:TitlesOfParts/vt:vector/vt:lpstr/text()}] {
	set translation [string trim [UnicodeCharsToHex [$node nodeValue]] 1]
	if {[regexp -all {\s+} $translation]} {
	  set translation \"$translation\"
	}
	puts "msgcat::mcset $lang [string trim Sheet1 1] $translation"
      }
      $doc delete
    }
    close $fd
  }

  vfs::zip::Unmount $mnt xlsx
}

