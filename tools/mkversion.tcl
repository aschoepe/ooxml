#!/bin/sh
#\
exec tclsh "$0" "$@"

#
# 2024 Alexander Schöpe, Bochum, DE
# mkversion.c from fossil-scm ported to vanilla tcl
#

proc hash {zIn N} {
  set zL {}
  foreach c [split $zIn {}] {
    lappend zL [scan $c %c]
  }
  lappend zL 0
  set zIn $zL
  unset zL

  set s {}
  for {set m 0} {$m < 256} {incr m} {
    lappend s $m
  }
  lappend s 0

  for {set t 0; set j 0; set n 0; set m 0} {$m < 256} {incr m; incr n} {
    set j [expr {($j + [lindex $s $m] + [lindex $zIn $n]) & 0xff}]
    if {[lindex $zIn $n] == 0} {
      set n -1
    }
    set t [lindex $s $j]
    set s [lreplace $s $j $j [lindex $s $m]]
    set s [lreplace $s $m $m $t]
  }

  set i 0
  set j 0
  set zOut {}
  for {set n 0} {$n < $N-2} {incr n 2} {
    incr i
    set t [lindex $s $i]
    set j [expr {($j + $t) & 0xff}]
    set s [lreplace $s $i $i [lindex $s $j]]
    set s [lreplace $s $j $j $t]
    set t [expr {($t + [lindex $s $i]) & 0xff}]
    append zOut [string index 0123456789abcdef [expr {($t >> 4) & 0xf}]]
    append zOut [string index 0123456789abcdef [expr {$t & 0xf}]]
  }

  return $zOut
}

set manifest {}

if {![catch {open manifest.uuid r} fd]} {
  if {![eof $fd]} {
    if {[gets $fd uuid] > 0} {
      lappend manifest MANIFEST_UUID $uuid
      lappend manifest MANIFEST_VERSION [format {[%10.10s]} $uuid]
      if {[info exists ::env(SOURCE_DATE_EPOCH)] && [string is wideinteger $::env(SOURCE_DATE_EPOCH)]} {
        set ctime $::env(SOURCE_DATE_EPOCH)
      } else {
        set ctime [clock seconds]
      }
      append uuid $ctime
      lappend manifest FOSSIL_BUILD_HASH [hash $uuid 33]
    }
  }
  close $fd
} else {
  puts stderr "can't open file manifest.uuid"
}

if {![catch {open manifest r} fd]} {
  while {![eof $fd]} {
    if {[gets $fd line] > 0} {
      if {[string match {D *} $line]} {
	set ctime [clock scan [lindex [split [lindex $line 1] .] 0] -gmt 1]
	lappend manifest MANIFEST_DATE [clock format $ctime -format {%Y-%m-%d %H:%M:%S} -gmt 1]
	lappend manifest MANIFEST_YEAR [clock format $ctime -format {%Y} -gmt 1]
	lappend manifest MANIFEST_NUMERIC_DATE [clock format $ctime -format {%Y%m%d} -gmt 1]
	lappend manifest MANIFEST_NUMERIC_TIME [clock format $ctime -format {%H%M%S} -gmt 1]
	break
      }
    }
  }
  close $fd
} else {
  puts stderr "can't open file manifest"
}

if {![catch {open configure.ac r} fd]} {
  while {![eof $fd]} {
    if {[gets $fd line] > 0} {
      if {[string match {AC_INIT*} $line]} {
	if {[regexp {AC_INIT\(\[.*?\],\[(.*?)\]\)} $line match version]} {
	  set versionList [lrange [split $version.0.0.0 .] 0 3]
	  lappend manifest RELEASE_VERSION $version
	  lappend manifest RELEASE_VERSION_NUMBER [format {%d%02d%d%d} {*}$versionList]
	  lappend manifest RELEASE_RESOURCE_VERSION [format {%d,%d,%d,%d} {*}$versionList]
	}
	break
      }
    }
  }
  close $fd
} else {
  puts stderr "can't open file manifest"
}

set template {namespace eval ::PKG {

  # TIP 599: Extended build information
  # https://core.tcl-lang.org/tips/doc/trunk/tip/599.md

  proc build-info { {cmd {}} } {
    set uuid MANIFEST_UUID
    set checkin [string map {[ {} ] {}} {MANIFEST_VERSION}]
    set build FOSSIL_BUILD_HASH
    set datetime [string map {{ } T} {MANIFEST_DATE}]Z
    set version RELEASE_VERSION
    set compiler {tcl.noarch}

    switch -- $cmd {
      commit {
        return $uuid
      }
      version - patchlevel {
        return $version
      }
      compiler {
        return $compiler
      }
      default {
        return ${version}+${checkin}.${datetime}.${compiler}
      }
    }
  }
}}

if {[llength $argv] == 1} {
  set pkgName [lindex $argv 0]
  lappend manifest PKG $pkgName
  puts [string map $manifest $template]

} else {
  foreach {n v} $manifest {
    puts [list $n $v]
  }
}