#!/bin/sh
#\
exec tclsh "$0" "$@"

# Copyright (C) 2019 Alexander Schoepe, Bochum, DE

package require Tcl 8.6-

switch -- [set cmd [lindex $argv 0]] {
  cp {
    set src [lrange $argv 1 end-1]
    set dst [lindex $argv end]
    if {![file exists $dst]} {
      file mkdir $dst
    }
    if {[file exists $dst] && [file isdirectory $dst]} {
      file copy -force -- {*}$src $dst
    }
  }
  default {
    exit 1
  }
}
