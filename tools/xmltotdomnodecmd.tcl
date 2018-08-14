#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

#
#  xmltotdomnodecmd.tcl <filename>
#  read xml file and build a source code to generate the original xml file
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


package require tdom 0.9.0-

set file $argv
set cmds {}

if {[catch {open $file r} fd]} {
  puts stderr "open: $fd"
  return
}

fconfigure $fd -encoding utf-8

if {[catch {dom parse [read $fd]} doc]} {
  puts stderr "dom parse: $doc"
  return
}

proc Preparse { node {level 0} } {
  global cmds

  if {$level != 0} {
    switch -- [$node nodeType] {
      ELEMENT_NODE {
        set name [$node nodeName]
      }
      TEXT_NODE {
        set name @textNode
      }
      CDATA_SECTION_NODE {
        set name @cdataNode
      }
      COMMENT_NODE {
        set name @commentNode
      }
      PROCESSING_INSTRUCTION_NODE {
        set name @piNode
      }
    }
    if {$name ni $cmds} {
      lappend cmds $name
    }
  }
  if {[$node hasChildNodes]} {
    foreach child [$node childNodes] {
      Preparse $child [expr {$level + 1}]
    }
  }
}

proc Parse { node {level 0} } {
  global cmds

  set brace 0
  set indent [string repeat { } [expr {2 * $level}]]
  set attributes {}
  foreach attr [$node attributes] {
    set attr0 [lindex $attr 0]
    set attr1 [lindex $attr 1]
    if {$attr0 ne {} && $attr1 ne {}} {
      if {$attr0 eq $attr1} {
	set attr xmlns:$attr0
      } else {
	set attr $attr1:$attr0
      }
    } else {
      set attr $attr0
    }
    lappend attributes $attr [$node getAttribute $attr]
  }
  if {$level == 0} {
    puts "set root \[\$doc documentElement\]"
    foreach {n v} $attributes {
      puts "\$root setAttribute [list $n] [list $v]"
    }
    puts "\$root appendFromScript \{"
  } else {
    switch -- [$node nodeType] {
      ELEMENT_NODE {
	set cmd "Tag_[$node nodeName] $attributes"
	puts "${indent}[string trim $cmd] \{"
	set brace 1
      }
      TEXT_NODE {
	puts "${indent}Text [$node nodeValue]"
      }
      CDATA_SECTION_NODE {
	puts "${indent}CData [$node nodeValue]"
      }
      COMMENT_NODE {
	puts "${indent}Comment [$node nodeValue]"
      }
      PROCESSING_INSTRUCTION_NODE {
	puts "${indent}PInstr [$node nodeValue]"
      }
    }
  }
  if {[$node hasChildNodes]} {
    foreach child [$node childNodes] {
      Parse $child [expr {$level + 1}]
    }
  }
  if {$brace} {
    puts "${indent}\}"
  }
  if {$level == 0} {
    puts "\}"
  }
}

set root [$doc documentElement]
puts "package require tdom 0.9.0-"
puts "set encoding utf-8"
puts "dom setResultEncoding \$encoding"
Preparse $root
foreach cmd [lsort $cmds] {
  switch -- $cmd {
    @textNode {
      puts "dom createNodeCmd textNode Text"
    }
    @cdataNode {
      puts "dom createNodeCmd cdataNode CData"
    }
    @commentNode {
      puts "dom createNodeCmd commentNode Comment"
    }
    @piNode {
      puts "dom createNodeCmd piNode PInstr"
    }
    default {
      puts "dom createNodeCmd -tagName $cmd elementNode Tag_$cmd"
    }
  }
}
puts "set doc \[dom createDocument [$root nodeName]\]"
Parse $root
puts "fconfigure stdout -encoding \$encoding"
puts "puts \[\$root asXML -indent 2 -xmlDeclaration 1 -encString \[string toupper \$encoding\]\]"
$doc delete
