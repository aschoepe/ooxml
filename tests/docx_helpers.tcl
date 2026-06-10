# helpers.tcl --
#
#   Shared test utilities for ooxml::docx tests.

package require tcltest 2.5

namespace eval ::docxtest {
    variable tmpdir [file join [::tcltest::configure -tmpdir] docxtest_tmp]
    # Remove stale files from previous runs so the ecma-99.1 sweep
    # only validates documents produced by the current test run.
    if {[file isdirectory $tmpdir]} {
        file delete -force $tmpdir
    }
    file mkdir $tmpdir
    variable counter 0

    namespace export newdoc tmpfile xpath countNodes haspart \
                     writedoc ziplist zipread xmlparse cleanup withDoc \
                     writeTextFile writeTinyPng
}

proc ::docxtest::newdoc {args} {
    return [::ooxml::docx::docx new {*}$args]
}

proc ::docxtest::tmpfile {{suffix .docx}} {
    variable tmpdir
    variable counter
    incr counter
    return [file join $tmpdir "test_${counter}${suffix}"]
}

proc ::docxtest::xpath {doc xpath {part word/document.xml}} {
    return [$doc xpath $xpath $part]
}

proc ::docxtest::countNodes {doc xpath {part word/document.xml}} {
    return [expr {int([$doc xpath "count($xpath)" $part])}]
}

proc ::docxtest::haspart {doc part} {
    return [expr {$part in [$doc xmlparts]}]
}

proc ::docxtest::writedoc {doc} {
    set f [tmpfile]
    $doc write $f
    return $f
}

proc ::docxtest::ziplist {file} {
    set result [list]
    if {![catch {set data [exec unzip -Z1 $file]}]} {
        foreach entry [split $data \n] {
            set entry [string trim $entry]
            if {$entry ne ""} {
                lappend result $entry
            }
        }
    } else {
        set fd [open "|unzip -l [list $file]" r]
        while {[gets $fd line] >= 0} {
            if {[regexp {\d+\s+\d[\d-]+\s+\d{2}:\d{2}\s+(.+)$} $line -> name]} {
                lappend result [string trim $name]
            }
        }
        catch {close $fd}
    }
    return $result
}

proc ::docxtest::zipread {file entry} {
    # Escape glob-special chars so unzip doesn't treat them as wildcards
    regsub -all {\[} $entry {\\[} entry
    regsub -all {\]} $entry {\\]} entry
    return [exec unzip -p $file $entry]
}

proc ::docxtest::xmlparse {xml} {
    return [dom parse $xml]
}

proc ::docxtest::writeTextFile {contents {suffix .xml}} {
    set f [tmpfile $suffix]
    set fd [open $f w]
    try {
        fconfigure $fd -encoding utf-8
        puts -nonewline $fd $contents
    } finally {
        close $fd
    }
    return $f
}

proc ::docxtest::writeTinyPng {} {
    set f [tmpfile .png]
    set fd [open $f wb]
    try {
        # Tcl 9 removed the "binary" encoding alias. Opening the channel in
        # binary mode is sufficient and keeps the helper compatible with both
        # Tcl 8.6 and Tcl 9.
        puts -nonewline $fd [binary decode base64 \
            {iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO0PpX0AAAAASUVORK5CYII=}]
    } finally {
        close $fd
    }
    return $f
}

proc ::docxtest::cleanup {doc} {
    catch {$doc destroy}
}

proc ::docxtest::withDoc {varname body args} {
    upvar 1 $varname doc
    set doc [newdoc {*}$args]
    set code [catch {uplevel 1 $body} result opts]
    catch {$doc destroy}
    return -options $opts $result
}

namespace eval :: {
    namespace import -force ::docxtest::*
}
