
# This tool unzips every zip archive into a directory and pretty
# prints the XML files in this new directory hierarchy. If called
# just with a zip archive (mainly meat for open office documents) the
# directory to unpack the content is the base-name of the file name of
# the archive. The second argument overwrites this default with an
# directory name. If the directory exists, its content will be
# removed!

package require Tcl 9.0-
package require tdom
package require fileutil::traverse

if {$argc < 1 || $argc > 2} {
    puts stderr "usage: $argv0 <xslx-file> ?<dirname>?"
}

proc unzip {zipfile dir} {
    file mkdir $dir
    set base [file join [zipfs root] pp]
    set baselen [string length $base]
    zipfs mount $zipfile $base
    ::fileutil::traverse filesInZip $base
    filesInZip foreach file {
        set target [file join $dir [string range $file $baselen+1 end]]
        if {[file isdirectory $file]} {
            file mkdir $target
            continue
        }
        set in [open $file rb]
        set out [open $target wb]
        fcopy $in $out
        close $in
        close $out
        #puts [file join $dir $target]
        #puts $target
    }
}

proc pp {xmlfile} {
    if {[file isdirectory $xmlfile]} return
    if {[catch {set data [::tdom::xmlReadFile $xmlfile]} errMsg]} {
        puts "cannot read $xmlfile (probably a binary file) - $errMsg"
        return 
    }
    if {[catch {dom parse $data doc}]} {
        puts "$xmlfile failed to parse"
        return
    }
    set fd [open $xmlfile w+]
    $doc asXML -channel $fd
    close $fd
}

if {$argc == 1} {
    set file [lindex $argv 0]
    set basedir [file dirname $file]
    set name [file rootname [file tail $file]]
    set dir [file join $basedir $name]
} else {
    set dir [lindex $argv 1]
}
if {[file exists $dir]} {
    file delete -force $dir
}
file mkdir $dir
unzip $file $dir
::fileutil::traverse xmlfiles $dir
xmlfiles foreach xmlfile {
    pp $xmlfile
}
