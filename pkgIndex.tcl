if {![package vsatisfies [package provide Tcl] 8.6]} return
package ifneeded ooxml 1.3 [list source [file join $dir ooxml.tcl]]
