if 0 {
if {[catch {
    package require ooxml
} errMsg]} {
    if {[catch {
        source [file join [file dirname [info script]] ../ooxml.tcl]
    } err2]} {
        puts $err2
        error $errMsg
    }
}
}
source ../ooxml.tcl

proc doit {rows cols} {
    set spreadsheet [::ooxml::xl_write new -creator "ich"]
    set sheet [$spreadsheet worksheet "Meine Daten"]
    set row 0
    while {$row < $rows} {
        set col 0
        while {$col < $cols} {
            $spreadsheet cell $sheet "Cell $row,$col"
            incr col
        }
        $spreadsheet row $sheet
        incr row
    }
    $spreadsheet write "$rows-$cols.xlsx"
    $spreadsheet destroy
}

#    10 10
#    100 10
#    500 10
foreach {rows cols} {
    1000 10
    3000 10
    5000 10
} {
    set msec [lindex [time {doit $rows $cols}] 0]
    set nrcells [expr {$rows * $cols}]
    set perCell [expr {$msec / $nrcells}]
    set seconds [format "%.3f" [expr {double($msec)/1000000}]]
    puts "$rows/$cols: $nrcells cells, $seconds seconds, $perCell per cell"
}
