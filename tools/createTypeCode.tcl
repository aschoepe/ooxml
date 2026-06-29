#!/bin/sh
#\
exec tclsh "$0" "$@"

# This script reads a simple xsd enumeration type definition from
# stdin and generates a Tcl type check proc from it.
#
# It is intended to be called from emacs per shell-command-on-region
# (M-|) with an xsd type definition as region.

package require tdom

set result ""
set xml {<xsd:umbrella xmlns:xsd="http://www.w3.org/2001/XMLSchema">}
append xml [read stdin]
append xml "</xsd:umbrella>"
set doc [dom parse $xml]
$doc selectNodesNamespaces {xsd http://www.w3.org/2001/XMLSchema}

foreach simpleType [$doc selectNodes /xsd:umbrella/xsd:simpleType] {
    append result "proc ::ooxml::docx::lib::[$simpleType @name] {value} \{\n"
    append result "    set values \{\n"
    foreach enumeration [$simpleType selectNodes xsd:restriction/xsd:enumeration] {
        append result "        [$enumeration @value]\n"
    }
    append result "    \}\n"
    append result "    if {\$value in \$values} {
        return \$value
    }
    error \"unknown [$simpleType @name] value \\\"\$value\\\", expected one of\\
            \[AllowedValues \$values\]\"\n"
    append result "\}\n"
}
puts $result
