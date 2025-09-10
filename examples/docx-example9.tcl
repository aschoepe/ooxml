#!/bin/sh
#\
exec tclsh8.6 "$0" "$@"

source ../ooxml.tcl
source ../ooxml-docx.tcl
namespace import ::ooxml::docx::docx

set doc [docx new]

# Title
$doc paragraph "OMML"

# Inline: E = m c^2
$doc paragraph ""
$doc math {
    $doc mrun "E = m"
    $doc msup { $doc mrun "c" } { $doc mrun "2" }
}

# Display: Sum_{i=1}^{n} i^2
$doc math -display 1 -jc center {
    $doc mnary -char "∑" -limLoc undOvr \
      -sub { $doc mrun "i=1" } \
      -sup { $doc mrun "n" } {
        $doc msup { $doc mrun "i" } { $doc mrun "2" }
    } 
}

# Display: lim_{x->0} (sin x)/x
$doc math -display 1 {
    $doc mlimlow { $doc mrun "lim" } { $doc mrun "x→0" }
    $doc mrun " "
    $doc mfrac -type bar { $doc mfunc { $doc mrun "sin" } { $doc mrun "x" } } { $doc mrun "x" }
}

# Display: integral_0^π sin x dx
$doc math -display 1 {
    $doc mnary -char "∫" -limLoc subSup \
      -sub { $doc mrun "0" } \
      -sup { $doc mrun "π" } {
        $doc mrun "sin "
        $doc mrun "x "
        $doc mrun "dx"
    }
}

# Display: nth-root(sin x)
$doc math -display 1 {
    $doc mroot { $doc mrun "n" } { $doc mfunc { $doc mrun "sin" } { $doc mrun "x" } }
}

# Display: lim^{n→∞}
$doc math -display 1 {
    $doc mlimupp { $doc mrun "lim" } { $doc mrun "n→∞" }
}

$doc write docx-example9.docx
$doc destroy
