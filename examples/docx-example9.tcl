#!/bin/sh
#\
exec tclsh "$0" "$@"

cd [file dirname [info script]]
source ../ooxml.tcl
source ../ooxml-docx.tcl
source strings.tcl
namespace import ::ooxml::docx::docx

set docx [docx new]

featuresCovered "OMML related methods." 

# Title
$docx paragraph "OMML"

# Inline: E = m c^2
$docx paragraph ""
$docx math {
    $docx mrun "E = m"
    $docx msup { $docx mrun "c" } { $docx mrun "2" }
}

# Display: Sum_{i=1}^{n} i^2
$docx math -display 1 -jc center {
    $docx mnary -char "∑" -limLoc undOvr \
      -sub { $docx mrun "i=1" } \
      -sup { $docx mrun "n" } {
        $docx msup { $docx mrun "i" } { $docx mrun "2" }
    } 
}

# Display: lim_{x->0} (sin x)/x
$docx math -display 1 {
    $docx mlimlow { $docx mrun "lim" } { $docx mrun "x→0" }
    $docx mrun " "
    $docx mfrac -type bar { $docx mfunc { $docx mrun "sin" } { $docx mrun "x" } } { $docx mrun "x" }
}

# Display: integral_0^π sin x dx
$docx math -display 1 {
    $docx mnary -char "∫" -limLoc subSup \
      -sub { $docx mrun "0" } \
      -sup { $docx mrun "π" } {
        $docx mrun "sin "
        $docx mrun "x "
        $docx mrun "dx"
    }
}

# Display: nth-root(sin x)
$docx math -display 1 {
    $docx mroot { $docx mrun "n" } { $docx mfunc { $docx mrun "sin" } { $docx mrun "x" } }
}

# Display: lim^{n→∞}
$docx math -display 1 {
    $docx mlimupp { $docx mrun "lim" } { $docx mrun "n→∞" }
}

$docx write docx-example9.docx
$docx destroy
