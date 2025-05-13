#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2025 Rolf Ade, DE
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

package require Tcl 8.6.7-
package require tdom 0.9.6-

namespace eval ::ooxml::docx {
    # To ensure it exists
}

namespace eval ::ooxml::docx::lib {

    namespace export docx OptVal NoCheck CT_* ST_* W3CDTF AllowedValues

}

# Method option handling helper procs and value checks
proc ::ooxml::docx::lib::AllowedValues {values {word "or"}} {
    return "[join [lrange $values 0 end-1] ", "] $word [lindex $values end]"
}

proc ::ooxml::docx::lib::OptVal {arglist {prefix ""}} {
    if {[llength $arglist] % 2 != 0} {
        if {$prefix ne ""} {append prefix " "}
        error "invalid arguments: expected ${prefix}?-option value ?-option value? ..?"
    }
    uplevel "array set opts [list $arglist]"
}

proc ::ooxml::docx::lib::NoCheck {value} {
    return $value
}

proc ::ooxml::docx::lib::ST_AlignH {value} {
    set values {
        left
        right
        center
        inside
        outside
    }
    if {$value in $values} {
        return $value
    }
    error "unknown alignH value \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_AlignV {value} {
    set values {
        top
        bottom
        center
        inside
        outside
    }
    if {$value in $values} {
        return $value
    }
    error "unknown alignV value \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_Anchor {value} {
    set values {
        text
        margin
        page
    }
    if {$value in $values} {
        return $value
    }
    error "unknown vAnchor or hAnchor type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_BlackWhiteMode {value} {
    set values {
        auto
        black
        blackGray
        blackWhite
        clr
        gray
        grayWhite
        hidden
        invGray
        ltGray
        white
    }
    if {$value in $values} {
        return $value
    }
    error "unknown back and white mode \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_Border {value} {
    # The xsd list a lot more but they seemed not to be supported
    # anymore by recent docx renderer.
    set values {
        nil
        none
        single
        thick
        double
        dotted
        dashed
        dotDash
        dotDotDash
        triple
        thinThickSmallGap
        thickThinSmallGap
        thinThickThinSmallGap
        thinThickMediumGap
        thickThinMediumGap
        thinThickThinMediumGap
        thinThickLargeGap
        thickThinLargeGap
        thinThickThinLargeGap
        wave
        doubleWave
        dashSmallGap
        dashDotStroked
        threeDEmboss
        threeDEngrave
        outset
        inset
    }
    if {$value in $values} {
        return $value
    }
    error "unknown border type value \"$value\", expected one of:\
               [AllowedValues $values]"
}
 
proc ::ooxml::docx::lib::ST_ChapterSep {value} {
    set values {
        hyphen
        period
        colon
        emDash
        enDash
    }
    if {$value in $values} {
        return $value
    }
    error "unknown section separator character \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_DecimalNumber {value} {
    if {![string is integer -strict $value]} {
        error "expected integer but got \"$value\""
    }
    return $value
}

proc ::ooxml::docx::lib::ST_DropCap {value} {
    set values {
        none
        drop
        margin
    }
    if {$value in $values} {
        return $value
    }
    error "unknown dropCap type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_EighthPointMeasure {value} {
    if {[string is integer -strict $value] && $value >= 0} {
        return $value
    }
    error "\"$value\" is not a valid measure value - value must be an\
               integer"
}

proc ::ooxml::docx::lib::ST_Emu {value} {
    if {[string is integer -strict $value] && $value >= 0} {
        return $value
    }
    if {![regexp {[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
        error "\"$value\" is not a valid measure value - value must be an\
                   integer or match the regular expression\
                   \[0-9\]+(\.\[0-9\]+)?(mm|cm|in|pt|pc|pi)"
    }
    scan $value "%f%s" value unit
    switch $unit {
        mm {set factor 36000}
        cm {set factor 360000}
        in {set factor 914400}
        pt {set factor 12700}
        pc {set factor 152400}
        pi {set factor 152400}
    }
    return [expr {round($value*$factor)}]
}

proc ::ooxml::docx::lib::ST_HeightRule {value} {
    set values {
        auto
        exact
        atLeast
    }
    if {$value in $values} {
        return $value
    }
    error "unknown height rule value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_HexColor {value} {
    if {$value ne "auto"} {
        return $value
    }
    if {[string length $value] != 6 || ![string is xdigit]} {
        error "unknown color value \"$value\", should be \"auto\" or a hex value in\
                   RRGGBB format."
    }
    return $value
}

proc ::ooxml::docx::lib::ST_HighlightColor {value} {
    set values {
        black
        blue
        cyan
        green
        magenta
        red
        yellow
        white
        darkBlue
        darkCyan
        darkGreen
        darkMagenta
        darkRed
        darkYellow
        darkGray
        lightGray
        none
    }
    if {$value in $values} {
        return $value
    }
    error "unknown highlight color value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_Jc {value} {
    set values {
        start
        center
        end
        both
        mediumKashida
        distribute
        numTab
        highKashida
        lowKashida
        thaiDistribute
        left
        right
    }
    if {$value in $values} {
        return $value
    }
    error "unknown align value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_JcTable {value} {
    set values {
        center
        end
        left
        right
        start
    }
    if {$value in $values} {
        return $value
    }
    error "unknown table align value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::CT_OnOff {value} {
    if {![string is boolean -strict $value]} {
        error "expected a Tcl boolean value"
    }
    if {$value} {
        return "on"
    } else {
        return "off"
    }
}

proc ::ooxml::docx::lib::ST_MeasurementOrPercent {value} {
    if {[string is integer -strict $value]} {
        return $value
    }
    if {[regexp -- {-?[0-9]+(\.[0-9]+)?%} $value]} {
        return $value
    }
    if {![regexp -- {-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
        error "\"$value\" is not a valid measure or perent value - value must\
              be an integer or a numeric value directly followed either a\
              percent sign (%) or one of the units mm, cm, in, pt, pc or pi"
    }
    return $value

}

proc ::ooxml::docx::lib::ST_Merge {value} {
    set values {
        continue
        restart
    }
    if {$value in $values} {
        return $value
    }
    error "unknown merge value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_NumberFormat {value} {
    set values {
        decimal
        upperRoman
        lowerRoman
        upperLetter
        lowerLetter
        ordinal
        cardinalText
        ordinalText
        hex
        chicago
        ideographDigital
        japaneseCounting
        aiueo
        iroha
        decimalFullWidth
        decimalHalfWidth
        japaneseLegal
        japaneseDigitalTenThousand
        decimalEnclosedCircle
        decimalFullWidth2
        aiueoFullWidth
        irohaFullWidth
        decimalZero
        bullet
        ganada
        chosung
        decimalEnclosedFullstop
        decimalEnclosedParen
        decimalEnclosedCircleChinese
        ideographEnclosedCircle
        ideographTraditional
        ideographZodiac
        ideographZodiacTraditional
        taiwaneseCounting
        ideographLegalTraditional
        taiwaneseCountingThousand
        taiwaneseDigital
        chineseCounting
        chineseLegalSimplified
        chineseCountingThousand
        koreanDigital
        koreanCounting
        koreanLegal
        koreanDigital2
        vietnameseCounting
        russianLower
        russianUpper
        none
        numberInDash
        hebrew1
        hebrew2
        arabicAlpha
        arabicAbjad
        hindiVowels
        hindiConsonants
        hindiNumbers
        hindiCounting
        thaiLetters
        thaiNumbers
        thaiCounting
        bahtText
        dollarText
        custom
    }
    if {$value in $values} {
        return $value
    }
    error "unknown numbering format \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_PageOrientation {value} {
    if {$value ni {landscape portrait}} {
        error "unknown symbol \"$value\", expected \"landscape\"\
                   or \"portrait\""
    }
    return $value
}

proc ::ooxml::docx::lib::ST_PointMeasure {value} {
    if {[string is integer -strict $value] && $value >= 0} {
        return $value
    }
    error "\"$value\" is not a valid measure value - value must be an\
               integer"
}

proc ::ooxml::docx::lib::ST_RelFromH {value} {
    set values {
        margin
        page
        column
        character
        leftMargin
        rightMargin
        insideMargin
        outsideMargin
    }
    if {$value in $values} {
        return $value
    }
    error "unknown horizonal relative from value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_RelFromV {value} {
    set values {
        margin
        page
        paragraph
        line
        topMargin
        bottomMargin
        insideMargin
        outsideMargin
    }
    if {$value in $values} {
        return $value
    }
    error "unknown vertical relative from value \"$value\", expected one of:\
               [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_SignedTwipsMeasure {value} {
    if {[string is integer -strict $value]} {
        return $value
    }
    if {![regexp -- {-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
        error "\"$value\" is not a valid measure value - value must match\
               the regular expression \[0-9\]+(\.\[0-9\]+)?(mm|cm|in|pt|pc|pi)"
    }
    return $value
}

proc ::ooxml::docx::lib::ST_TabJc {value} {
    set values {
        clear
        start
        center
        end
        decimal
        bar
        num
        left
        right
    }
    if {$value in $values} {
        return $value
    }
    error "unknown tab stop type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_TabTlc {value} {
    set values {
        none
        dot
        hyphen
        underscore
        heavy
        middleDot
    }
    if {$value in $values} {
        return $value
    }
    error "unknown tab fill type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_TblLayoutType {value} {
    set values {
        fixed
        autofit
    }
    if {$value in $values} {
        return $value
    }
    error "unknown table layout type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_TblWidth {value} {
    set values {
        nil
        pct
        percent
        dxa
        measure
        auto
    }
    if {$value in $values} {
        switch $value {
            "percent" {return pct}
            "measure" {return dxa}
            default {return $value}
        }
    }
    error "unknown table width value type \"$value\", expected one of\
            [AllowedValues $values]"
}

# ST_HpsMeasure accepts exactly the same value as ST_TwipsMeasure.
# The difference is only the interpretation of the integer (only)
# values. For ST_HpsMeasure the number specifies half points
# (1/144 of an inch), for ST_TwipsMeasure the number specifies
# twentieths of a point (equivalent to 1/1440th of an inch).
proc ::ooxml::docx::lib::ST_TwipsMeasure {value} {
    if {[string is integer -strict $value] && $value >= 0} {
        return $value
    }
    if {![regexp {[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
        error "\"$value\" is not a valid measure value - value must match\
               the regular expression \[0-9\]+(\.\[0-9\]+)?(mm|cm|in|pt|pc|pi)"
    }
    return $value
}

proc ::ooxml::docx::lib::ST_Underline {value} {
    set values {
        single
        words
        double
        thick
        dotted
        dottedHeavy
        dash
        dashedHeavy
        dashLong
        dashLongHeavy
        dotDash
        dashDotHeavy
        dotDotDash
        dashDotDotHeavy
        wave
        wavyHeavy
        wavyDouble
        none
    }
    if {$value ni $values} {
        error "unkown underline value \"$value\", expected one of\
                  [AllowedValues $values]"
    }
    return $value
}

proc ::ooxml::docx::lib::ST_Wrap {value} {
    set values {
        auto
        notBeside
        around
        tight
        through
        none
    }
    if {$value in $values} {
        return $value
    }
    error "unknown wrap type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_WrapText {value} {
    set values {
        bothSides
        left
        right
        largest
    }
    if {$value in $values} {
        return $value
    }
    error "unknown wrapText value \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_XAlign {value} {
    set values {
        left
        center
        right
        inside
        outside
    }
    if {$value in $values} {
        return $value
    }
    error "unknown xAlign type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::ST_YAlign {value} {
    set values {
        inline
        top
        center
        bottom
        inside
        outside
    }
    if {$value in $values} {
        return $value
    }
    error "unknown yAlign type \"$value\", expected one of\
            [AllowedValues $values]"
}

proc ::ooxml::docx::lib::W3CDTF {value} {
    if {[catch {
        set value [clock format [clock scan $value] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
    }]} {
        error "invalid datetime value \"$value\", expected everything which is\
                accepted by \[clock scan\]"
    }
    return $value
}

package provide ::ooxml::docx::lib 1.8.1
