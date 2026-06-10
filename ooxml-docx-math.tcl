#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2025 Rolf Ade, DE
#  Copyright (C) 2025, 2026 Miguel Bañón
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

namespace eval ::ooxml::docx {
    
    # ── Fraction properties ──────────────────────────────────────────
    set properties(fPr) {
        -type {m:type ST_FType}
    }

    # ── Math run properties ──────────────────────────────────────────
    set properties(mrun1) {
        -lit {m:lit CT_OnOff}
        | {
            {
                -scr {m:scr ST_MScript}
                -sty {m:sty ST_MathStyle}
            }
            {
                -nor {m:nor CT_OnOff}
            }
        }
    }

    set properties(mrun2) {
        -aln {m:aln CT_OnOff}
    }
    
    # ── N-ary operator properties ────────────────────────────────────
    set properties(naryPr) {
        -char {m:chr NoCheck}
        -limLoc {m:limLoc ST_LimLoc}
        -grow {m:grow CT_OnOff}
        -subHide {m:subHide CT_OnOff}
        -supHide {m:supHide CT_OnOff}
    }

    # ── Accent properties (Phase 1) ─────────────────────────────────
    set properties(accPr) {
        -char {m:chr NoCheck}
    }

    # ── Bar properties (Phase 1) ─────────────────────────────────────
    set properties(barPr) {
        -pos {m:pos ST_TopBot}
    }

    # ── Delimiter properties (Phase 1) ───────────────────────────────
    set properties(dPr) {
        -begChr {m:begChr NoCheck}
        -endChr {m:endChr NoCheck}
        -sepChr {m:sepChr NoCheck}
        -grow   {m:grow   CT_OnOff}
        -shp    {m:shp    ST_Shp}
    }

    # ── Grouping character properties (Phase 1) ──────────────────────
    set properties(groupChrPr) {
        -char   {m:chr    NoCheck}
        -pos    {m:pos    ST_TopBot}
        -vertJc {m:vertJc ST_TopBot}
    }

    # ── Matrix column properties (Phase 1) ───────────────────────────
    set properties(mcPr) {
        -count {m:count ST_Integer255}
        -mcJc  {m:mcJc  ST_McJc}
    }

    # ── Tag registration ─────────────────────────────────────────────
    # Original tags
    foreach tag {
        m:acc m:accPr m:aln
        m:bar m:barPr m:baseJc m:begChr m:brk
        m:chr m:count
        m:d m:dPr m:deg m:degHide m:den
        m:e m:endChr
        m:f m:fName m:fPr m:func
        m:groupChr m:groupChrPr m:grow
        m:jc
        m:lim m:limLoc m:limLow m:limUpp m:lit
        m:m m:mPr m:mc m:mcJc m:mcPr m:mcs
        m:mr m:nary m:naryPr m:nor m:num
        m:oMath m:oMathPara m:oMathParaPr
        m:pos
        m:r m:rPr m:rad m:radPr
        m:sSub m:sSubSup m:sSup m:scr m:sepChr m:shp m:sty m:sub m:subHide m:sup m:supHide
        m:t m:type
        m:vertJc
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(m) elementNode Tag_$tag
    }

}

# OMML related methods, the initial code contributed by Miguel Bañón
oo::define ooxml::docx::docx {

    # Guard: math sub-methods must be called from within a math zone
    method RequireMathZone {} {
        my variable inMathZone
        set methodName [lindex [info level [expr {[info level] - 1}]] 1]
        if {![info exists inMathZone] || !$inMathZone} {
            error "$methodName must be called from within a math zone\
                   (inside the script argument of the math method)"
        }
    }

    method EvalChildScript {tag script} {
        my variable body
        set savedbody $body
        Tag_$tag {
            set body [dom fromScriptContext]
            if {$script ne ""
                && [catch {uplevel 2 [list eval $script]} errMsg]} {
                set body $savedbody
                error $errMsg
            }
        }
        set body $savedbody
    }
    
    method Mt {text} {
        set atts ""
        if {[string index $text 0] eq " " || [string index $text end] eq " "} {
            lappend atts xml:space preserve
        }
        Tag_m:t $atts {
            Text [dom clearString -replace $text]
        }
    }

    # Inline or display math zone.
    #   doc math { ...OMML... }            ;# inline (m:oMath) in current paragraph
    #   doc math -display 1 { ...OMML... } ;# display (m:oMathPara/m:oMath) on its own paragraph
    # Options:
    #   -display CT_OnOff   (default off)
    #   -jc      ST_MJc     (left|center|right|centerGroup) - only used for display
    method math {args} {
        my variable body
        my variable valPrefix
        my variable inMathZone

        set savedValPrefix $valPrefix
        set savedInMathZone [expr {[info exists inMathZone] && $inMathZone}]
        set valPrefix m
        set inMathZone 1
        if {[catch {
            set script [lindex $args end]
            OptVal [lrange $args 0 end-1] "math" "script"

            set display [my EatOption -display CT_OnOff]   ;# "1" or "0"
            set jc      [my EatOption -jc ST_MJc]          ;# left|center|right|centerGroup
            my CheckRemainingOpts

            if {$display eq "1"} {
                $body appendFromScript {
                    Tag_w:p {
                        Tag_m:oMathPara {
                            if {$jc ne ""} {
                                Tag_m:oMathParaPr { Tag_m:jc m:val $jc }
                            }
                            my EvalChildScript m:oMath $script
                        }
                    }
                }
            } else {
                # Inline math inside current paragraph; create one if needed
                [my LastParagraph 1] appendFromScript {
                    my EvalChildScript m:oMath $script
                }
            }
        } errMsg]} {
            set valPrefix $savedValPrefix
            set inMathZone $savedInMathZone
            return -code error $errMsg
        }
        set valPrefix $savedValPrefix
        set inMathZone $savedInMathZone
    }

    # ──────────────────────────────────────────────────────────────────
    #  Accent: <m:acc><m:accPr><m:chr m:val="̂"/></m:accPr>
    #          <m:e>base</m:e></m:acc>
    #
    # Puts an accent (hat, tilde, dot, arrow, …) above the base.
    # The default accent character (when -char is omitted) is the
    # combining circumflex accent (U+0302), per ECMA-376 §22.1.2.1.
    #
    # Usage:
    #   $doc maccent { $doc mrun "x" }               ;# x-hat (default)
    #   $doc maccent -char "\u0303" { $doc mrun "x" } ;# x-tilde
    #   $doc maccent -char "\u0307" { $doc mrun "x" } ;# x-dot
    #   $doc maccent -char "\u20D7" { $doc mrun "v" } ;# vector arrow
    # ──────────────────────────────────────────────────────────────────
    method maccent {args} {
        my variable body
        variable ::ooxml::docx::properties
        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "baseScript"
            set baseScript [lindex $args end]
            $body appendFromScript {
                Tag_m:acc {
                    Tag_m:accPr {
                        my Create $properties(accPr)
                    }
                    my EvalChildScript m:e $baseScript
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # ──────────────────────────────────────────────────────────────────
    #  Bar (overline / underline):
    #   <m:bar><m:barPr><m:pos m:val="top"/></m:barPr>
    #          <m:e>base</m:e></m:bar>
    #
    # Default position is "top" (overline) per ECMA-376 §22.1.2.7.
    #
    # Usage:
    #   $doc mbar { $doc mrun "x" }                ;# overline (default)
    #   $doc mbar -pos bot { $doc mrun "x" }       ;# underline
    # ──────────────────────────────────────────────────────────────────
    method mbar {args} {
        my variable body
        variable ::ooxml::docx::properties
        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "baseScript"
            set baseScript [lindex $args end]
            $body appendFromScript {
                Tag_m:bar {
                    Tag_m:barPr {
                        my Create $properties(barPr)
                    }
                    my EvalChildScript m:e $baseScript
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # ──────────────────────────────────────────────────────────────────
    #  Delimiter: <m:d>
    #   Wraps one or more sub-expressions in matching delimiters
    #   (parentheses, brackets, braces, |·|, ‖·‖, ⌊·⌋, etc.)
    #
    #   The default delimiters are parentheses "(" and ")" per
    #   ECMA-376 §22.1.2.24. For a single argument, the script body
    #   creates content directly. For multiple arguments separated
    #   by the separator character, use mdelimarg to wrap each one
    #   in its own m:e child.
    #
    # Options:
    #   -begChr  string    opening delimiter (default "(")
    #   -endChr  string    closing delimiter (default ")")
    #   -sepChr  string    separator between arguments (default "|")
    #   -grow    CT_OnOff  delimiters grow to match content height
    #   -shp     ST_Shp    centered or match
    #
    # Usage:
    #   # Simple parentheses: (x+1)
    #   $doc mdelim { $doc mrun "x+1" }
    #
    #   # Brackets: [a, b]
    #   $doc mdelim -begChr "\[" -endChr "\]" {
    #       $doc mdelimarg { $doc mrun "a" }
    #       $doc mdelimarg { $doc mrun "b" }
    #   }
    #
    #   # Absolute value: |x|
    #   $doc mdelim -begChr "|" -endChr "|" { $doc mrun "x" }
    #
    #   # Braces with grow
    #   $doc mdelim -begChr "\{" -endChr "\}" -grow 1 {
    #       $doc mrun "expr"
    #   }
    #
    #   # Norm: ‖v‖
    #   $doc mdelim -begChr "\u2016" -endChr "\u2016" { $doc mrun "v" }
    # ──────────────────────────────────────────────────────────────────
    method mdelim {args} {
        my variable body
        variable ::ooxml::docx::properties
        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "script"
            set script [lindex $args end]
            $body appendFromScript {
                Tag_m:d {
                    Tag_m:dPr {
                        my Create $properties(dPr)
                    }
                    # Run script at the m:d level so that mdelimarg
                    # creates m:e children directly under m:d.
                    set savedbody $body
                    set dNode [dom fromScriptContext]
                    set body $dNode
                    if {$script ne ""
                        && [catch {uplevel 1 [list eval $script]} errMsg]} {
                        set body $savedbody
                        error $errMsg
                    }
                    set body $savedbody
                    # If user did not use mdelimarg, no m:e exists
                    # yet — wrap all non-dPr children in a single m:e.
                    set hasE 0
                    foreach child [$dNode childNodes] {
                        if {[$child nodeName] eq "m:e"} {
                            set hasE 1
                            break
                        }
                    }
                    if {!$hasE} {
                        set loose [list]
                        foreach child [$dNode childNodes] {
                            if {[$child nodeName] ne "m:dPr"} {
                                lappend loose $child
                            }
                        }
                        set ns "http://schemas.openxmlformats.org/officeDocument/2006/math"
                        set eNode [[$dNode ownerDocument] createElementNS $ns "m:e"]
                        $dNode appendChild $eNode
                        foreach child $loose {
                            $eNode appendChild $child
                        }
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # Helper: creates a delimiter argument (m:e) inside mdelim.
    # Each call wraps its script in an <m:e> child of the
    # enclosing <m:d>.
    #
    # For multi-argument delimiters the main mdelim script should
    # consist solely of mdelimarg calls:
    #
    #   $doc mdelim -sepChr "," {
    #       $doc mdelimarg { $doc mrun "a" }
    #       $doc mdelimarg { $doc mrun "b" }
    #       $doc mdelimarg { $doc mrun "c" }
    #   }
    method mdelimarg {script} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                my EvalChildScript m:e $script
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # Fraction: <m:f> with optional <m:fPr><m:type m:val="..."/></m:fPr>
    # Usage: doc mfrac { ...numerator... } { ...denominator... } ?-type bar|lin|noBar|skw?
    method mfrac {args} {
        my variable body
        variable ::ooxml::docx::properties
        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-2] "" "numScript denScript"
            set numScript [lindex $args end-1]
            set denScript [lindex $args end]
            $body appendFromScript {
                Tag_m:f {
                    Tag_m:fPr {
                        my Create $properties(fPr)
                    }
                    my EvalChildScript m:num $numScript
                    my EvalChildScript m:den $denScript
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # ──────────────────────────────────────────────────────────────────
    #  Grouping character (underbrace / overbrace with label):
    #   <m:groupChr><m:groupChrPr>…</m:groupChrPr>
    #              <m:e>base</m:e></m:groupChr>
    #
    # The default character is U+23DF (bottom curly bracket ⏟) per
    # ECMA-376 §22.1.2.41, positioned at the bottom with the label
    # above.
    #
    # Options:
    #   -char    string    grouping character (default ⏟)
    #   -pos     ST_TopBot position of the character (top|bot, default bot)
    #   -vertJc  ST_TopBot vertical justification of the base relative
    #                      to the grouping character (top|bot)
    #
    # Usage:
    #   # Underbrace (default)
    #   $doc mgroupchr { $doc mrun "a+b+c" }
    #
    #   # Overbrace
    #   $doc mgroupchr -char "\u23DE" -pos top -vertJc bot {
    #       $doc mrun "x+y"
    #   }
    # ──────────────────────────────────────────────────────────────────
    method mgroupchr {args} {
        my variable body
        variable ::ooxml::docx::properties
        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "baseScript"
            set baseScript [lindex $args end]
            $body appendFromScript {
                Tag_m:groupChr {
                    Tag_m:groupChrPr {
                        my Create $properties(groupChrPr)
                    }
                    my EvalChildScript m:e $baseScript
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # Lower limit wrapper: <m:limLow><m:e>base</m:e><m:lim>low</m:lim></m:limLow>
    method mlimlow {baseScript lowScript} {
        if {[catch {
            my RequireMathZone
            my variable body
            $body appendFromScript {
                Tag_m:limLow {
                    my EvalChildScript m:e $baseScript
                    my EvalChildScript m:lim $lowScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # Upper limit wrapper: <m:limUpp><m:e>base</m:e><m:lim>up</m:lim></m:limUpp>
    method mlimupp {baseScript upScript} {
        if {[catch {
            my RequireMathZone
            my variable body
            $body appendFromScript {
                Tag_m:limUpp {
                    my EvalChildScript m:e $baseScript
                    my EvalChildScript m:lim $upScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # ──────────────────────────────────────────────────────────────────
    #  Matrix: <m:m><m:mPr>…</m:mPr><m:mr>…</m:mr>…</m:m>
    #
    #  The matrix body is a list of row scripts. Each row script
    #  creates cells via mmatcell, which wraps content in an <m:e>.
    #
    # Options:
    #   -baseJc  ST_BaseJc  vertical alignment of the matrix relative to
    #                       the math baseline (top|center|bottom)
    #   -mcJc    ST_McJc   column justification (left|center|right)
    #   -count   0..255    number of adjacent columns sharing the same
    #                       justification (default: omitted)
    #
    # Usage:
    #   # 2×2 identity matrix
    #   $doc mmatrix {
    #       {
    #           $doc mmatcell { $doc mrun "1" }
    #           $doc mmatcell { $doc mrun "0" }
    #       }
    #       {
    #           $doc mmatcell { $doc mrun "0" }
    #           $doc mmatcell { $doc mrun "1" }
    #       }
    #   }
    #
    #   # Centered columns
    #   $doc mmatrix -mcJc center {
    #       { $doc mmatcell { $doc mrun "a" }
    #         $doc mmatcell { $doc mrun "b" } }
    #       { $doc mmatcell { $doc mrun "c" }
    #         $doc mmatcell { $doc mrun "d" } }
    #   }
    #
    #   # Wrap in brackets for a proper matrix notation: [M]
    #   $doc mdelim -begChr "\[" -endChr "\]" {
    #       $doc mmatrix {
    #           { $doc mmatcell { $doc mrun "a" }
    #             $doc mmatcell { $doc mrun "b" } }
    #           { $doc mmatcell { $doc mrun "c" }
    #             $doc mmatcell { $doc mrun "d" } }
    #       }
    #   }
    # ──────────────────────────────────────────────────────────────────
    method mmatrix {args} {
        my variable body
        variable ::ooxml::docx::properties

        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "rowScripts"
            set rowScripts [lindex $args end]
            if {![llength $rowScripts]} {
                error "mmatrix requires at least one row"
            }
            set baseJc [my EatOption -baseJc ST_BaseJc]
            $body appendFromScript {
                Tag_m:m {
                    # Matrix properties
                    Tag_m:mPr {
                        if {$baseJc ne ""} {
                            Tag_m:baseJc m:val $baseJc
                        }
                        # Column specification — only emitted if the
                        # user provided -mcJc or -count
                        set hasMcPr 0
                        foreach {opt _} [array get opts] {
                            if {$opt in {-mcJc -count}} {
                                set hasMcPr 1
                                break
                            }
                        }
                        if {$hasMcPr} {
                            Tag_m:mcs {
                                Tag_m:mc {
                                    Tag_m:mcPr {
                                        my Create $properties(mcPr)
                                    }
                                }
                            }
                        }
                    }
                    # Rows — each rowScript is evaluated to produce
                    # mmatcell calls that create m:e children of m:mr
                    foreach rowScript $rowScripts {
                        Tag_m:mr {
                            set savedbody $body
                            set body [dom fromScriptContext]
                            if {[catch {uplevel 1 [list eval $rowScript]} errMsg]} {
                                set body $savedbody
                                error $errMsg
                            }
                            set body $savedbody
                        }
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # Helper: creates a matrix cell (m:e) inside a matrix row (m:mr).
    method mmatcell {script} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                my EvalChildScript m:e $script
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # Function application: <m:func><m:fName>name</m:fName><m:e>arg</m:e></m:func>
    method mfunc {nameScript argScript} {
        my variable body
        if {[catch {
            my RequireMathZone 
            $body appendFromScript {
                Tag_m:func {
                    my EvalChildScript m:fName $nameScript
                    my EvalChildScript m:e $argScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # n-ary operator (sum/product/integral/⋂ etc.)
    # Usage:
    #   doc mnary { base } -char "∑" -limLoc undOvr -sub { i=1 } -sup { n } -grow 1 -subHide 0 -supHide 0
    method mnary {args} {
        my variable body
        variable ::ooxml::docx::properties

        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "baseScript"
            set baseScript [lindex $args end]
            set subSc  [my EatOption -sub      NoCheck]
            set supSc  [my EatOption -sup      NoCheck]
            $body appendFromScript {
                Tag_m:nary {
                    Tag_m:naryPr {
                        my Create $properties(naryPr)
                    }
                    my EvalChildScript m:sub $subSc
                    my EvalChildScript m:sup $supSc
                    my EvalChildScript m:e $baseScript
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # nth-root: <m:rad><m:deg>n</m:deg><m:e>…</m:e></m:rad>
    # Per ECMA-376 §22.1.2.88, m:radPr is optional and omitted when
    # there are no properties to set.
    method mroot {degreeScript radicandScript} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                Tag_m:rad {
                    my EvalChildScript m:deg $degreeScript
                    my EvalChildScript m:e $radicandScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    method mrun {args} {
        variable ::ooxml::docx::properties
        
        if {[catch {
            my RequireMathZone
            OptVal [lrange $args 0 end-1] "" "text"
            set text [lindex $args end]
            set brk    [my EatOption -brk     CT_OnOff]
            set brkAt  [my EatOption -brkAt   ST_Integer255]

            Tag_m:r {
                Tag_m:rPr {
                    my Create $properties(mrun1)
                    if {$brk eq "1"} {
                        if {$brkAt ne ""} {
                            Tag_m:brk m:alnAt $brkAt
                        } else {
                            Tag_m:brk
                        }
                    }
                    my Create $properties(mrun2)
                }
                my Mt $text
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    # Subscript: <m:sSub><m:e>base</m:e><m:sub>sub</m:sub></m:sSub>
    method msub {baseScript subScript} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                Tag_m:sSub {
                    my EvalChildScript m:e $baseScript
                    my EvalChildScript m:sub $subScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # Sub+Sup: <m:sSubSup><m:e>base</m:e><m:sub>…</m:sub><m:sup>…</m:sup></m:sSubSup>
    method msubsup {baseScript subScript supScript} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                Tag_m:sSubSup {
                    my EvalChildScript m:e $baseScript
                    my EvalChildScript m:sub $subScript
                    my EvalChildScript m:sup $supScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # Superscript: <m:sSup><m:e>base</m:e><m:sup>sup</m:sup></m:sSup>
    method msup {baseScript supScript} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                Tag_m:sSup {
                    my EvalChildScript m:e $baseScript
                    my EvalChildScript m:sup $supScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

    # Square root: <m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>
    #              <m:deg/><m:e>…</m:e></m:rad>
    # Per ECMA-376 §22.1.2.89, degHide must be explicitly set and
    # an empty m:deg element must be present for strict conformance.
    method msqrt {radicandScript} {
        my variable body
        if {[catch {
            my RequireMathZone
            $body appendFromScript {
                Tag_m:rad {
                    Tag_m:radPr {
                        Tag_m:degHide m:val 1
                    }
                    Tag_m:deg
                    my EvalChildScript m:e $radicandScript
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }            
    }

}
