#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2025 Rolf Ade, DE
#  Copyright (C) 2025 Miguel Bañón
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

# OMML related methods, the initially code contributed by Miguel Bañón
oo::define ooxml::docx::docx {
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

        set savedValPrefix $valPrefix
        set valPrefix m
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
            return -code error $errMsg
        }
        set valPrefix $savedValPrefix
    }

    # Fraction: <m:f> with optional <m:fPr><m:type m:val="..."/></m:fPr>
    # Usage: doc mfrac { ...numerator... } { ...denominator... } ?-type bar|lin|noBar|skw?
    method mfrac {args} {
        my variable body
        variable ::ooxml::docx::properties
        if {[catch {
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

    # Lower limit wrapper: <m:limLow><m:e>base</m:e><m:lim>low</m:lim></m:limLow>
    method mlimlow {baseScript lowScript} {
        my variable body
        $body appendFromScript {
            Tag_m:limLow {
                my EvalChildScript m:e $baseScript
                my EvalChildScript m:lim $lowScript
            }
        }
    }

    # Upper limit wrapper: <m:limUpp><m:e>base</m:e><m:lim>up</m:lim></m:limUpp>
    method mlimupp {baseScript upScript} {
        my variable body
        $body appendFromScript {
            Tag_m:limUpp {
                my EvalChildScript m:e $baseScript
                my EvalChildScript m:lim $upScript
            }
        }
    }

    # Function application: <m:func><m:fName>name</m:fName><m:e>arg</m:e></m:func>
    method mfunc {nameScript argScript} {
        my variable body
        $body appendFromScript {
            Tag_m:func {
                my EvalChildScript m:fName $nameScript
                my EvalChildScript m:e $argScript
            }
        }
    }

    # n-ary operator (sum/product/integral/⋂ etc.)
    # Usage:
    #   doc mnary { base } -char "∑" -limLoc undOvr -sub { i=1 } -sup { n } -grow 1 -subHide 0 -supHide 0
    method mnary {args} {
        my variable body
        variable ::ooxml::docx::properties

        if {[catch {
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
    method mroot {degreeScript radicandScript} {
        my variable body
        $body appendFromScript {
            Tag_m:rad {
                my EvalChildScript m:deg $degreeScript
                my EvalChildScript m:e $radicandScript
            }
        }
    }

    method mrun {args} {
        my variable body
        variable ::ooxml::docx::properties
        
        if {[catch {
            OptVal [lrange $args 0 end-1] "" "<mrun script"
            set text [lindex $args end]
            set brk    [my EatOption -brk     CT_OnOff]
            set brkAt  [my EatOption -brkAt   ST_Integer255]

            Tag_m:r {
                Tag_m:rPr {
                    my Create $properties(mrun1)
                    if {$brk eq "on"} {
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
        $body appendFromScript {
            Tag_m:sSub {
                my EvalChildScript m:e $baseScript
                my EvalChildScript m:sub $subScript
            }
        }
    }

    # Sub+Sup: <m:sSubSup><m:e>base</m:e><m:sub>…</m:sub><m:sup>…</m:sup></m:sSubSup>
    method msubsup {baseScript subScript supScript} {
        my variable body
        $body appendFromScript {
            Tag_m:sSubSup {
                my EvalChildScript m:e $baseScript
                my EvalChildScript m:sub $subScript
                my EvalChildScript m:sup $supScript
            }
        }
    }

    # Superscript: <m:sSup><m:e>base</m:e><m:sup>sup</m:sup></m:sSup>
    method msup {baseScript supScript} {
        my variable body
        $body appendFromScript {
            Tag_m:sSup {
                my EvalChildScript m:e $baseScript
                my EvalChildScript m:sup $supScript
            }
        }
    }

    # Square root: <m:rad><m:e>…</m:e></m:rad>  (no degree)
    method msqrt {radicandScript} {
        my variable body
        $body appendFromScript {
            Tag_m:rad {
                my EvalChildScript m:e $radicandScript
            }
        }
    }

}
