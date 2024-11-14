#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2024 Rolf Ade, DE
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
package require tdom 0.9.0-
package require msgcat
package require ooxml

namespace eval ::ooxml {

    namespace export docx_write

    variable xmlns

    array set xmlns {
        o urn:schemas-microsoft-com:office:office
        v urn:schemas-microsoft-com:vml
        w http://schemas.openxmlformats.org/wordprocessingml/2006/main
        w10 urn:schemas-microsoft-com:office:word
        wp http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
        wps http://schemas.microsoft.com/office/word/2010/wordprocessingShape
        wpg http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
        mc http://schemas.openxmlformats.org/markup-compatibility/2006
        wp14 http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
        w14 http://schemas.microsoft.com/office/word/2010/wordml
    }
    
}

proc ooxml::InitDocxNodeCommands {} {
    set elementNodes {
        w:body w:p w:pPr w:pStyle w:r w:rPr w:t
    }
    namespace eval ::ooxml "dom createNodeCmd textNode Text; namespace export Text"
    foreach tag $elementNodes {
        namespace eval ::ooxml "dom createNodeCmd -tagName $tag elementNode Tag_$tag
                                namespace export Tag_$tag"
    }
    namespace eval ::ooxml "dom createNodeCmd textNode Text; namespace export Text"
}
::ooxml::InitDocxNodeCommands

proc ::ooxml::InitStaticDocx {} {
    if {[info exists ::ooxml::staticDocx]} return
    foreach {name xml} {
        [Content_Types].xml {
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="png" ContentType="image/png"/>
                <Default Extension="jpeg" ContentType="image/jpeg"/>
                <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
                <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
                <Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
                <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
            </Types>
        }
        _rels/.rels {
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
                <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
        }
        word/_rels/document.xml.rels {
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
                <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
            </Relationships>
        }
        docProps/app.xml {
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            </Properties>
        }
        docProps/core.xml {
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                <dcterms:created xsi:type="dcterms:W3CDTF">2024-10-30T15:52:52Z</dcterms:created>
                <dc:creator/>
                <dc:description/>
                <dc:language>de-DE</dc:language>
                <cp:lastModifiedBy/>
                <dcterms:modified xsi:type="dcterms:W3CDTF">2024-10-30T15:53:59Z</dcterms:modified>
                <cp:revision>1</cp:revision>
                <dc:subject/>
                <dc:title/>
            </cp:coreProperties>
        }
        word/fontTable.xml {        
            <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <w:font w:name="Times New Roman">
                    <w:charset w:val="00" w:characterSet="windows-1252"/>
                    <w:family w:val="roman"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Symbol">
                    <w:charset w:val="02"/>
                    <w:family w:val="roman"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Arial">
                    <w:charset w:val="00" w:characterSet="windows-1252"/>
                    <w:family w:val="swiss"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Liberation Serif">
                    <w:altName w:val="Times New Roman"/>
                    <w:charset w:val="01" w:characterSet="utf-8"/>
                    <w:family w:val="roman"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Liberation Sans">
                    <w:altName w:val="Arial"/>
                    <w:charset w:val="01" w:characterSet="utf-8"/>
                    <w:family w:val="swiss"/>
                    <w:pitch w:val="variable"/>
                </w:font>
            </w:fonts>
        }
        word/settings.xml {
            <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:zoom w:val="bestFit" w:percent="228"/>
                <w:defaultTabStop w:val="709"/>
                <w:autoHyphenation w:val="true"/>
                <w:compat>
                    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
                </w:compat>
            </w:settings>
        }
        word/styles.xml {
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">
                <w:docDefaults>
                    <w:rPrDefault>
                        <w:rPr>
                            <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:eastAsia="AR PL KaitiM GB" w:cs="FreeSans"/>
                            <w:kern w:val="2"/>
                            <w:sz w:val="24"/>
                            <w:szCs w:val="24"/>
                            <w:lang w:val="de-DE" w:eastAsia="zh-CN" w:bidi="hi-IN"/>
                        </w:rPr>
                    </w:rPrDefault>
                    <w:pPrDefault>
                        <w:pPr>
                            <w:widowControl/>
                            <w:suppressAutoHyphens w:val="true"/>
                        </w:pPr>
                    </w:pPrDefault>
                </w:docDefaults>
                <w:style w:type="paragraph" w:styleId="Normal">
                    <w:name w:val="Normal"/>
                    <w:qFormat/>
                    <w:pPr>
                        <w:widowControl/>
                        <w:bidi w:val="0"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:eastAsia="AR PL KaitiM GB" w:cs="FreeSans"/>
                        <w:color w:val="auto"/>
                        <w:kern w:val="2"/>
                        <w:sz w:val="24"/>
                        <w:szCs w:val="24"/>
                        <w:lang w:val="de-DE" w:eastAsia="zh-CN" w:bidi="hi-IN"/>
                    </w:rPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Berschrift1">
                    <w:name w:val="Überschrift"/>
                    <w:basedOn w:val="Normal"/>
                    <w:next w:val="Textkrper"/>
                    <w:qFormat/>
                    <w:pPr>
                        <w:keepNext w:val="true"/>
                        <w:spacing w:before="240" w:after="120"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans" w:eastAsia="AR PL KaitiM GB" w:cs="FreeSans"/>
                        <w:sz w:val="28"/>
                        <w:szCs w:val="28"/>
                    </w:rPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Textkrper">
                    <w:name w:val="Body Text"/>
                    <w:basedOn w:val="Normal"/>
                    <w:pPr>
                        <w:spacing w:lineRule="auto" w:line="276" w:before="0" w:after="140"/>
                    </w:pPr>
                    <w:rPr/>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Aufzhlung">
                    <w:name w:val="List"/>
                    <w:basedOn w:val="Textkrper"/>
                    <w:pPr/>
                    <w:rPr>
                        <w:rFonts w:cs="FreeSans"/>
                    </w:rPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Beschriftung">
                    <w:name w:val="Caption"/>
                    <w:basedOn w:val="Normal"/>
                    <w:qFormat/>
                    <w:pPr>
                        <w:suppressLineNumbers/>
                        <w:spacing w:before="120" w:after="120"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:cs="FreeSans"/>
                        <w:i/>
                        <w:iCs/>
                        <w:sz w:val="24"/>
                        <w:szCs w:val="24"/>
                    </w:rPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Verzeichnis">
                    <w:name w:val="Verzeichnis"/>
                    <w:basedOn w:val="Normal"/>
                    <w:qFormat/>
                    <w:pPr>
                        <w:suppressLineNumbers/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:cs="FreeSans"/>
                    </w:rPr>
                </w:style>
            </w:styles>
        }
    } {
        set ::ooxml::staticDocx($name) [dom parse $xml]
    }
}

::ooxml::InitStaticDocx

oo::class create ooxml::docx_write {
    constructor { args } {
        my variable document
        my variable body
        variable ::ooxml::xmlns

        ::ooxml::InitNodeCommands
        namespace import ::ooxml::Tag_*  ::ooxml::Text
        
        dom createDocument w:document document
        $document documentElement root
        foreach ns {
            o
            r
            v
            w
            w10
            wp
            wps
            wpg
            mc
            wp14
            w14
        } {
            $root setAttribute xmlns:$ns $xmlns($ns)
        }
        $root setAttribute mc:Ignorable "w14 wp14"
        $root appendFromScript Tag_w:body
        set body [$root firstChild]
    }

    destructor {
    }

    method importStyles {docxfile} {


    }

    method text {text args} {
        my variable body

        set len [llength $args]
        set idx 0
        for {set idx 0} {$idx < $len} {incr idx} {
            switch -- [set opt [lindex $args $idx]] {
                -style {
                    incr idx
                    if {!($idx < $len)} {
                        error "option '$opt': missing argument"
                    }            
                }
                default {
                    error "unknown option \"$opt\", should be: -style"
                }
            }
            set opts([string range $opt 1 end]) [lindex $args $idx]
        }
        
        $body appendFromScript {
            Tag_w:p {
                if {[info exists opts(style)]} {
                    Tag_w:pPr {
                        Tag_w:pStyle w:val $opts(style)
                        Tag_w:rPr
                    }
                }
                Tag_w:r {
                    my Wt $text
                }
            }
        }
    }

    method appendText {text args} {
        my variable body
        set len [llength $args]
        set idx 0
        for {set idx 0} {$idx < $len} {incr idx} {
            switch -- [set opt [lindex $args $idx]] {
                -style {
                    incr idx
                    if {!($idx < $len)} {
                        error "option '$opt': missing argument"
                    }            
                }
                default {
                    error "unknown option \"$opt\", should be: -style"
                }
            }
            set opts([string range $opt 1 end]) [lindex $args $idx]
        }

        # Identify the last paragraph
        set p [$body lastChild]
        while {$p ne ""} {
            if {[$p nodeType] ne "ELEMENT_NODE"} {
                set child [$p previousSibling]
                continue
            }
            if {[$p nodeName] ne "w:p"} {
                set child [$p previousSibling]
                continue
            }
            break
        }
        if {$p eq ""} {
            error "no paragraph to append to in the document"
        }
        $p appendFromScript {
            Tag_w:r {
                if {[info exists opts(style)]} {
                    Tag_w:rPr {
                        Tag_w:pStyle w:val $opts(style)
                        Tag_w:rPr
                    }
                }
                my Wt $text
            }
        }
    }

    method Wt {text} {
        set atts ""
        if {[string index $text 0] eq " " || [string index $text end] eq " "} {
            lappend atts xml:space preserve
        }
        #Tag_w:pPr
        Tag_w:t $atts {
            Text [dom clearString -replace $text]
        }
    }
        
    method write {file} {
        my variable document
        my variable body
        variable ::ooxml::xmlns
        variable ::ooxml::staticDocx

        ooxml::InitNodeCommands
        namespace import ::ooxml::Tag_* ::ooxml::Text

        # Initialize zip file
        set file [string trim $file]
        if {$file eq {}} {
            set file {spreadsheetml.xlsx}
        }
        if {[file extension $file] ne {.docx}} {
            append file .docx
        }
        if {[catch {open $file w} zf]} {
            error "cannot open file $file for writing"
        }
        fconfigure $zf -translation binary -eofchar {}

        foreach staticFile {
            [Content_Types].xml
            _rels/.rels
            docProps/app.xml
            docProps/core.xml
            word/_rels/document.xml.rels
            word/fontTable.xml
            word/settings.xml
            word/styles.xml
        } {
            ::ooxml::Dom2zip $zf $staticDocx($staticFile) $staticFile cd count
        }

        # word/document.xml
        ::ooxml::Dom2zip $zf $document "word/document.xml" cd count
        
        
        # Finalize zip.
        set cdoffset [tell $zf]
        set endrec [binary format a4ssssiis PK\05\06 0 0 $count $count [string length $cd] $cdoffset 0]
        puts -nonewline $zf $cd
        puts -nonewline $zf $endrec
        close $zf
        return 0
    }
}
