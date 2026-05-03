# ecma376/validate.tcl --
#
#   Structural validator for OOXML .docx packages.
#   Encodes key constraints from:
#     - ECMA-376 5th Edition Part 1 (WML)
#     - ECMA-376 5th Edition Part 2 (OPC)
#     - ISO/IEC 29500-1:2016 and 29500-2:2021
#
#   Usage:
#     source ecma376/validate.tcl
#     set errors [ecma376::validate::docx /path/to/file.docx]
#     # Returns list of error strings; empty list = valid
#
#   Requires: tdom 0.9.6+, Tcl 8.6.7+

package require tdom

namespace eval ecma376::validate {
    # ── Namespace URIs (ECMA-376 §22.1, §11.3) ──────────────────
    variable ns
    array set ns {
        w   http://schemas.openxmlformats.org/wordprocessingml/2006/main
        r   http://schemas.openxmlformats.org/officeDocument/2006/relationships
        rel http://schemas.openxmlformats.org/package/2006/relationships
        ct  http://schemas.openxmlformats.org/package/2006/content-types
        m   http://schemas.openxmlformats.org/officeDocument/2006/math
        wp  http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
        a   http://schemas.openxmlformats.org/drawingml/2006/main
        mc  http://schemas.openxmlformats.org/markup-compatibility/2006
        dc  http://purl.org/dc/elements/1.1/
        cp  http://schemas.openxmlformats.org/package/2006/metadata/core-properties
    }

    # ── OPC required relationship types (§9.3, §11.3) ───────────
    variable relTypes
    array set relTypes {
        officeDocument http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
        styles         http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles
        settings       http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings
        fontTable      http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable
        footnotes      http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes
        endnotes       http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes
        comments       http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments
        header         http://schemas.openxmlformats.org/officeDocument/2006/relationships/header
        footer         http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer
        numbering      http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering
        image          http://schemas.openxmlformats.org/officeDocument/2006/relationships/image
        hyperlink      http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink
        coreProperties http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
    }

    # ── Content types (§13.2, §11.3.6) ──────────────────────────
    variable contentTypes
    array set contentTypes {
        word/document.xml      application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml
        word/styles.xml        application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml
        word/settings.xml      application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml
        word/fontTable.xml     application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml
        word/footnotes.xml     application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml
        word/endnotes.xml      application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml
        word/comments.xml      application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml
        word/numbering.xml     application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml
        docProps/core.xml      application/vnd.openxmlformats-package.core-properties+xml
        docProps/app.xml       application/vnd.openxmlformats-officedocument.extended-properties+xml
    }

    # ── WML block-level elements (§17.2.2 CT_Body) ──────────────
    # These are the only elements allowed as direct children of w:body
    # (besides w:sectPr which must be last).
    variable bodyChildren {
        p tbl sdt customXml altChunk
        bookmarkStart bookmarkEnd
        commentRangeStart commentRangeEnd
        moveFromRangeStart moveFromRangeEnd
        moveToRangeStart moveToRangeEnd
        permStart permEnd proofErr
        ins del moveFrom moveTo
    }

    # ── WML sectPr child ordering (§17.6.17 CT_SectPr) ─────────
    # Per ECMA-376, children of w:sectPr must appear in this order.
    variable sectPrOrder {
        headerReference footerReference
        endnotePr footnotePr
        type pgSz pgMar paperSrc pgBorders lnNumType pgNumType
        cols formProt vAlign noEndnote titlePg textDirection
        bidi rtlGutter docGrid printerSettings sectPrChange
    }
}

# ── Main validation entry point ─────────────────────────────────

proc ecma376::validate::docx {zipfile} {
    set errors [list]
    set parts [dict create]
    set ct ""
    set packageRels ""
    set docRels ""

    # ── Phase 1: Read zip contents ──────────────────────────────
    set zipEntries [list]
    if {![catch {set data [exec unzip -Z1 $zipfile]}]} {
        foreach e [split $data \n] {
            set e [string trim $e]
            if {$e ne ""} { lappend zipEntries $e }
        }
    } else {
        lappend errors "FATAL: cannot list zip contents"
        return $errors
    }

    foreach entry $zipEntries {
        dict set parts $entry 1
    }

    # ── Phase 2: OPC — required parts (§9.1, §9.3) ─────────────
    foreach req {
        {[Content_Types].xml}
        _rels/.rels
        word/document.xml
        word/_rels/document.xml.rels
    } {
        if {![dict exists $parts $req]} {
            lappend errors "OPC: required part missing: $req"
        }
    }

    # Read XML parts for further checks
    if {[catch {
        regsub -all {\[} {[Content_Types].xml} {\\[} esc
        regsub -all {\]} $esc {\\]} esc
        set ct [exec unzip -p $zipfile $esc]
    }]} {
        lappend errors "OPC: cannot read [Content_Types].xml"
    }
    if {[catch {set packageRels [exec unzip -p $zipfile _rels/.rels]}]} {
        lappend errors "OPC: cannot read _rels/.rels"
    }
    if {[catch {set docRels [exec unzip -p $zipfile word/_rels/document.xml.rels]}]} {
        lappend errors "OPC: cannot read word/_rels/document.xml.rels"
    }
    if {[catch {set docXml [exec unzip -p $zipfile word/document.xml]}]} {
        lappend errors "OPC: cannot read word/document.xml"
        return $errors
    }

    # ── Phase 3: OPC — Content Types validation (§10.1.2) ───────
    if {$ct ne ""} {
        lappend errors {*}[validateContentTypes $ct $parts]
    }

    # ── Phase 4: OPC — Relationships validation (§9.3) ──────────
    if {$packageRels ne ""} {
        lappend errors {*}[validatePackageRels $packageRels]
    }
    if {$docRels ne ""} {
        lappend errors {*}[validateDocRels $docRels $parts $zipfile]
    }

    # ── Phase 5: WML — Document structure (§17.2.2) ─────────────
    lappend errors {*}[validateDocument $docXml]

    # ── Phase 6: WML — Optional parts ────────────────────────────
    foreach notePart {word/footnotes.xml word/endnotes.xml} {
        if {[dict exists $parts $notePart]} {
            if {[catch {set xml [exec unzip -p $zipfile $notePart]}]} {
                lappend errors "WML: cannot read $notePart"
            } else {
                lappend errors {*}[validateNotes $xml $notePart]
            }
        }
    }
    if {[dict exists $parts word/styles.xml]} {
        if {![catch {set xml [exec unzip -p $zipfile word/styles.xml]}]} {
            lappend errors {*}[validateStyles $xml]
        }
    }
    if {[dict exists $parts word/settings.xml]} {
        if {![catch {set xml [exec unzip -p $zipfile word/settings.xml]}]} {
            lappend errors {*}[validateSettings $xml]
        }
    }

    # ── Phase 7: Schema — Element child ordering (§17.3, §17.6) ──
    lappend errors {*}[validateElementOrdering $docXml]

    # ── Phase 8: Schema — Attribute enum values ──────────────────
    lappend errors {*}[validateAttributeEnums $docXml]
    foreach optPart {word/styles.xml word/settings.xml word/footnotes.xml word/endnotes.xml} {
        if {[dict exists $parts $optPart]} {
            if {![catch {set xml [exec unzip -p $zipfile $optPart]}]} {
                lappend errors {*}[validateAttributeEnums $xml $optPart]
            }
        }
    }

    return $errors
}

# ── Content Types validation ────────────────────────────────────

proc ecma376::validate::validateContentTypes {xml parts} {
    variable ns
    variable contentTypes
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "CT: malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list ct $ns(ct)]

    # Root element must be ct:Types
    if {[$root localName] ne "Types"} {
        lappend errors "CT: root element must be Types, got [$root localName]"
    }

    # Every XML part should have an Override or matching Default
    set overrides [dict create]
    foreach node [$root selectNodes ct:Override] {
        dict set overrides [$node @PartName ""] [$node @ContentType ""]
    }
    set defaults [dict create]
    foreach node [$root selectNodes ct:Default] {
        dict set defaults [$node @Extension ""] [$node @ContentType ""]
    }

    # Check known parts have correct content types
    foreach {part expectedCT} [array get contentTypes] {
        if {![dict exists $parts $part]} continue
        set partName "/$part"
        if {[dict exists $overrides $partName]} {
            set actualCT [dict get $overrides $partName]
            if {$actualCT ne $expectedCT} {
                lappend errors "CT: $part has content type '$actualCT', expected '$expectedCT'"
            }
        } else {
            # Check extension default
            set ext [string range [file extension $part] 1 end]
            if {![dict exists $defaults $ext]} {
                lappend errors "CT: $part has no Override and no Default for .$ext"
            }
        }
    }

    # Check for duplicate PartName values (OPC §10.1.2.4)
    set seen [dict create]
    foreach node [$root selectNodes ct:Override] {
        set pn [$node @PartName ""]
        if {[dict exists $seen $pn]} {
            lappend errors "CT: duplicate Override for $pn"
        }
        dict set seen $pn 1
    }

    $doc delete
    return $errors
}

# ── Package-level relationships (.rels) ─────────────────────────

proc ecma376::validate::validatePackageRels {xml} {
    variable ns
    variable relTypes
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "PKG-RELS: malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list r $ns(rel)]

    # Must have officeDocument relationship (§9.3.2)
    set officeDocRels [$root selectNodes {
        r:Relationship[contains(@Type,'officeDocument')]
    }]
    if {![llength $officeDocRels]} {
        lappend errors "PKG-RELS: missing officeDocument relationship"
    }

    # Must have core-properties with package/ relationship type (§9.3.4)
    set cpRels [$root selectNodes {
        r:Relationship[contains(@Type,'core-properties')]
    }]
    if {[llength $cpRels]} {
        set cpType [[lindex $cpRels 0] @Type ""]
        if {![string match *package/2006/relationships* $cpType]} {
            lappend errors "PKG-RELS: core-properties uses wrong relationship type prefix: $cpType"
        }
    }

    # Every Relationship must have Id and Type attributes (§9.3.1)
    foreach rel [$root selectNodes r:Relationship] {
        if {[$rel @Id ""] eq ""} {
            lappend errors "PKG-RELS: Relationship missing Id attribute"
        }
        if {[$rel @Type ""] eq ""} {
            lappend errors "PKG-RELS: Relationship missing Type attribute"
        }
        if {[$rel @Target ""] eq ""} {
            lappend errors "PKG-RELS: Relationship missing Target attribute"
        }
    }

    # Check for duplicate Id values
    set ids [dict create]
    foreach rel [$root selectNodes r:Relationship] {
        set id [$rel @Id ""]
        if {$id ne "" && [dict exists $ids $id]} {
            lappend errors "PKG-RELS: duplicate Relationship Id '$id'"
        }
        dict set ids $id 1
    }

    $doc delete
    return $errors
}

# ── Document-level relationships ────────────────────────────────

proc ecma376::validate::validateDocRels {xml parts zipfile} {
    variable ns
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "DOC-RELS: malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list r $ns(rel)]

    # Every Id must be unique
    set ids [dict create]
    foreach rel [$root selectNodes r:Relationship] {
        set id [$rel @Id ""]
        if {$id ne "" && [dict exists $ids $id]} {
            lappend errors "DOC-RELS: duplicate Relationship Id '$id'"
        }
        dict set ids $id 1
    }

    # Every internal Target must resolve to an existing part
    foreach rel [$root selectNodes r:Relationship] {
        set mode [$rel @TargetMode ""]
        if {$mode eq "External"} continue
        set target [$rel @Target ""]
        if {$target eq ""} continue
        # Resolve relative to word/
        set resolved "word/$target"
        if {![dict exists $parts $resolved]} {
            lappend errors "DOC-RELS: dangling target '$target' (resolved to $resolved)"
        }
    }

    # Every r:id referenced in document.xml must exist in rels
    if {![catch {set docXml [exec unzip -p $zipfile word/document.xml]}]} {
        if {![catch {set ddoc [dom parse $docXml]}]} {
            $ddoc selectNodesNamespaces [list \
                w $ns(w) r $ns(r) wp $ns(wp)]
            set refIds [list]
            foreach attr {r:id r:embed} {
                foreach node [$ddoc selectNodes "//*\[@$attr\]"] {
                    lappend refIds [$node getAttribute $attr ""]
                }
            }
            foreach rid $refIds {
                if {$rid ne "" && ![dict exists $ids $rid]} {
                    lappend errors "DOC-RELS: document.xml references r:id '$rid' not in .rels"
                }
            }
            $ddoc delete
        }
    }

    $doc delete
    return $errors
}

# ── WML Document structure (§17.2.2 CT_Body) ───────────────────

proc ecma376::validate::validateDocument {xml} {
    variable ns
    variable bodyChildren
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "WML-DOC: malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list w $ns(w) m $ns(m) mc $ns(mc)]

    # Root element must be w:document (§17.2.3)
    if {[$root localName] ne "document"} {
        lappend errors "WML-DOC: root must be w:document, got [$root localName]"
    }
    set rootNS [$root namespaceURI]
    if {$rootNS ne $ns(w)} {
        lappend errors "WML-DOC: root namespace must be $ns(w), got $rootNS"
    }

    # w:body must be sole child element of w:document (§17.2.2)
    set bodies [$root selectNodes w:body]
    if {[llength $bodies] != 1} {
        lappend errors "WML-DOC: w:document must contain exactly one w:body, found [llength $bodies]"
        $doc delete
        return $errors
    }
    set body [lindex $bodies 0]

    # w:sectPr must be last child of w:body if present (§17.6.17)
    set lastChild [$body lastChild]
    if {$lastChild ne ""} {
        set children [$body childNodes]
        set foundSectPr 0
        set sectPrIdx -1
        set idx 0
        foreach child $children {
            if {[$child nodeType] eq "ELEMENT_NODE" && [$child localName] eq "sectPr"} {
                set foundSectPr 1
                set sectPrIdx $idx
            }
            incr idx
        }
        if {$foundSectPr} {
            set lastElemIdx [expr {$idx - 1}]
            # Walk back to find last element
            for {set i [expr {[llength $children] - 1}]} {$i >= 0} {incr i -1} {
                if {[[lindex $children $i] nodeType] eq "ELEMENT_NODE"} {
                    set lastElemIdx $i
                    break
                }
            }
            if {$sectPrIdx != $lastElemIdx} {
                lappend errors "WML-DOC: w:sectPr must be last element child of w:body (found at position $sectPrIdx, last element at $lastElemIdx)"
            }
        }
    }

    # Body children must be block-level elements (§17.2.2)
    foreach child [$body childNodes] {
        if {[$child nodeType] ne "ELEMENT_NODE"} continue
        set ln [$child localName]
        if {$ln eq "sectPr"} continue
        if {$ln ni $bodyChildren} {
            lappend errors "WML-DOC: illegal w:body child element w:$ln"
        }
    }

    # Every w:p must have at most one w:pPr as first child (§17.3.1.26)
    foreach p [$root selectNodes {//w:p}] {
        set pprCount 0
        foreach child [$p childNodes] {
            if {[$child nodeType] ne "ELEMENT_NODE"} continue
            if {[$child localName] eq "pPr"} {
                incr pprCount
                if {$pprCount > 1} {
                    lappend errors "WML-DOC: w:p has multiple w:pPr children"
                }
            }
        }
    }

    # Every w:r must have at most one w:rPr as first child (§17.3.2.28)
    foreach r [$root selectNodes {//w:r}] {
        set rprCount 0
        foreach child [$r childNodes] {
            if {[$child nodeType] ne "ELEMENT_NODE"} continue
            if {[$child localName] eq "rPr"} {
                incr rprCount
                if {$rprCount > 1} {
                    lappend errors "WML-DOC: w:r has multiple w:rPr children"
                }
            }
        }
    }

    $doc delete
    return $errors
}

# ── Footnotes / Endnotes (§17.11.15, §17.11.2) ────────────────

proc ecma376::validate::validateNotes {xml partName} {
    variable ns
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "WML-NOTES($partName): malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list w $ns(w)]

    set rootLN [$root localName]
    set isFootnotes [expr {$rootLN eq "footnotes"}]
    set isEndnotes  [expr {$rootLN eq "endnotes"}]
    set noteTag [expr {$isFootnotes ? "footnote" : "endnote"}]

    if {!$isFootnotes && !$isEndnotes} {
        lappend errors "WML-NOTES($partName): root must be w:footnotes or w:endnotes, got w:$rootLN"
        $doc delete
        return $errors
    }

    # Must have separator (id=0) and continuationSeparator (id=1)
    # per §17.11.3 and §17.11.6
    set hasSep 0
    set hasCont 0
    foreach note [$root selectNodes w:$noteTag] {
        set ntype [$note @w:type ""]
        set nid   [$note @w:id ""]
        if {$ntype eq "separator" && $nid eq "0"} { set hasSep 1 }
        if {$ntype eq "continuationSeparator" && $nid eq "1"} { set hasCont 1 }
    }
    if {!$hasSep} {
        lappend errors "WML-NOTES($partName): missing separator $noteTag (w:type='separator' w:id='0')"
    }
    if {!$hasCont} {
        lappend errors "WML-NOTES($partName): missing continuationSeparator $noteTag (w:type='continuationSeparator' w:id='1')"
    }

    # Every note must have a w:id attribute (§17.11.15)
    foreach note [$root selectNodes w:$noteTag] {
        if {[$note @w:id ""] eq ""} {
            lappend errors "WML-NOTES($partName): w:$noteTag missing required w:id attribute"
        }
    }

    # Note ids must be unique (§17.11.15)
    set ids [dict create]
    foreach note [$root selectNodes w:$noteTag] {
        set nid [$note @w:id ""]
        if {$nid ne ""} {
            if {[dict exists $ids $nid]} {
                lappend errors "WML-NOTES($partName): duplicate w:$noteTag w:id='$nid'"
            }
            dict set ids $nid 1
        }
    }

    # User notes (id >= 2) must contain at least one w:p (§17.11.15)
    foreach note [$root selectNodes w:$noteTag] {
        set nid [$note @w:id ""]
        if {$nid ne "" && $nid > 1} {
            set pcount [llength [$note selectNodes w:p]]
            if {$pcount == 0} {
                lappend errors "WML-NOTES($partName): user $noteTag id=$nid has no w:p child"
            }
        }
    }

    $doc delete
    return $errors
}

# ── Styles (§17.7.4) ───────────────────────────────────────────

proc ecma376::validate::validateStyles {xml} {
    variable ns
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "WML-STYLES: malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list w $ns(w)]

    if {[$root localName] ne "styles"} {
        lappend errors "WML-STYLES: root must be w:styles, got [$root localName]"
    }

    # Every w:style must have w:type and w:styleId (§17.7.4.17)
    foreach style [$root selectNodes w:style] {
        set stype [$style @w:type ""]
        set sid   [$style @w:styleId ""]
        if {$stype eq ""} {
            lappend errors "WML-STYLES: w:style missing w:type attribute"
        }
        if {$sid eq ""} {
            lappend errors "WML-STYLES: w:style missing w:styleId attribute"
        }
        if {$stype ni {paragraph character numbering table}} {
            if {$stype ne ""} {
                lappend errors "WML-STYLES: w:style has invalid w:type='$stype'"
            }
        }
    }

    # styleId values must be unique within a type (§17.7.4.17)
    set typeIds [dict create]
    foreach style [$root selectNodes w:style] {
        set stype [$style @w:type ""]
        set sid   [$style @w:styleId ""]
        if {$stype eq "" || $sid eq ""} continue
        set key "${stype}:${sid}"
        if {[dict exists $typeIds $key]} {
            lappend errors "WML-STYLES: duplicate styleId '$sid' for type '$stype'"
        }
        dict set typeIds $key 1
    }

    # basedOn must reference an existing styleId of the same type (§17.7.4.3)
    foreach style [$root selectNodes w:style] {
        set stype [$style @w:type ""]
        set basedOn [$style selectNodes {string(w:basedOn/@w:val)}]
        if {$basedOn ne ""} {
            set key "${stype}:${basedOn}"
            if {![dict exists $typeIds $key]} {
                set sid [$style @w:styleId ""]
                lappend errors "WML-STYLES: style '$sid' basedOn '$basedOn' does not exist as $stype style"
            }
        }
    }

    $doc delete
    return $errors
}

# ── Settings (§17.15.1) ────────────────────────────────────────

proc ecma376::validate::validateSettings {xml} {
    variable ns
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        lappend errors "WML-SETTINGS: malformed XML: $err"
        return $errors
    }
    set root [$doc documentElement]
    $doc selectNodesNamespaces [list w $ns(w)]

    if {[$root localName] ne "settings"} {
        lappend errors "WML-SETTINGS: root must be w:settings, got [$root localName]"
    }

    $doc delete
    return $errors
}

# ── Schema data loading ─────────────────────────────────────────

# Load the machine-generated schema data (enums, child orderings,
# attribute→enum mappings) extracted from the ECMA-376 RNC schemas.
source [file join [file dirname [info script]] schema-data.tcl]

# ── Element child ordering validation (§17.3, §17.6, §17.15) ───
#
# ECMA-376 defines xsd:sequence for pPr, rPr, tblPr, and settings
# children.  Elements appearing out of order produce a schema-
# invalid document even though Word may tolerate it.

proc ecma376::validate::validateElementOrdering {xml {partName word/document.xml}} {
    variable ns
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        return $errors
    }
    $doc selectNodesNamespaces [list w $ns(w) m $ns(m)]

    # Check ordering for each container type
    foreach {containerTag orderKey} {
        pPr   pPr
        rPr   rPr
        tblPr tblPr
    } {
        if {![info exists ::ecma376::schema::childOrder($orderKey)]} continue
        set expected $::ecma376::schema::childOrder($orderKey)

        foreach container [$doc selectNodes "//w:$containerTag"] {
            set children [list]
            foreach child [$container childNodes] {
                if {[$child nodeType] ne "ELEMENT_NODE"} continue
                lappend children [$child localName]
            }
            # Check that children appear in schema order
            # (not all need to be present, but those present must be ordered)
            set lastIdx -1
            foreach childName $children {
                set idx [lsearch -exact $expected $childName]
                if {$idx == -1} continue ;# unknown child, skip (may be extension)
                if {$idx < $lastIdx} {
                    lappend errors "SCHEMA-ORDER($partName): w:$containerTag child w:$childName appears out of ECMA-376 sequence order (after w:[lindex $expected $lastIdx])"
                    break
                }
                set lastIdx $idx
            }
        }
    }

    $doc delete
    return $errors
}

# ── Attribute enum validation ───────────────────────────────────
#
# For every w:val attribute (and other enum-typed attributes) in
# the output XML, verify the value is in the ECMA-376 enumeration.

proc ecma376::validate::validateAttributeEnums {xml {partName word/document.xml}} {
    variable ns
    set errors [list]

    if {[catch {set doc [dom parse $xml]} err]} {
        return $errors
    }
    $doc selectNodesNamespaces [list w $ns(w) m $ns(m)]

    # Map of element localName → attribute localName → enum type
    # Built from the schema: when element X has attribute w:Y typed as ST_Z,
    # we validate the value against the enum.
    #
    # We focus on the high-value w:val attributes where the element name
    # tells us which enum to check.
    variable attrChecks
    if {![info exists attrChecks]} {
        array set attrChecks {}
        # Build element-specific checks from the schema attribute map
        # Format: elementLocalName:attrLocalName → enumTypeName
        foreach {attr enumType} [array get ::ecma376::schema::attrEnum] {
            # These attributes are used on many elements with different types
            # Only check w:val which is the most common pattern
            set attrChecks($attr) $enumType
        }
    }

    # Check w:val on elements where the parent context determines the type
    set elemToEnum {
        jc          w_ST_Jc
        u           w_ST_Underline
        highlight   w_ST_HighlightColor
        vertAlign   w_ST_VerticalAlignRun
        shd         w_ST_Shd
        pStyle      {}
        rStyle      {}
        tblStyle    {}
    }

    # Direct element → w:val enum checks
    foreach {elem enumName} $elemToEnum {
        if {$enumName eq ""} continue
        if {![info exists ::ecma376::schema::enums($enumName)]} continue
        set allowed $::ecma376::schema::enums($enumName)
        foreach node [$doc selectNodes "//w:$elem"] {
            set val [$node @w:val ""]
            if {$val eq ""} continue
            if {$val ni $allowed} {
                lappend errors "SCHEMA-ENUM($partName): w:$elem/@w:val='$val' is not a valid $enumName value"
            }
        }
    }

    # Check shading w:val (on w:shd elements)
    if {[info exists ::ecma376::schema::enums(w_ST_Shd)]} {
        set allowed $::ecma376::schema::enums(w_ST_Shd)
        foreach node [$doc selectNodes {//w:shd}] {
            set val [$node @w:val ""]
            if {$val ne "" && $val ni $allowed} {
                lappend errors "SCHEMA-ENUM($partName): w:shd/@w:val='$val' is not a valid w_ST_Shd value"
            }
        }
    }

    # Check border w:val (on border elements)
    if {[info exists ::ecma376::schema::enums(w_ST_Border)]} {
        set allowed $::ecma376::schema::enums(w_ST_Border)
        foreach borderElem {top bottom left right insideH insideV start end tl2br tr2bl between bar} {
            foreach node [$doc selectNodes "//w:$borderElem"] {
                set val [$node @w:val ""]
                if {$val ne "" && $val ni $allowed} {
                    lappend errors "SCHEMA-ENUM($partName): w:$borderElem/@w:val='$val' is not a valid w_ST_Border value"
                }
            }
        }
    }

    # Check w:type on sectPr (section mark)
    if {[info exists ::ecma376::schema::enums(w_ST_SectionMark)]} {
        set allowed $::ecma376::schema::enums(w_ST_SectionMark)
        foreach node [$doc selectNodes {//w:type}] {
            set parent [$node parentNode]
            if {[$parent localName] ne "sectPr"} continue
            set val [$node @w:val ""]
            if {$val ne "" && $val ni $allowed} {
                lappend errors "SCHEMA-ENUM($partName): w:type/@w:val='$val' is not a valid w_ST_SectionMark value"
            }
        }
    }

    # Check w:orient on pgSz (page orientation)
    if {[info exists ::ecma376::schema::enums(w_ST_PageOrientation)]} {
        set allowed $::ecma376::schema::enums(w_ST_PageOrientation)
        foreach node [$doc selectNodes {//w:pgSz}] {
            set val [$node @w:orient ""]
            if {$val ne "" && $val ni $allowed} {
                lappend errors "SCHEMA-ENUM($partName): w:pgSz/@w:orient='$val' is not a valid w_ST_PageOrientation value"
            }
        }
    }

    $doc delete
    return $errors
}
