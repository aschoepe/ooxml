#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2024, 2025 Rolf Ade, DE
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
package require ooxml
source [file join [file dir [info script]] ooxml-docx-lib.tcl]
package require ooxml::docx::lib

namespace eval ::ooxml::docx {

    namespace export docx

    # These are acually used as namespace variables
    variable prefixnslist
    variable properties
    variable staticDocx
    variable xmlns
    # These variables are only used in the namespace setup and
    # defined here only to avoid changes of possibly global variables
    # with the same name for Tcl < 9.0
    # (See https://core.tcl-lang.org/tips/doc/trunk/tip/278.md
    # "Fix Variable Name Resolution Quirks")
    variable BorderOpts
    variable borderOptions
    variable property
    variable option
    variable name
    variable xml
    variable tag

    array set xmlns {
        a http://schemas.openxmlformats.org/drawingml/2006/main
        ct http://schemas.openxmlformats.org/package/2006/content-types
        m  http://schemas.openxmlformats.org/officeDocument/2006/math
        mc http://schemas.openxmlformats.org/markup-compatibility/2006
        o urn:schemas-microsoft-com:office:office
        pic http://schemas.openxmlformats.org/drawingml/2006/picture
        r http://schemas.openxmlformats.org/officeDocument/2006/relationships
        rel http://schemas.openxmlformats.org/package/2006/relationships
        sl http://schemas.openxmlformats.org/schemaLibrary/2006/main
        v urn:schemas-microsoft-com:vml
        w http://schemas.openxmlformats.org/wordprocessingml/2006/main
        w10 urn:schemas-microsoft-com:office:word
        w14 http://schemas.microsoft.com/office/word/2010/wordml
        wp http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
        wp14 http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
        wpg http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
        wps http://schemas.microsoft.com/office/word/2010/wordprocessingShape
    }
    set prefixnslist [array get xmlns]
    # Most WordprocessingML elements have a sequence content model. So
    # the basic rule is to keep the order of the options below as is
    # (and to care that new options are inserted at the right place).
    # Exceptions are locally noted.

    set properties(abstractNumStyle) {
        -numberFormat {w:numFmt ST_NumberFormat}
        -levelRestart {w:lvlRestart ST_DecimalNumber}
        -levelText {w:lvlText NoCheck}
        -align {w:lvlJc ST_Jc}
    }

    set properties(cell1) {
        -cellWidth {w:tcW {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -span {w:gridSpan ST_DecimalNumber}
        -hspan {w:hMerge ST_Merge}
        -vspan {w:vMerge ST_Merge}
    }

    set properties(cell2) {
        -shading {w:shd {
            {type val} {ST_Shd !}
            color ST_HexColor
            fill ST_HexColor}}
    }

    set properties(cell3) {
        -textDirection {w:textDirection ST_TextDirection}
        -tcFitText {w:tcFitText CT_OnOff}
        -vAlign {w:vAlign ST_VerticalJc}
        -hideMark {w:hideMark CT_OnOff}
    }

    set properties(cellMargins) {
        -cellMarginTop {w:top {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -cellMarginStart {w:start {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -cellMarginLeft {w:left {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -cellMarginBottom {w:bottom {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -cellMarginEnd {w:end {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -cellMarginRight {w:right {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
    }

    set properties(paragraph1) {
        -pstyle {w:pStyle PStyle}
        -keepNext {w:keepNext CT_OnOff}
        -keepLines {w:keepLines CT_OnOff}
        -pageBreakBefore {w:pageBreakBefore CT_OnOff}
        -textframe {w:framePr {
            {width w} ST_TwipsMeasure
            {height h} ST_TwipsMeasure
            wrap ST_Wrap
            vAnchor ST_Anchor
            hAnchor ST_Anchor
            xAlign ST_XAlign
            yAlign ST_YAlign
            dropCap ST_DropCap
            lines ST_DecimalNumber
            vSpace ST_TwipsMeasure
            hSpace ST_TwipsMeasure
            hrule ST_HeightRule
            anchorLock CT_OnOff}}
        -widowControl {w:widowControl CT_OnOff}
        +w:numPr {
            -level {w:ilvl ST_DecimalNumber}
            -numberingStyle {w:numId ST_DecimalNumber}
        }
        -suppressLineNumbers {w:suppressLineNumbers CT_OnOff}
    }

    set properties(paragraph2) {
        -shading {w:shd {
            {type val} {ST_Shd !}
            color ST_HexColor
            fill ST_HexColor}}
    }

    set properties(paragraph3) {
        -bidi {w:bidi CT_OnOff}
        -spacing {w:spacing {
            after ST_TwipsMeasure
            before ST_TwipsMeasure
            line ST_TwipsMeasure
            lineRule ST_LineSpacingRule}}
        -indentation {w:ind {
            end ST_SignedTwipsMeasure
            endChars ST_DecimalNumber
            firstLine ST_SignedTwipsMeasure
            firstLineChars ST_DecimalNumber
            hanging ST_SignedTwipsMeasure
            hangingChars ST_DecimalNumber
            start ST_SignedTwipsMeasure
            startChars ST_DecimalNumber}}
        -align {w:jc ST_Jc}
    }

    set properties(row) {
        -wBefore {w:wBefore {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -wAfter {w:wAfter {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -cantSplit {w:cantSplit CT_OnOff}
        -rowHeight {w:trHeight {
            {value val} ST_TwipsMeasure
            hRule ST_HeightRule}}
        -headerrow {w:tblHeader CT_OnOff}
        -cellSpacing {w:tblCellSpacing {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -align {w:jc ST_JcTable}
    }

    # According to the specification the children elements of rPr are
    # in unspecified order. Since Miguel uses an in this case limited
    # validator which wrongly insists on a certain order we follow
    # this for now.
    set properties(run) {
        -cstyle {w:rStyle RStyle}
        -font {w:rFonts NoCheck RFonts}
        -bold {{w:b w:bCs} CT_OnOff}
        -italic {{w:i w:iCs} CT_OnOff}
        -strike {w:strike CT_OnOff}
        -dstrike {w:dstrike CT_OnOff}
        -emboss {w:emboss CT_OnOff}
        -noProof {w:noProof CT_OnOff}
        -color {w:color ST_HexColor}
        -fontsize {{w:sz w:szCs} ST_HpsMeasure}
        -highlight {w:highlight ST_HighlightColor}
        -underline {w:u ST_Underline}
        -verticalAlign {w:vertAlign ST_VerticalAlignRun}
        -rtl {w:rtl CT_OnOff}
    }

    set properties(sectionsetup1) {
        +w:footnotePr {
            -fn_pos {w:pos ST_FtnPos}
            -fn_numFmt {w:numFmt ST_NumberFormat}
            -fn_numStart {w:numStart ST_DecimalNumber}
            -fn_numRestart {w:numRestart ST_RestartNumber}
        }
        +w:endnotePr {
            -en_pos {w:pos ST_EdnPos}
            -en_numFmt {w:numFmt ST_NumberFormat}
            -en_numStart {w:numStart ST_DecimalNumber}
            -en_numRestart {w:numRestart ST_RestartNumber}
        }
        -sectionType {w:type ST_SectionMark}
        -sizeAndOrientation {w:pgSz {
            {height h} ST_TwipsMeasure
            {orientation orient} ST_PageOrientation
            {width w} ST_TwipsMeasure}}
        -margins {w:pgMar {
            bottom {ST_TwipsMeasure ! 1134}
            footer {ST_TwipsMeasure ! 0}
            gutter {ST_TwipsMeasure ! 0}
            header {ST_TwipsMeasure ! 0}
            left {ST_TwipsMeasure ! 1134}
            right {ST_TwipsMeasure ! 1134}
            top {ST_TwipsMeasure ! 1134}}}
        -paperSource {w:paperSrc {
            first ST_DecimalNumber
            other ST_DecimalNumber
        }}
    }

    set properties(sectionsetup2) {
        -pageNumbering {w:pgNumType {
            chapSep ST_ChapterSep
            chapStyle ST_DecimalNumber
            fmt ST_NumberFormat
            start ST_DecimalNumber
        }}
        -vAlign {w:vAlign ST_VerticalJc}
        -docGrid {w:docGrid {
            type ST_DocGrid
            linePitch ST_DecimalNumber
            charSpace ST_DecimalNumber
        }}
    }

    set properties(settings) {
        -writeProtection {w:writeProtection {
            recommended CT_OnOff
            algorithmName NoCheck
            hashValue XSD_base64Binary
            saltValue XSD_base64Binary
            spinCount ST_DecimalNumber
            cryptProviderType ST_CryptProv
            cryptAlgorithmClass ST_AlgClass
            cryptAlgorithmType ST_AlgType
            cryptAlgorithmSid ST_DecimalNumber
            cryptSpinCount ST_DecimalNumber
            cryptProvider NoCheck
            algIdExt XSD_hexBinary
            algIdExtSource NoCheck
            cryptProviderTypeExt XSD_hexBinary
            cryptProviderTypeExtSource NoCheck
            hash XSD_base64Binary
            salt XSD_base64Binary
        }}
        -view {w:view ST_View}
        -zoom {w:zoom {
            {type val} ST_Zoom
            percent ST_DecimalNumberOrPercent
        }}
        -removePersonalInformation {w:removePersonalInformation CT_OnOff}
        -removeDateAndTime {w:removeDateAndTime CT_OnOff}
        -doNotDisplayPageBoundaries {w:doNotDisplayPageBoundaries CT_OnOff}
        -displayBackgroundShape {w:displayBackgroundShape CT_OnOff}
        -printPostScriptOverText {w:printPostScriptOverText CT_OnOff}
        -printFractionalCharacterWidth {w:printFractionalCharacterWidth CT_OnOff}
        -printFormsData {w:printFormsData CT_OnOff}
        -embedTrueTypeFonts {w:embedTrueTypeFonts CT_OnOff}
        -embedSystemFonts {w:embedSystemFonts CT_OnOff}
        -saveSubsetFonts {w:saveSubsetFonts CT_OnOff}
        -saveFormsData {w:saveFormsData CT_OnOff}
        -mirrorMargins {w:mirrorMargins CT_OnOff}
        -alignBordersAndEdges {w:alignBordersAndEdges CT_OnOff}
        -bordersDoNotSurroundHeader {w:bordersDoNotSurroundHeader CT_OnOff}
        -bordersDoNotSurroundFooter {w:bordersDoNotSurroundFooter CT_OnOff}
        -gutterAtTop {w:gutterAtTop CT_OnOff}
        -hideSpellingErrors {w:hideSpellingErrors CT_OnOff}
        -hideGrammaticalErrors {w:hideGrammaticalErrors CT_OnOff}
        / w:activeWritingStyle
        -proofState {w:proofState {
            spelling ST_Proof
            grammar ST_Proof
        }}
        -formsDesign {w:formsDesign CT_OnOff}
        -linkStyles {w:linkStyles CT_OnOff}
        / w:stylePaneFormatFilter
        / w:stylePaneSortMethod
        -documentType {w:documentType NoCheck}
        / w:revisionView
        -trackRevisions {w:trackRevisions CT_OnOff}
        -doNotTrackMoves {w:doNotTrackMoves CT_OnOff}
        -doNotTrackFormatting {w:doNotTrackFormatting CT_OnOff}
        -documentProtection {w:documentProtection {
            edit ST_DocProtect
            formatting CT_OnOff
            enforcement CT_OnOff
        }}
        -autoFormatOverride {w:autoFormatOverride CT_OnOff}
        -styleLockTheme {w:styleLockTheme CT_OnOff}
        -styleLockQFSet {w:styleLockQFSet CT_OnOff}
        -defaultTabStop {w:defaultTabStop ST_TwipsMeasure}
        -autoHyphenation {w:autoHyphenation CT_OnOff}
        -consecutiveHyphenLimit {w:consecutiveHyphenLimit ST_DecimalNumber}
        -hyphenationZone {w:hyphenationZone ST_TwipsMeasure}
        -doNotHyphenateCaps {w:doNotHyphenateCaps CT_OnOff}
        -showEnvelope {w:showEnvelope CT_OnOff}
        -summaryLength {w:summaryLength ST_DecimalNumberOrPercent}
        -clickAndTypeStyle {w:clickAndTypeStyle NoCheck}
        -defaultTableStyle {w:defaultTableStyle NoCheck}
        -evenAndOddHeaders {w:evenAndOddHeaders CT_OnOff}
        -bookFoldRevPrinting {w:bookFoldRevPrinting CT_OnOff}
        -bookFoldPrinting {w:bookFoldPrinting CT_OnOff}
        -bookFoldPrintingSheets {w:bookFoldPrintingSheets ST_DecimalNumber}
        -drawingGridHorizontalSpacing {w:drawingGridHorizontalSpacing ST_TwipsMeasure}
        -drawingGridVerticalSpacing {w:drawingGridVerticalSpacing ST_TwipsMeasure}
        -displayHorizontalDrawingGridEvery {w:displayHorizontalDrawingGridEvery ST_DecimalNumber}
        -displayVerticalDrawingGridEvery {w:displayVerticalDrawingGridEvery ST_DecimalNumber}
        -doNotUseMarginsForDrawingGridOrigin {w:doNotUseMarginsForDrawingGridOrigin CT_OnOff}
        -drawingGridHorizontalOrigin {w:drawingGridHorizontalOrigin ST_TwipsMeasure}
        -drawingGridVerticalOrigin {w:drawingGridVerticalOrigin ST_TwipsMeasure}
        -doNotShadeFormData {w:doNotShadeFormData CT_OnOff}
        -noPunctuationKerning {w:noPunctuationKerning CT_OnOff}
        -characterSpacingControl {w:characterSpacingControl ST_CharacterSpacing}
        -printTwoOnOne {w:printTwoOnOne CT_OnOff}
        -strictFirstAndLastChars {w:strictFirstAndLastChars CT_OnOff}
        / w:noLineBreaksAfter
        / w:noLineBreaksBefore
        -savePreviewPicture {w:savePreviewPicture CT_OnOff}
        -doNotValidateAgainstSchema {w:doNotValidateAgainstSchema CT_OnOff}
        -saveInvalidXml {w:saveInvalidXml CT_OnOff}
        -ignoreMixedContent {w:ignoreMixedContent CT_OnOff}
        -alwaysShowPlaceholderText {w:alwaysShowPlaceholderText CT_OnOff}
        -doNotDemarcateInvalidXml {w:doNotDemarcateInvalidXml CT_OnOff}
        -saveXmlDataOnly {w:saveXmlDataOnly CT_OnOff}
        -useXSLTWhenSaving {w:useXSLTWhenSaving CT_OnOff}
        -saveThroughXslt {w:saveThroughXslt CT_OnOff}
        -showXMLTags {w:showXMLTags CT_OnOff}
        -alwaysMergeEmptyNamespace {w:alwaysMergeEmptyNamespace CT_OnOff}
        -updateFields {w:updateFields CT_OnOff}
        / w:hdrShapeDefaults
        / w:footnotePr
        / w:endnotePr
        / w:compat
        / w:docVars
        / w:rsids
        / m:mathPr
        -attachedSchema {w:attachedSchema NoCheck}
        / w:themeFontLang
        / w:clrSchemeMapping
        -doNotIncludeSubdocsInStats {w:doNotIncludeSubdocsInStats CT_OnOff}
        -doNotAutoCompressPictures {w:doNotAutoCompressPictures CT_OnOff}
        / w:forceUpgrade
        / w:captions
        / w:readModeInkLockDown
        / w:smartTagType
        / sl:schemaLibrary
        / w:shapeDefaults
        -doNotEmbedSmartTags {w:doNotEmbedSmartTags CT_OnOff}
        -decimalSymbol {w:decimalSymbol NoCheck}
        -listSeparator {w:listSeparator NoCheck}
    }

    set properties(table1) {
        -style {w:tblStyle TStyle}
        -width {w:tblW {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -align {w:jc ST_JcTable}
    }

    set properties(table2) {
        -layout {w:tblLayout {
            type ST_TblLayoutType}}
    }

    set properties(table3) {
        -look {w:tblLook {
            firstRow CT_OnOff
            lastRow CT_OnOff
            firstColumn CT_OnOff
            lastColumn CT_OnOff
            noHBand CT_OnOff
            noVBand CT_OnOff
        }}
        -caption {w:tblCaption NoCheck}
    }

    set properties(xfrm) {
        -dimension {a:ext {
            {width -cx} ST_Emu
            {height -cy} ST_Emu
        }}
    }

    # This creates
    # properties(paragraphBorders)
    # properties(sectionBorders)
    # properties(tableBorders)
    # properties(cellBorders)
    set BorderOpts {
        color ST_HexColor
        {borderwidth sz} ST_EighthPointMeasure
        space ST_PointMeasure
        {type val} ST_Border
    }
    foreach {property borderOptions} {
        paragraphBorders {top left bottom right between bar}
        sectionBorders {top left bottom right}
        tableBorders {top start left bottom end right insideH insideV}
        cellBorders {top start left bottom end right insideH insideV tl2br tr2bl}
    } {
        foreach option $borderOptions {
            lappend properties($property) \
                -${option}Border [list w:$option $BorderOpts]
        }
    }

    foreach {name xml} {
        [Content_Types].xml {
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
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
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
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
            </cp:coreProperties>
        }
        word/fontTable.xml {
            <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <w:font w:name="Times New Roman">
                    <w:charset w:val="00"/>
                    <w:family w:val="roman"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Symbol">
                    <w:charset w:val="02"/>
                    <w:family w:val="roman"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Arial">
                    <w:charset w:val="00"/>
                    <w:family w:val="swiss"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Liberation Serif">
                    <w:altName w:val="Times New Roman"/>
                    <w:charset w:val="01"/>
                    <w:family w:val="roman"/>
                    <w:pitch w:val="variable"/>
                </w:font>
                <w:font w:name="Liberation Sans">
                    <w:altName w:val="Arial"/>
                    <w:charset w:val="01"/>
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
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>
        }
    } {
        set staticDocx($name) $xml
    }

    foreach tag {
        w:abstractNum w:abstractNumId w:active w:activeRecord
        w:activeWritingStyle w:addressFieldName
        w:adjustLineHeightInTable w:adjustRightInd w:alias w:aliases
        w:alignBordersAndEdges w:alignTablesRowByRow w:allowPNG
        w:allowSpaceOfSameStyleInTable w:altChunk w:altChunkPr
        w:altName w:alwaysMergeEmptyNamespace
        w:alwaysShowPlaceholderText w:annotationRef
        w:applyBreakingRules w:attachedSchema w:attachedTemplate
        w:attr w:autoCaption w:autoCaptions
        w:autofitToFirstFixedWidthCell w:autoFormatOverride
        w:autoHyphenation w:autoRedefine w:autoSpaceDE w:autoSpaceDN
        w:autoSpaceLikeWord95 w:b w:background
        w:balanceSingleByteDoubleByteWidth w:bar w:basedOn w:bCs w:bdo
        w:bdr w:behavior w:behaviors w:between w:bibliography w:bidi
        w:bidiVisual w:blockQuote w:body w:bodyDiv w:bookFoldPrinting
        w:bookFoldPrintingSheets w:bookFoldRevPrinting w:bookmarkEnd
        w:bookmarkStart w:bordersDoNotSurroundFooter
        w:bordersDoNotSurroundHeader w:bottom w:br w:cachedColBalance
        w:calcOnExit w:calendar w:cantSplit w:caps w:caption
        w:captions w:category w:cellDel w:cellIns w:cellMerge
        w:characterSpacingControl w:charset w:checkBox w:checked
        w:checkErrors w:citation w:clickAndTypeStyle
        w:clrSchemeMapping w:cnfStyle w:col w:colDelim w:color w:cols
        w:column w:comboBox w:comment w:commentRangeEnd
        w:commentRangeStart w:commentReference w:comments w:compat
        w:compatSetting w:connectString w:consecutiveHyphenLimit
        w:contentPart w:contextualSpacing w:continuationSeparator
        w:control w:convMailMergeEsc w:cr w:cs w:customXml
        w:customXmlDelRangeEnd w:customXmlDelRangeStart
        w:customXmlInsRangeEnd w:customXmlInsRangeStart
        w:customXmlMoveFromRangeEnd w:customXmlMoveFromRangeStart
        w:customXmlMoveToRangeEnd w:customXmlMoveToRangeStart
        w:customXmlPr w:dataBinding w:dataSource w:dataType w:date
        w:dateFormat w:dayLong w:dayShort w:ddList w:decimalSymbol
        w:default w:defaultTableStyle w:defaultTabStop w:del
        w:delInstrText w:delText w:description w:destination w:dir
        w:dirty w:displayBackgroundShape w:displayHangulFixedWidth
        w:displayHorizontalDrawingGridEvery
        w:displayVerticalDrawingGridEvery w:div w:divBdr w:divId
        w:divs w:divsChild w:docDefaults w:docGrid w:docPart
        w:docPartBody w:docPartCategory w:docPartGallery w:docPartList
        w:docPartObj w:docPartPr w:docParts w:docPartUnique w:document
        w:documentProtection w:documentType w:docVar w:docVars
        w:doNotAutoCompressPictures w:doNotAutofitConstrainedTables
        w:doNotBreakConstrainedForcedTable w:doNotBreakWrappedTables
        w:doNotDemarcateInvalidXml w:doNotDisplayPageBoundaries
        w:doNotEmbedSmartTags w:doNotExpandShiftReturn
        w:doNotHyphenateCaps w:doNotIncludeSubdocsInStats
        w:doNotLeaveBackslashAlone w:doNotOrganizeInFolder
        w:doNotRelyOnCSS w:doNotSaveAsSingleFile w:doNotShadeFormData
        w:doNotSnapToGridInCell w:doNotSuppressBlankLines
        w:doNotSuppressIndentation w:doNotSuppressParagraphBorders
        w:doNotTrackFormatting w:doNotTrackMoves
        w:doNotUseEastAsianBreakRules
        w:doNotUseHTMLParagraphAutoSpacing
        w:doNotUseIndentAsNumberingTabStop w:doNotUseLongFileNames
        w:doNotUseMarginsForDrawingGridOrigin
        w:doNotValidateAgainstSchema w:doNotVertAlignCellWithSp
        w:doNotVertAlignInTxbx w:doNotWrapTextWithPunct w:drawing
        w:drawingGridHorizontalOrigin w:drawingGridHorizontalSpacing
        w:drawingGridVerticalOrigin w:drawingGridVerticalSpacing
        w:dropDownList w:dstrike w:dynamicAddress w:eastAsianLayout
        w:effect w:em w:embedBold w:embedBoldItalic w:embedItalic
        w:embedRegular w:embedSystemFonts w:embedTrueTypeFonts
        w:emboss w:enabled w:encoding w:end w:endnote
        w:endnoteRef w:endnoteReference w:endnotes w:entryMacro
        w:equation w:evenAndOddHeaders w:exitMacro w:family w:ffData
        w:fHdr w:fieldMapData w:fitText w:flatBorders w:fldChar
        w:fldData w:fldSimple w:font w:fonts w:footerReference
        w:footnote w:footnoteLayoutLikeWW8 w:footnoteRef
        w:footnoteReference w:footnotes w:forceUpgrade
        w:forgetLastTabAlignment w:format w:formProt w:formsDesign
        w:frame w:frameLayout w:framePr w:frameset w:framesetSplitbar
        w:ftr w:gallery w:glossaryDocument w:gridAfter w:gridBefore
        w:gridCol w:gridSpan w:group w:growAutofit w:guid
        w:gutterAtTop w:hdr w:hdrShapeDefaults w:header
        w:headerReference w:headers w:headerSource w:helpText w:hidden
        w:hideGrammaticalErrors w:hideMark w:hideSpellingErrors
        w:highlight w:hMerge w:hps w:hpsBaseText w:hpsRaise
        w:hyperlink w:hyphenationZone w:i w:iCs w:id
        w:ignoreMixedContent w:ilvl w:imprint w:ind w:ins w:insideH
        w:insideV w:instrText w:isLgl w:jc w:keepLines w:keepNext
        w:kern w:kinsoku w:label w:lang w:lastRenderedPageBreak
        w:latentStyles w:layoutRawTableWidth w:layoutTableRowsApart
        w:left w:legacy w:lid w:lineWrapLikeWord6 w:link
        w:linkedToFile w:linkStyles w:linkToQuery w:listEntry
        w:listItem w:listSeparator w:lnNumType w:lock w:locked
        w:longDesc w:lsdException w:lvl w:lvlJc w:lvlOverride
        w:lvlPicBulletId w:lvlRestart w:lvlText w:mailAsAttachment
        w:mailMerge w:mailSubject w:mainDocumentType w:mappedName
        w:marBottom w:marH w:marLeft w:marRight w:marTop w:marW
        w:matchSrc w:maxLength w:mirrorIndents w:mirrorMargins
        w:monthLong w:monthShort w:moveFrom w:moveFromRangeEnd
        w:moveFromRangeStart w:moveTo w:moveToRangeEnd
        w:moveToRangeStart w:movie w:multiLevelType w:mwSmallCaps
        w:name w:next w:noBorder w:noBreakHyphen w:noColumnBalance
        w:noEndnote w:noExtraLineSpacing w:noLeading
        w:noLineBreaksAfter w:noLineBreaksBefore w:noProof
        w:noPunctuationKerning w:noResizeAllowed w:noSpaceRaiseLower
        w:noTabHangInd w:notTrueType w:noWrap w:nsid w:num w:numbering
        w:numberingChange w:numFmt w:numId w:numIdMacAtCleanup
        w:numPicBullet w:numRestart w:numStart w:numStyleLink w:object
        w:objectEmbed w:objectLink w:odso w:oMath w:optimizeForBrowser
        w:outline w:outlineLvl w:overflowPunct w:p w:pageBreakBefore
        w:panose1 w:paperSrc w:permEnd w:permStart w:personal
        w:personalCompose w:personalReply w:pgBorders w:pgMar w:pgNum
        w:pgNumType w:pgSz w:pict w:picture w:pitch w:pixelsPerInch
        w:placeholder w:pos w:position w:pPrChange w:pPrDefault
        w:printBodyTextBeforeHeader w:printColBlack w:printerSettings
        w:printFormsData w:printFractionalCharacterWidth
        w:printPostScriptOverText w:printTwoOnOne w:proofErr
        w:proofState w:pStyle w:ptab w:qFormat w:query w:r
        w:readModeInkLockDown w:recipientData w:recipients w:relyOnVML
        w:removeDateAndTime w:removePersonalInformation w:result
        w:revisionView w:rFonts w:richText w:right w:rPrChange
        w:rPrDefault w:rsid w:rsidRoot w:rsids w:rStyle w:rt w:rtl
        w:rtlGutter w:ruby w:rubyAlign w:rubyBase w:rubyPr
        w:saveFormsData w:saveInvalidXml w:savePreviewPicture
        w:saveSmartTagsAsXml w:saveSubsetFonts w:saveThroughXslt
        w:saveXmlDataOnly w:scrollbar w:sdt w:sdtContent w:sdtEndPr
        w:sdtPr w:sectPr w:sectPrChange w:selectFldWithFirstOrLastChar
        w:semiHidden w:separator w:settings w:shadow w:shapeDefaults
        w:shapeLayoutLikeWW8 w:shd w:showBreaksInFrames w:showEnvelope
        w:showingPlcHdr w:showXMLTags w:sig w:size w:sizeAuto
        w:smallCaps w:smartTag w:smartTagPr w:smartTagType
        w:snapToGrid w:softHyphen w:sourceFileName w:spaceForUL
        w:spacing w:spacingInWholePoints w:specVanish
        w:splitPgBreakAndParaMark w:src w:start w:startOverride
        w:statusText w:storeMappedDataAs w:strictFirstAndLastChars
        w:strike w:style w:styleLink w:styleLockQFSet w:styleLockTheme
        w:stylePaneFormatFilter w:stylePaneSortMethod w:styles
        w:subDoc w:subFontBySize w:suff w:summaryLength
        w:suppressAutoHyphens w:suppressBottomSpacing
        w:suppressLineNumbers w:suppressOverlap
        w:suppressSpacingAtTopOfPage w:suppressSpBfAfterPgBrk
        w:suppressTopSpacing w:suppressTopSpacingWP
        w:swapBordersFacingPages w:sym w:sz w:szCs w:t w:tab
        w:tabIndex w:table w:tag w:targetScreenSz w:tbl w:tblCaption
        w:tblCellSpacing w:tblDescription w:tblGrid w:tblGridChange
        w:tblHeader w:tblInd w:tblLayout w:tblLook w:tblOverlap w:tblPr
        w:tblpPr w:tblPrChange w:tblPrEx w:tblPrExChange w:tblStyle
        w:tblStyleColBandSize w:tblStylePr w:tblStyleRowBandSize
        w:tblW w:tc w:tcFitText w:tcPrChange w:tcW w:temporary
        w:text w:textAlignment w:textboxTightWrap w:textDirection
        w:textInput w:themeFontLang w:title w:titlePg w:tl2br w:tmpl
        w:top w:topLinePunct w:tr w:tr2bl w:trackRevisions w:trHeight
        w:trPrChange w:truncateFontHeightsLikeWP6 w:txbxContent
        w:type w:types w:u w:udl w:uiPriority w:ulTrailSpace
        w:underlineTabInNumList w:unhideWhenUsed w:uniqueTag
        w:updateFields w:useAltKinsokuLineBreakRules
        w:useAnsiKerningPairs w:useFELayout w:useNormalStyleForList
        w:usePrinterMetrics w:useSingleBorderforContiguousCells
        w:useWord97LineBreakRules w:useWord2002TableStyleRules
        w:useXSLTWhenSaving w:vAlign w:vanish w:vertAlign w:view
        w:viewMergedData w:vMerge w:w w:wAfter w:wBefore w:webHidden
        w:webSettings w:widowControl w:wordWrap w:wpJustification
        w:wpSpaceWidth w:wrapTrailSpaces w:writeProtection w:yearLong
        w:yearShort w:zoom
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(w) elementNode Tag_$tag
    }

    foreach tag {
        w:footnotePr w:endnotePr w:numPr w:pBdr w:pPr w:rPr w:tabs
        w:tblBorders w:tblCellMar w:tcBorders w:tcMar w:tcPr
        w:trPr
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(w) -notempty elementNode Tag_$tag
    }

    foreach tag {
        Relationships Relationship
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(rel) elementNode Tag_$tag
    }

    foreach tag {
        wp:align wp:anchor wp:cNvGraphicFramePr wp:docPr
        wp:effectExtent wp:extent wp:inline wp:positionH wp:positionV
        wp:posOffset wp:simplePos wp:wrapNone wp:wrapSquare
        wp:wrapTopAndBottom
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(wp) elementNode Tag_$tag
    }

    foreach tag {
        a:avLst a:blip a:ext a:fillRect a:graphic a:graphicData
        a:graphicFrameLocks a:off a:picLocks a:prstGeom a:stretch
        a:xfrm
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(a) elementNode Tag_$tag
    }

    foreach tag {
        pic:pic pic:blipFill pic:nvPicPr pic:cNvPr pic:cNvPicPr pic:spPr
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(pic) elementNode Tag_$tag
    }

    foreach tag {
        wps:bodyPr wps:cNvSpPr wps:spPr wps:txbx wps:wsp
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(wps) elementNode Tag_$tag
    }

    foreach tag {
        Default Override
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(ct) elementNode Tag_$tag
    }


    dom createNodeCmd textNode Text
    namespace export Tag_* Text
}

oo::class create ooxml::docx::docx {

    constructor { args } {
        my variable docs
        my variable body
        my variable context
        my variable binparts
        my variable ignorable
        my variable setuproot
        my variable pagesetup
        my variable sectionsetup
        my variable impPagesetup
        my variable tablecontext
        my variable bookmarks
        my variable id
        my variable valPrefix

        variable ::ooxml::docx::xmlns
        variable ::ooxml::docx::staticDocx
        variable ::ooxml::docx::prefixnslist

        namespace import ::ooxml::docx::lib::*
        namespace import ::ooxml::docx::Tag_*  ::ooxml::docx::Text

        # Setup defaults/initial values
        #
        # Sensible default according to field
        set ignorable "w14 wp14 wpg wps"
        set context document
        set valPrefix w
        set pagesetup ""
        set sectionsetup ""
        set impPagesetup ""
        set tablecontext ""
        # Since the "general page setup" (WordprocessingML does not
        # really have a concept for that) is a child of w:body after
        # the last paragraph it seems handy to not reallyq insert that
        # into the tree at the moment the user calls pagesetup but
        # just before serializing - as the user consecutive add
        # content we can just append to the w:body without looking at
        # every place if there is already a page setup child and we
        # have to insert new content before that element. The current
        # pagesetup/sectionsetup user definition is stored in
        # according variables to be applied later. To give the user
        # error feedback at the place he provides commands an
        # auxiliary document is used and the definition is "tested"
        # with that.
        set setupdoc [dom createDocumentNS $xmlns(w) w:umbrella]
        set setuproot [$setupdoc documentElement]
        foreach ns {o m r v w w10 wp wps wpg mc wp14 w14} {
            $setuproot setAttributeNS "" xmlns:$ns $xmlns($ns)
        }

        if {[llength $args] % 2} {
            # User gave a docx file name to start from
            set startdocx [lindex $args end]
            set args [lrange $args 0 end-1]
            my InitFromDocx $startdocx
        } else {
            # Create a docx from scratch
            # Ensure that we can use our prefixes for XPath queries
            foreach auxFile [array names staticDocx] {
                set docs($auxFile) [dom parse $staticDocx($auxFile)]
                $docs($auxFile) selectNodesNamespaces $prefixnslist
            }
            set document [dom createDocumentNS $xmlns(w) w:document]
            set docs(word/document.xml) $document
            $docs(word/document.xml) selectNodesNamespaces $prefixnslist
            $document documentElement root
            foreach ns {o m r v w10 wp wps wpg mc wp14 w14} {
                $root setAttributeNS "" xmlns:$ns $xmlns($ns)
            }
            $root appendFromScript Tag_w:body
            set body [$root firstChild]
            my Ignorable
        }
        my configure {*}$args
    }

    destructor {
        my variable docs
        my variable setuproot

        foreach part [array names docs] {
            $docs($part) delete
        }
        [$setuproot ownerDocument] delete
    }

    method Add2Content_Types {file} {
        my variable docs
        variable ::ooxml::docx::xmlns

        if {[string index $file 0] ne "/"} {
            error "file paths to be added to \[ContentType\].xml must start\
                   with a slash (/)"
        }
        set ctRoot [$docs(\[Content_Types\].xml) documentElement]
        if {[$ctRoot selectNodes {count(ct:Override[@PartName=$file])}] > 0} {
            return
        }
        if {[string range $file 0 11] eq "/word/media/"} {
            set suffix [string tolower [string range [file extension $file] 1 end]]
            if {[$ctRoot selectNodes {
                count(ct:Default[@Extension=$suffix])
            }] > 0} {
                return
            }
        }
        switch -glob $file {
            "/_rels/.rels" {
                set type "application/vnd.openxmlformats-package.relationships+xml"
            }
            "/docProps/core.xml" {
                set type "application/vnd.openxmlformats-package.core-properties+xml"
            }
            "/docProps/app.xml" {
                set type "application/vnd.openxmlformats-officedocument.extended-properties+xml"
            }
            "/word/_rels/document.xml.rels" {
                set type "application/vnd.openxmlformats-package.relationships+xml"
            }
            "/word/comments.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
            }
            "/word/document.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            }
            "/word/endnotes.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
            }
            "/word/fontTable.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
            }
            "/word/footnotes.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
            }
            "/word/footer*.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
            }
            "/word/header*.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
            }
            "/word/media/*" {
                set contentType [MimeType $suffix]
                $ctRoot insertBeforeFromScript {
                    Tag_Default Extension $suffix ContentType $contentType
                } [$ctRoot selectNodes {ct:Override[1]}]
                return
            }
            "/word/numbering.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
            }
            "/word/settings.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
            }
            "/word/styles.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
            }
            "/word/stylesWithEffects.xml" {
                set type "application/vnd.ms-word.stylesWithEffects+xml"
            }
            "/word/theme/theme1.xml" {
                set type "application/vnd.openxmlformats-officedocument.theme+xml"
            }
            "/word/webSettings.xml" {
                set type "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
            }
            default {
                error "cannot find the ContentType for '$file'"
            }
        }
        $ctRoot appendFromScript {
            Tag_Override PartName $file ContentType $type
        }
    }

    method Add2Relationships {type target} {
        my variable docs
        my variable context
        variable ::ooxml::docx::xmlns

        set thisfile "word/_rels/$context.xml.rels"
        if {![info exists docs($thisfile)]} {
            set docs($thisfile) [dom parse {
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
            }]
        }
        set relsRoot [$docs($thisfile) documentElement]

        # If there is already an entry for this type and target just
        # return the rId
        set fqtype http://schemas.openxmlformats.org/officeDocument/2006/relationships/$type
        set result [$relsRoot selectNodes -namespaces "rel $xmlns(rel)" {
            string(rel:Relationship[@Type=$fqtype and @Target=$target]/@Id)
        }]
        if {$result ne ""} {
            return $result
        }

        # Find the maximum numeric part of any existing rId in this relationships part
        set nr 0
        foreach rel [$relsRoot childNodes] {
            set id [$rel @Id ""]
            if {[regexp {^rId([0-9]+)$} $id -> n]} {
                if {$n > $nr} { set nr $n }
            }
        }
        incr nr
        set attlist [list \
             Id rId$nr \
             Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/$type \
             Target $target]
        if {$type eq "hyperlink"} {
            lappend attlist TargetMode External
        }
        $relsRoot appendFromScript {
            Tag_Relationship {*}$attlist
        }
        if {$type eq "hyperlink"} {
            return rId$nr
        }
        my Add2Content_Types "/word/$target"
        return rId$nr
    }

    method Anchor {name} {
        my Prepare 1
        # Defaults
        array set anchorAtts {
            behindDoc 0
            distT 0
            distB 0
            distL 0
            distR 0
            hidden 0
            locked 0
            layoutInCell 0
            allowOverlap 1
            relativeHeight 1
            simplePos 0
        }
        # Updated by optional user provided values
        array set anchorAtts [my CheckedAttlist [my EatOption -anchorData] {
            -behindDoc CT_Boolean
            -distT ST_Emu
            -distB ST_Emu
            -distL ST_Emu
            -distR ST_Emu
            -hidden CT_Boolean
            -locked CT_Boolean
            -layoutInCell CT_Boolean
            -allowOverlap CT_Boolean
            -relativeHeight CT_UnsignedInt
        } -anchorData]
        Tag_wp:anchor [array get anchorAtts] {
            Tag_wp:simplePos x 0 y 0
            Tag_wp:positionH [my Option -positionH relativeFrom ST_RelFromH "column"] {
                lassign [my OneOff {-alignH ST_AlignH} {-posOffsetH ST_Emu}] \
                    alignH posOffsetH
                if {$alignH eq "" && $posOffsetH eq ""} {
                    set alignH "center"
                }
                foreach value [list $alignH $posOffsetH] tag {Tag_wp:align Tag_wp:posOffset} {
                    if {$value ne ""} {
                        $tag {Text $value}
                        break
                    }
                }
            }
            Tag_wp:positionV [my Option -positionV relativeFrom ST_RelFromV "paragraph"] {
                lassign [my OneOff {-alignV ST_AlignV} {-posOffsetV ST_Emu}] \
                    alignV posOffsetV
                if {$alignV eq "" && $posOffsetV eq ""} {
                    set alignV "center"
                }
                foreach value [list $alignV $posOffsetV] tag {Tag_wp:align Tag_wp:posOffset} {
                    if {$value ne ""} {
                        $tag {Text $value}
                        break
                    }
                }
            }
            set thisOptionValue [my PeekOption -dimension]
            set attlist [my CheckedAttlist $thisOptionValue {
                {width -cx} ST_Emu
                {height -cy} ST_Emu
            } -dimension]
            if {[llength $attlist] != 4} {
                error "the -dimension option expects both keys \"width\" and\
                      \"height\" to be given"
            }
            Tag_wp:extent {*}$attlist
            Tag_wp:effectExtent l 0 t 0 r 0 b 0
            set wrapMode [my EatOption -wrapMode]
            switch $wrapMode {
                "" -
                "none" {
                    if {[my EatOption -wrapData] ne ""} {
                        error "the option \"-wrapMode none does not expect \"-wrapData\""
                    }
                    Tag_wp:wrapNone
                }
                "square" {
                    set attlist [my CheckedAttlist [my EatOption -wrapData] {
                        -wrapText ST_WrapText
                        -distT ST_Emu
                        -distB ST_Emu
                        -distL ST_Emu
                        -distR ST_Emu
                    } -wrapData]
                    Tag_wp:wrapSquare {*}$attlist
                }
                "topBottom" {
                    Tag_wp:wrapTopAndBottom
                }
                default {
                    error "unknown -wrapMode option value \"$wrapMode\", expect\
                         one of [AllowedValues {none square topBottom}]"
                }
            }
            Tag_wp:docPr id [my NextId docPr] name $name
            Tag_wp:cNvGraphicFramePr {
                Tag_a:graphicFrameLocks noChangeAspect 1
            }
            set anchor [dom fromScriptContext]
        }
        return $anchor
    }

    method CallType {type value errtext} {
        # A few value checks need access to the docx object internal
        # data and therefor are implemented as object methods. The
        # bulk of the type checks are "static" procs in the
        # ooxml::docx namespace.
        set error 0
        if {[catch {set ooxmlvalue [$type $value]} errMsg]} {
            if {![llength [info procs ::ooxml::docx::lib::$type]]} {
                if {[catch {set ooxmlvalue [my $type $value]} errMsg]} {
                    set error 1
                }
            } else {
                set error 1
            }
            if {$error} {
                if {[info exists ::ooxml::docx::docgen] && $::ooxml::docx::docgen} {
                    error "===$type=== $errMsg"
                } else {
                    error "$errtext: $errMsg"
                }
            }
        }
        return $ooxmlvalue
    }

    method CheckedAttlist {optionValue attdefs option} {
        variable valPrefix

        set attlist ""
        if {[catch {array set atts $optionValue}]} {
            set keys ""
            foreach {attdata type} $attdefs {
                lappend keys [string trimleft [lindex $attdata 0] -]
            }
            error "the value \"$optionValue\" given to the \"$option\" \
                           option is invalid, expected is a key value pairs\
                           list with keys out of [AllowedValues $keys]"
        }
        foreach {attdata typedata} $attdefs {
            if {[llength $attdata] == 2} {
                lassign $attdata key attname
            } else {
                if {[string index $attdata 0] eq "-"} {
                    set key [string range $attdata 1 end]
                } else {
                    set key $attdata
                }
                set attname $attdata
            }
            lassign $typedata type flags default
            if {![info exists atts($key)]} {
                if {$flags eq "!"} {
                    if {[llength $typedata] == 2} {
                        error "the argument \"$optionValue\" given to the\
                                \"$option\" option is invalid: the mandatory\
                                key \"$key\" in the argument\
                                is invalid"
                    } else {
                        set atts($key) $default
                    }
                } else {
                    continue
                }
            }
            set ooxmlvalue [my CallType $type $atts($key) \
                                "the argument \"$optionValue\" given to the\
                                \"$option\" option is invalid: the value\
                                given to the key \"$key\" invalid"]
            if {[string index $attname 0] eq "-"} {
                lappend attlist [string range $attname 1 end] $ooxmlvalue
            } else {
                lappend attlist ${valPrefix}:$attname $ooxmlvalue
            }
            unset atts($key)
        }
        # Check if there are unknown keys left in the key
        # values list
        set remainigKeys [array names atts]
        if {[llength $remainigKeys] != 0} {
            set keys ""
            foreach {attdata type} $attdefs {
                lappend keys [lindex $attdata 0]
            }
            if {[llength $remainigKeys] == 1} {
                error "unknown key \"[lindex $remainigKeys 0]\" in\
                                the value \"$optionValue\" of the option\
                                \"$option\", the expected keys are\
                                [AllowedValues $keys]"
            } else {
                error "unknown keys [AllowedValues $remainigKeys and] in\
                               the value \"$optionValue\" of the option\
                               \"$option\", the expected keys are\
                               [AllowedValues $keys]"
            }
        }
        return $attlist
    }

    method CheckRemainingOpts {} {
        my Prepare
        set nrRemainigOpts [llength [array names opts]]
        if {$nrRemainigOpts == 0} return
        if {$nrRemainigOpts == 1} {
            set text "unknown option: [lindex [array names opts] 0]"
        } else {
            set text "unknown options: [AllowedValues [array names opts] and]"
        }
        set knownOps [lsort [array names optsknown]]
        set knownOpsLen [llength $knownOps]
        if {$knownOpsLen} {
            if {$knownOpsLen > 1} {
                append text ", known options are [AllowedValues $knownOps and]"
            } else {
                append text ", known option is $knownOps"
            }
        }
        uplevel [list error $text]
    }

    method CheckrId {type rId} {
        my variable docs
        variable ::ooxml::docx::xmlns

        set relsRoot [$docs(word/_rels/document.xml.rels) documentElement]
        set type "$xmlns(r)/$type"
        if {![$relsRoot selectNodes -namespaces "rel $xmlns(rel)" {
            count(rel:Relationship[@Id=$rId and @Type=$type])
        }]} {
            error "invalid rId $rId"
        }
    }

    method ControlChar {c {n 1}} {
        set p [my LastParagraph]
        for {set i 0} {$i<$n} {incr i} {
            $p appendFromScript {
                Tag_w:r {
                    Tag_w:$c
                }
            }
        }
    }

    method Create switchActionList {
        variable valPrefix
        my Prepare
        foreach {opt optdata} $switchActionList {
            # If the opt is just / then this is a placeholder for an
            # element which currently cannot be created by an option.
            # The optdata gives the element name.
            if {$opt eq "/"} {
                continue
            }
            # If the option description starts with a + then this
            # describes not an option which would create a child with
            # one or several attributes. It describes a child (with
            # the opt value stripped by the leading + as name) with
            # child nodes and the optdata are the ordinary description
            # of the options to create those children with values in
            # attributes.
            if {[string index $opt 0] eq "+"} {
                Tag_[string range $opt 1 end] {
                    my Create $optdata
                }
                continue
            }
            # If the opt is just | then optdata is a list of mutually
            # exclusive usual option descriptions.
            if {$opt eq "|"} {
                # We check the last child of the context to see if the
                # option group has created something (and so at least
                # one option of this group was given).
                set lastChild [[dom fromScriptContext] lastChild]
                set seen 0
                set groupOptions [list]
                foreach desc $optdata {
                    array unset startOptsknown
                    array set startOptsknown [array get optsknown]
                    my Create $desc
                    set thisopts [list]
                    foreach thisopt [array names optsknown] {
                        if {![info exists startOptsknown($thisopt)]} {
                            lappend thisopts $thisopt
                        }
                    }
                    lappend groupOptions [lsort $thisopts]
                    set thisLast [[dom fromScriptContext] lastChild]
                    if {$lastChild ne $thisLast} {
                        incr seen
                        set lastChild $thisLast
                    }
                }
                if {$seen > 1} {
                    error "The options [join $groupOptions " and "] are\
                           mutually exclusive."
                }
                continue
            }
            # The "normal" cases
            set optsknown($opt) ""
            if {![info exists opts($opt)]} {
                continue
            }
            set value $opts($opt)
            switch [llength $optdata] {
                2 {
                    lassign $optdata tags attdefs
                    if {[llength $attdefs] == 1} {
                        # An element with just w:val as attribute and
                        # the attdefs gives the value type.
                        set ooxmlvalue [my CallType $attdefs $value \
                                            "the value \"$value\" given to\
                                             the \"$opt\" option is invalid"]
                        foreach tag $tags {
                            Tag_$tag ${valPrefix}:val $ooxmlvalue
                        }
                        unset opts($opt)
                        continue
                    }
                    # This handles the (rare) case of tags with just one
                    # attribute to set and that attribute is not w:val
                    if {[llength $attdefs] == 2} {
                        lassign $attdefs attname atttype
                        set ooxmlvalue [my CallType $atttype $value \
                                            "the value \"$value\" given to the \"$opt\"\
                                              option is invalid"]
                        if {[string index $attname 0] eq "-"} {
                            set attname [string range $attname 1 end]
                        } else {
                            set attname ${valPrefix}:$attname
                        }
                        foreach tag $tags {
                            Tag_$tag $attname $ooxmlvalue
                        }
                        unset opts($opt)
                        continue
                    }
                    # For now the code assumes always several atts
                    # and therefore the value to the option is always
                    # handled as a key value pairs list.
                    set attlist [my CheckedAttlist $value $attdefs $opt]
                    foreach tag $tags {
                        Tag_$tag {*}$attlist
                    }
                    unset opts($opt)
                }
                3 {
                    my [lindex $optdata 2] $value
                    unset opts($opt)

                }
                default {
                    error "invalid properties value"
                }
            }
        }
    }

    method CreateComment {{commentID ""}} {
        my variable docs
        my variable ignorable
        variable ::ooxml::docx::xmlns
        variable ::ooxml::docx::prefixnslist

        my Prepare
        if {![info exists docs(word/comments.xml)]} {
            my Add2Relationships comments comments.xml
            set document [dom createDocumentNS $xmlns(w) w:comments]
            set docs(word/comments.xml) $document
            $docs(word/comments.xml) selectNodesNamespaces $prefixnslist
            $document documentElement commentsroot
            foreach ns {o m r v w10 wp wps wpg mc wp14 w14} {
                $commentsroot setAttributeNS {} xmlns:$ns $xmlns($ns)
            }
            my Ignorable word/comments.xml
        }
        set comments [$docs(word/comments.xml) documentElement]
        if {$commentID eq ""} {
            set commentID [my NextId comments]
        }
        $comments appendFromScript {
            set author [my EatOption -author NoCheck "Unknown"]
            Tag_w:comment w:id $commentID \
                w:date [my EatOption -date ST_DateTime] \
                w:author $author \
                w:initials [my EatOption -initials NoCheck]
        }
        return [list [$comments lastChild] $commentID]
    }

    method EatOption {option {type ""} {default ""} {deleteOption 1}} {
        my Prepare
        set optsknown($option) ""
        if {[info exists opts($option)]} {
            set value $opts($option)
            if {$type ne ""} {
                set value [my CallType $type $value "option $option"]
            }
            if {$deleteOption} {
                unset opts($option)
            }
            return $value
        }
        return $default
    }

    method FootnoteEndnote {type refstyle script} {
        my variable docs
        my variable body
        my variable ignorable
        variable ::ooxml::docx::xmlns
        variable ::ooxml::docx::prefixnslist

        set types "${type}s"
        set notePart "word/$types.xml"
        if {![info exists docs($notePart)]} {
            my Add2Relationships $types $types.xml
            set document [dom createDocumentNS $xmlns(w) w:$types]
            set docs($notePart) $document
            $docs($notePart) selectNodesNamespaces $prefixnslist
            $document documentElement notesroot
            foreach ns {o m r v w10 wp wps wpg mc wp14 w14} {
                $notesroot setAttributeNS {} xmlns:$ns $xmlns($ns)
            }
            my Ignorable $notePart
        } else {
            set notesroot [$docs($notePart) documentElement]
        }

        foreach {thisid sepType} {
            0 separator
            1 continuationSeparator
        } {
            if {![$notesroot selectNodes "count(w:$type\[@w:id=$thisid])"]} {
                $notesroot appendFromScript {
                    Tag_w:$type w:type separator w:id $thisid {
                        Tag_w:p { Tag_w:r { Tag_w:$sepType } }
                    }
                }
            }
        }
        if {![info exists id($types)] || $id($types) < 1} {
            set id($types) 1
        }

        set thisid [my NextId $types]
        $notesroot appendFromScript {
            Tag_w:$type w:id $thisid
        }
        set savedbody $body
        set body [$notesroot lastChild]
        if {[catch {uplevel 2 [list eval $script]} errMsg errVals]} {
            set body $savedbody
            my ProcessErrorinfo $type
            return -code error $errMsg
        }
        set body $savedbody
        set firstp [$notesroot selectNodes {w:p[1]}]
        if {$firstp eq ""} {
            $notesroot appendFromScript Tag_w:p
            set firstp [$notesroot lastChild]
        }
        $firstp insertBeforeFromScript {
            Tag_w:r {
                if {$refstyle ne ""} {
                    Tag_w:rPr {
                        Tag_w:rStyle w:val $refstyle
                    }
                }
                Tag_w:${type}Ref
            }
        } [$firstp selectNodes {w:r[1]}]
        return $thisid
    }

    method GetDocDefault {styles} {
        # styles has the content model sequence:
        # docDefaults? latentStyles? styles*
        set docDefaults [$styles firstChild]
        if {$docDefaults == "" || [$docDefaults localName] ne "docDefaults"} {
            set nextNode $docDefaults
            $styles insertBeforeFromScript Tag_w:docDefaults $nextNode
            set docDefaults [$styles firstChild]
        }
        return $docDefaults
    }

    method HeaderFooter {what script} {
        my variable docs
        my variable body
        my variable context
        my variable ignorable
        variable ::ooxml::docx::xmlns
        variable ::ooxml::docx::prefixnslist

        set have [lsort -dictionary [array names docs word/$what*]]
        if {![llength $have]} {
            set nr 1
        } else {
            set last [string range [lindex $have end] 0 end-4]
            set nr [string range $last [expr {5 + [string length $what]}] end]
            incr nr
        }
        set rId [my Add2Relationships $what $what$nr.xml]
        if {$what eq "header"} {
            set elnName "hdr"
        } else {
            set elnName "ftr"
        }
        set document [dom createDocumentNS $xmlns(w) w:$elnName]
        set partName "word/$what$nr.xml"
        set docs($partName) $document
        $docs($partName) selectNodesNamespaces $prefixnslist
        $document documentElement root
        foreach ns {o m r v w10 wp wps wpg mc wp14 w14} {
            $root setAttributeNS {} xmlns:$ns $xmlns($ns)
        }
        my Ignorable $partName
        set savedbody $body
        set savedcontext $context
        set body $root
        set context $what$nr
        if {[catch {uplevel 2 [list eval $script]} errMsg errVals]} {
            set body $savedbody
            set context $savedcontext
            my ProcessErrorinfo $what
            return -code error $errMsg
        }
        set body $savedbody
        set context $savedcontext
        return $rId
    }

    method Ignorable {{file ""}} {
        my variable docs
        my variable ignorable
        variable ::ooxml::docx::xmlns

        if {$file eq ""} {
            set files [array names docs word/*.xml]
        } else {
            set files [list $file]
        }
        foreach doc $files  {
            set thisdoc $docs($doc)
            set storedprefixns [$thisdoc selectNodesNamespaces]
            $thisdoc selectNodesNamespaces ""
            set thisroot [$thisdoc documentElement]
            if {[catch {$thisroot selectNodes mc:*}]} {
                # XML namespace prefix mc not defined
                $thisdoc selectNodesNamespaces $storedprefixns
                continue
            }
            $thisdoc selectNodesNamespaces $storedprefixns
            array unset prefixlookup
            array set prefixlookup [join [$thisroot selectNodes namespace::*]]
            set knownprefixes ""
            foreach prefix $ignorable {
                if {[info exists prefixlookup(xmlns:$prefix)]} {
                    lappend knownprefixes $prefix
                }
            }
            if {![llength $knownprefixes]} {
                if {[$thisroot hasAttributeNS $xmlns(mc) Ignorable]} {
                    $thisroot removeAttributeNS $xmlns(mc) Ignorable
                }
                continue
            }
            $thisroot setAttributeNS \
                $xmlns(mc) mc:Ignorable $knownprefixes
        }
    }

    method Image_anchor {rId file} {
        my Prepare
        set anchor [my Anchor [file tail $file]]
        $anchor appendFromScript {
            my Image_graphic $rId $file
        }
    }

    method Image_graphic {rId file} {
        my Prepare
        Tag_a:graphic {
            Tag_a:graphicData uri "http://schemas.openxmlformats.org/drawingml/2006/picture" {
                Tag_pic:pic {
                    Tag_pic:nvPicPr {
                        Tag_pic:cNvPr id [my NextId pic] name [file rootname [file tail $file]]
                        Tag_pic:cNvPicPr {
                            Tag_a:picLocks  noChangeAspect 1 noChangeArrowheads 1
                        }
                    }
                    Tag_pic:blipFill {
                        Tag_a:blip r:embed $rId
                        Tag_a:stretch {
                            Tag_a:fillRect
                        }
                    }
                    Tag_pic:spPr {*}[my Option -bwMode bwMode ST_BlackWhiteMode "auto"] {
                        my SpPr_Content
                    }
                }
            }
        }
    }

    method Image_inline {rId file} {
        my Prepare
        Tag_wp:inline {
            set thisOptionValue [my PeekOption -dimension]
            set attlist [my CheckedAttlist $thisOptionValue {
                {width -cx} ST_Emu
                {height -cy} ST_Emu
            } -dimension]
            if {[llength $attlist] != 4} {
                error "the -dimension option expects both keys \"width\" and\
                      \"height\" to be given"
            }
            Tag_wp:extent {*}$attlist
            Tag_wp:docPr id [my NextId docPr] name [file tail $file]
            my Image_graphic $rId $file
        }
    }

    method InitFromDocx {docxfile} {
        my variable docs
        my variable body
        my variable binparts
        my variable impPagesetup
        my variable id
        variable ::ooxml::docx::prefixnslist

        if {[catch {::ooxml::ZipOpen $docxfile} errMsg]} {
            return -code error "Cannot open '$docxfile': $errMsg"
        }
        set id(image) 0
        foreach part [::ooxml::ZipMembers] {
            if {[catch {set docs($part) [::ooxml::ZipReadParse $part]}]} {
                # Non XML file; add to binparts
                set binparts($part) [::ooxml::ZipReadBinaryFile $part]
            }
            if {[regexp {^word/media/image(\d+)} $part -> n]} {
                if {$id(image) < $n} {
                    set id(image) $n
                }
            }
        }
        ::ooxml::ZipClose

        # Check that (by the spec) required parts of the docx are
        # actually exist (the implementation expects them as present).
        foreach musthave {
            [Content_Types].xml
            word/document.xml
            word/_rels/document.xml.rels
        } {
            if {![info exists docs($musthave)]} {
                return -code error "Invalid ooxml docx file '$docxfile':\
                                    missing required part '$musthave'"
            }
        }
        # Ensure that a word/styles.xml is there
        if {![info exists docs(word/styles.xml)]} {
            variable ::ooxml::docx::staticDocx
            set docs(word/styles.xml) [dom parse $staticDocx(word/styles.xml)]
        }
        # Ensure that we can use our prefixes for XPath queries
        foreach xmlpart [array names docs] {
            $docs($xmlpart) selectNodesNamespaces $prefixnslist
        }
        # Move a present sectPr elemet away (to restore it at write
        # time, if the user has not specified another page setup).
        set body [$docs(word/document.xml) selectNodes {w:document/w:body[last()]}]
        if {$body eq ""} {
            return -code error "Invalid ooxml docx file '$docxfile':\
                no body element in word/document.xml"
        }
        set lastElm [$body selectNodes {w:*[last()]}]
        if {[$lastElm localName] eq "sectPr"} {
            set impPagesetup $lastElm
            $body removeChild $lastElm
        }
        # Seed index counters
        foreach key {docPr pic bookmarks textboxes comments footnotes endnotes} {
            set id($key) 0
        }
        foreach part [array names docs "word/*.xml"] {
            foreach {key tag att} {
                docPr wp:docPr id
                pic pic:cNvPr id
                bookmarks w:bookmarkStart w:id
            } {
                foreach attribute [$docs($part) selectNodes //$tag/@$att] {
                    lassign $attribute name v
                    if {[string is integer -strict $v] && $v > $id($key)} {
                        set id($key) $v
                    }
                }
            }
            foreach attribute [$docs($part) selectNodes //wp:docPr/@name] {
                lassign $attribute name value
                if {[regexp {^Textbox\s+(\d+)$} $value -> n] && $n > $id(textboxes)} {
                    set id(textboxes) $n
                }
            }
        }
        foreach {key part tag} {
            comments word/comments.xml w:comment
            footnotes word/footnotes.xml w:footnote
            endnotes word/endnotes.xml w:endnote
        } {
            if {![info exists docs($part)]} continue
            foreach attribute [$docs($part) selectNodes //$tag/@w:id] {
                lassign $attribute name v
                if {[string is integer -strict $v] && $v > $id($key)} {
                    set id($key) $v
                }
            }
        }
    }

    method LastParagraph {{create 0}} {
        my variable body

        set p [$body lastChild]
        while {$p ne ""} {
            if {[$p nodeType] ne "ELEMENT_NODE"} {
                set p [$p previousSibling]
                continue
            }
            if {[$p nodeName] ne "w:p"} {
                # Perhaps it would be more wise to raise error in this
                # situation or to create a new one?
                set p [$p previousSibling]
                continue
            }
            break
        }
        if {$p eq ""} {
            if {!$create} {
                error "no paragraph to append to in the document"
            }
            $body appendFromScript {
                Tag_w:p
            }
            set p [$body lastChild]
        }
        return $p
    }

    method NextId {domain} {
        my variable id

        # Return the next unique id
        return [incr id($domain)]
    }

    method OneOff {opta optb} {
        my Prepare
        lassign $opta optiona typea
        lassign $optb optionb typeb
        set a [my EatOption $optiona $typea]
        set b [my EatOption $optionb $typeb]
        if {$a ne "" && $b ne ""} {
            error "the options $optiona and $optionb are mutually exclusive"
        }
        return [list $a $b]
    }

    method Option {option attname type {default ""}} {
        my Prepare
        set optsknown($option) ""
        if {[info exists opts($option)]} {
            set ooxmlvalue [my CallType $type $opts($option) \
                                "the value given to the \"$option\" option\
                                 is invalid"]
            unset opts($option)
        } else {
            set ooxmlvalue $default
        }
        return [list $attname $ooxmlvalue]
    }

    method PeekOption {option {type ""} {default ""}} {
        my Prepare
        return [my EatOption $option $type $default 0]
    }


    method Prepare {{properties 0}} {
        set script "upvar opts opts; upvar optsknown optsknown;"
        if {$properties} {
            append script "variable ::ooxml::docx::properties"
        }
        uplevel [list eval $script]
    }

    method ProcessErrorinfo {what} {
        upvar errMsg errMsg
        upvar errVals errVals

        set errorinfo [dict get $errVals -errorinfo]
        set ind [string first "(\"eval\" body line " $errorinfo]
        if {$ind > -1} {
            set errMsg [string range $errorinfo 0 $ind-1]
            if {[regexp {(\d+)} [string range $errorinfo $ind end] --> linenr]} {
                append errMsg "($what script body line $linenr)"
            }
        }
    }

    method PPr {} {
        my Prepare 1
        Tag_w:pPr {
            my Create $properties(paragraph1)
            Tag_w:pBdr {
                my Create $properties(paragraphBorders)
            }
            my Create $properties(paragraph2)
            Tag_w:tabs {
                my Tabs [my EatOption -tabs]
            }
            my Create $properties(paragraph3)
        }
    }

    method PStyle {value} {
        return [my StyleCheck paragraph $value]
    }

    method RFonts {value} {
        Tag_w:rFonts \
            w:ascii $value \
            w:hAnsi $value \
            w:eastAsia $value \
            w:cs $value
    }

    method RPr {} {
        my Prepare 1
        Tag_w:rPr {
            my Create $properties(run)
        }
    }

    method RStyle {value} {
        return [my StyleCheck character $value]
    }

    method SectionCommon {} {
        my Prepare 1
        Tag_w:sectPr {
            foreach what {Header Footer} {
                foreach type {even default first} {
                    set value [my EatOption -$type$what]
                    if {$value eq ""} {
                        continue
                    }
                    if {$type eq "even"} {
                        my settings -evenAndOddHeaders on
                    }
                    my CheckrId [string tolower $what] $value
                    Tag_w:[string tolower $what]Reference w:type $type r:id $value
                }
            }
            my Create $properties(sectionsetup1)
            Tag_w:pgBorders {
                my Create $properties(sectionBorders)
            }
            my Create $properties(sectionsetup2)
        }
    }

    method SpPr_Content {} {
        my Prepare 1
        Tag_a:xfrm {
            Tag_a:off x 0 y 0
            my Create $properties(xfrm)
        }
        Tag_a:prstGeom prst "rect" {
            Tag_a:avLst
        }
    }

    method StyleCheck {type value} {
        my variable docs

        set styles [$docs(word/styles.xml) documentElement]
        set style [$styles selectNodes {
            w:style[@w:type=$type and @w:styleId=$value]
        }]
        if {![llength $style]} {
            set style [$styles selectNodes {
                w:style[@w:type=$type][w:name[@w:val=$value]]
            }]
        }
        if {![llength $style]} {
            error "unknown $type style \"$value\""
        }
        return [[lindex $style 0] selectNodes string(@w:styleId)]
    }

    method Tabs {tabsdata} {
        foreach tabstop $tabsdata {
            if {[llength $tabstop] == 1} {
                OptVal [list -tabs [list pos $tabstop type "start"]]
            } else {
                OptVal [list -tabs [list $tabstop]]
            }
            my Create {
                -tabs {w:tab {
                    leader ST_TabTlc
                    pos ST_SignedTwipsMeasure
                    {type val} ST_TabJc
                }}
            }
            unset opts
        }
    }

    method TblPr {} {
        my Prepare 1
        Tag_w:tblPr {
            my Create $properties(table1)
            Tag_w:tblBorders {
                my Create $properties(tableBorders)
            }
            my Create $properties(table2)
            Tag_w:tblCellMar {
                my Create $properties(cellMargins)
            }
            my Create $properties(table3)
        }

    }

    method TblStylePr {conditionData} {
        set values {
            wholeTable
            firstRow
            lastRow
            firstCol
            lastCol
            band1Vert
            band2Vert
            band1Horz
            band2Horz
            neCell
            nwCell
            seCell
            swCell
        }
        foreach {types styledata} $conditionData {
            foreach type $types {
                if {$type ni $values} {
                    error "unknown table overwrite style \"$type\", expected\
                           one of [AllowedValues $values]"
                }
                OptVal $styledata
                array unset optsknown
                Tag_w:tblStylePr w:type $type {
                    my PPr
                    my RPr
                    my TblPr
                    my TrPr
                    my TcPr
                }
                my CheckRemainingOpts
            }
        }
    }

    method TcPr {} {
        my Prepare 1
        Tag_w:tcPr {
            my Create $properties(cell1)
            Tag_w:tcBorders {
                my Create $properties(cellBorders)
            }
            my Create $properties(cell2)
            Tag_w:tcMar {
                my Create $properties(cellMargins)
            }
            my Create $properties(cell3)
        }
    }

    method TrPr {} {
        my Prepare 1
        Tag_w:trPr {
            my Create $properties(row)
        }
    }

    method TStyle {value} {
        return [my StyleCheck table $value]
    }

    method Wt {text} {
        # Just not exactly that easy.
        # Handle at least \n \r \t and \f special
        set pos 0
        set end [string length $text]
        foreach part [split $text "\n\r\t\f"] {
            set len [string length $part]
            if {$len} {
                set atts ""
                if {[string index $part 0] eq " " || [string index $part end] eq " "} {
                    lappend atts xml:space preserve
                }
                Tag_w:t $atts {
                    Text [dom clearString -replace $part]
                }
            }
            incr pos $len
            switch [string index $text $pos] {
                "\n" Tag_w:br
                "\r" Tag_w:cr
                "\t" Tag_w:tab
                "\f" {Tag_w:br w:type "page"}
            }
            incr pos
        }
    }

    method addXML {node xmlstr} {
        set doc [dom parse [subst -nocommands -nobackslashes {<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">$xmlstr</w:document>}]]
        foreach child [[$doc documentElement] childNodes] {
            $node appendChild $child
        }
        $doc delete
    }

    method append {text args} {
        if {[catch {
            OptVal $args "text"
            set p [my LastParagraph]
            $p appendFromScript {
                Tag_w:r {
                    my RPr
                    my Wt $text
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method br {{n 1}} {
        my ControlChar br $n
    }

    method comment {args} {
        my variable body
        my variable context

        if {[catch {
            OptVal [lrange $args 0 end-1] "" "<comment script>"
            lassign [my CreateComment] comment id
            my CheckRemainingOpts
            set script [lindex $args end]
            set savedbody $body
            set savedcontext $context
            set body $comment
            # The nested catch is needed to ensure body is set back
            if {[catch {
                uplevel [list eval $script]
            } errMsg errVals]} {
                set body $savedbody
                set context $savedcontext
                my ProcessErrorinfo "comment"
                error $errMsg
            }
        } errMsg]} {
            return -code error $errMsg
        }
        set body $savedbody
        set context $savedcontext
        # Add the comment mark to the document
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:r {
                Tag_w:commentReference w:id $id
            }
        }
    }

    method commentrangeend {id args} {
        my variable body
        my variable context
        my variable commentranges

        if {[catch {
            if {![info exists commentranges($id)]} {
                error "no open comment range with the id '$id'"
            }
            OptVal [lrange $args 0 end-1]
            lassign [my CreateComment $id] comment id
            my CheckRemainingOpts
            set script [lindex $args end]
            set savedbody $body
            set savedcontext $context
            set body $comment
            # The nested catch is needed to ensure body is set back
            if {[catch {
                uplevel [list eval $script]
            } errMsg errVals]} {
                set body $savedbody
                set context $savedcontext
                my ProcessErrorinfo "commentend"
                error $errMsg
            }
        } errMsg]} {
            return -code error $errMsg
        }
        set body $savedbody
        set context $savedcontext
        # Add the comment mark to the document
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:commentRangeEnd w:id $id
            Tag_w:r {
                Tag_w:commentReference w:id $id
            }
        }
        unset commentranges($id)
    }


    method commentrangestart {{returnvar ""}} {
        my variable body
        my variable context
        my variable commentranges

        if {$returnvar ne ""} {
            upvar $returnvar id
        }
        set id [my NextId comments]
        # Add the comment range start mark to the document
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:commentRangeStart w:id $id
        }
        set commentranges($id) ""
        return $id
    }

    method configure {args} {
        my variable docs
        my variable ignorable

        if {[catch {
            OptVal $args
            # We need to look if -ignorable was given to be able to
            # handle the empty string as value."
            if {[info exists opts(-ignorable)]} {
                set ignorable [my EatOption -ignorable]
                my Ignorable
            }
            set coreroot [$docs(docProps/core.xml) documentElement]
            $coreroot appendFromScript {
                foreach elem {
                    cp:category
                    cp:contentStatus
                    dcterms:created
                    dc:creator
                    dc:description
                    dc:identifier
                    cp:keywords
                    dc:language
                    cp:lastModifiedBy
                    cp:lastPrinted
                    dcterms:modified
                    cp:revision
                    dc:subject
                    dc:title
                    cp:version
                } {
                    # Hm. According to opc-coreProperties.xsd
                    # cp:keywords should have value elements as
                    # children, but I don't see this in what
                    # libreoffice generates.
                    lassign [split $elem :] prefix option
                    set value [my EatOption -$option]
                    if {$value ne ""} {
                        set attlist ""
                        if {$prefix eq "dcterms"} {
                            set value [W3CDTF $value]
                            lappend attlist xsi:type dcterms:W3CDTF
                        }
                        foreach currentnode [$coreroot selectNodes {*[local-name()=$option]}] {
                            $currentnode delete
                        }
                        ::tdom::fsnewNode $elem $attlist {Text $value}
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method endnote {args} {
        if {[catch {
            set script [lindex $args end]
            OptVal [lrange $args 0 end-1] "endnote" "script"
            set id [my FootnoteEndnote endnote \
                        [my EatOption -refstyle RStyle] $script]
            set p [my LastParagraph 1]
            $p appendFromScript {
                Tag_w:r {
                    my RPr
                    Tag_w:endnoteReference w:id $id
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method field {field {switches ""} args} {
        my variable docs
        if {[catch {
            set nfield [string toupper $field]
            set values {
                AUTHOR
                CREATEDATE
                DATE
                FILESIZE
                NUMPAGES
                PAGE
                REF
                SAVEDATE
                SECTION
                SEQ
                TIME
                TITLE
                TOC
                USERNAME
            }
            if {$nfield ni $values} {
                return -code error "Unknown field type '$field', expected one\
                                         out of [AllowedValues $values]"
            }
            set dirtyAttrib ""
            if {$nfield in {REF SEQ TOC}} {
                set dirtyAttrib {w:dirty "true"}
            }
            if {$switches ne ""} {
                set nfield "$nfield $switches"
            }
            OptVal $args "field switches"
            set p [my LastParagraph 1]
            $p appendFromScript {
                Tag_w:r {
                    my RPr
                    Tag_w:fldChar w:fldCharType "begin" {*}$dirtyAttrib
                }
                Tag_w:r {
                    OptVal $args
                    my RPr
                    Tag_w:instrText xml:space preserve {
                        Text "$nfield"
                    }
                }
                Tag_w:r {
                    OptVal $args
                    my RPr
                    Tag_w:fldChar w:fldCharType "separate"
                }
                Tag_w:r {
                    OptVal $args
                    my RPr
                    Tag_w:fldChar w:fldCharType "end"
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method footer {script {returnvar ""}} {
        if {$returnvar ne ""} {
            upvar $returnvar result
        }
        if {[catch {
            set result [my HeaderFooter footer $script]
        } errMsg]} {
            return -code error $errMsg
        }
        return $result
    }

    method footnote {args} {
        if {[catch {
            set script [lindex $args end]
            OptVal [lrange $args 0 end-1] "footnote" "script"
            set id [my FootnoteEndnote footnote \
                        [my EatOption -refstyle RStyle] $script]
            set p [my LastParagraph 1]
            $p appendFromScript {
                Tag_w:r {
                    my RPr
                    Tag_w:footnoteReference w:id $id
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method header {script {returnvar ""}} {
        if {$returnvar ne ""} {
            upvar $returnvar result
        }
        if {[catch {
            set result [my HeaderFooter header $script]
        } errMsg]} {
            return -code error $errMsg
        }
        return $result
    }

    method image {file type args} {
        my variable binparts

        if {[catch {
            if {![file isfile $file] || ![file readable $file]} {
                error "cannot read file \"$file\""
            }
            if {$type ni {inline anchor}} {
                error "invalid image type \"$type\" (expected \"inline\" or \"anchor\")"
            }
            OptVal $args "file type"
            set imagename image[my NextId image]
            append imagename [string tolower [file extension $file]]
            set fd [open $file rb]
            set binparts(word/media/$imagename) [read $fd]
            close $fd
            set rId [my Add2Relationships image media/$imagename]
            set p [my LastParagraph 1]
            $p appendFromScript {
                Tag_w:r {
                    Tag_w:drawing {
                        my Image_$type $rId $file
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method import {what docxfile} {
        my variable docs
        variable ::ooxml::docx::xmlns

        ::ooxml::ZipOpen $docxfile
        set rId ""
        try {
            switch -glob $what {
                "styles" {
                    set what "word/styles.xml"
                }
                "numbering" {
                    set what "word/numbering.xml"
                }
                "header*" -
                "footer*" {
                    set what "word/$what.xml"
                }
            }
            set thisdoc [::ooxml::ZipReadParse $what]
            set target [string range $what 5 end]
            if {[info exists docs($what)]} {
                $docs($what) delete
                if {[string range $what 0 4] eq "word/"} {
                    set relsRoot [$docs(word/_rels/document.xml.rels) documentElement]
                    set rId [$relsRoot selectNodes -namespaces [list r $xmlns(rel)] {
                        string(r:Relationship[@Target=$target])
                    }]
                }
            } elseif {[string range $what 0 4] eq "word/"} {
                set thisrels [::ooxml::ZipReadParse "word/_rels/document.xml.rels"]
                set relsroot [$thisrels documentElement]
                set typeurl [$relsroot selectNodes -namespaces [list r $xmlns(rel)] {
                    string(r:Relationship[@Target=$target]/@Type)
                }]
                # The typeurl starts with
                # http://schemas.openxmlformats.org/officeDocument/2006/relationships/
                set rId [my Add2Relationships [string range $typeurl 68 end] $target]
                $thisrels delete
            }
            set docs($what) $thisdoc
        } finally {
            ::ooxml::ZipClose
        }
        return $rId
    }

    method jumpto {text name args} {
        my variable bookmarks

        if {[catch {
            OptVal $args "text mark"
            set p [my LastParagraph 1]
            $p appendFromScript {
                Tag_w:hyperlink w:anchor $name {
                    Tag_w:r {
                        my RPr
                        my Wt $text
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
        if {![info exists bookmarks($name)]} {
            set bookmarks($name) 0
        }
    }

    method mark {name} {
        if {[catch {
            my markstart $name
            my markend $name
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method markend {name} {
        my variable bookmarks

        if {![info exists bookmarks($name)] || $bookmarks($name) == 0} {
            return -code error "A mark span with the name \"$name\" is\
                                not started."
        }
        if {$bookmarks($name) < 0} {
            return -code error "The mark with the name \"$name\" is\
                                already ended."
        }
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:bookmarkEnd w:id $bookmarks($name)
        }
        set bookmarks($name) "-$bookmarks($name)"
    }

    method markstart {name} {
        my variable bookmarks

        if {[info exists bookmarks($name)] && $bookmarks($name) > 0} {
            return -code error "mark \"$name\" is not unique"
        }
        set p [my LastParagraph 1]
        set thisid [my NextId bookmarks]
        $p appendFromScript {
            Tag_w:bookmarkStart w:id $thisid w:name $name
        }
        set bookmarks($name) $thisid
    }

    method numbering {cmd args} {
        my variable docs
        variable ::ooxml::docx::properties
        variable ::ooxml::docx::prefixnslist

        if {![info exists docs(word/numbering.xml)]} {
            if {$cmd in {"abstractNumIds" "delete"}} {
                return
            }
            if {$cmd ne "abstractNum"} {
                return -code error "invalid subcommand \"$cmd\""
            }
            my Add2Relationships numbering numbering.xml
            set docs(word/numbering.xml) [dom parse {
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            }]
            $docs(word/numbering.xml) selectNodesNamespaces $prefixnslist
        }
        set numbering [$docs(word/numbering.xml) documentElement]
        switch $cmd {
            "abstractNum" {
                if {[llength $args] != 2} {
                    return -code error "wrong # of arguments, expecting abstractNumId <list with each element a level description>"
                }
                lassign $args id levelData
                set style [$numbering selectNodes {
                    w:abstractNum[@w:abstractNumId=$id]
                }]
                if {$style ne ""} {
                    error "abstractNum style id $id already exists"
                }
                if {[catch {
                    $numbering insertBeforeFromScript {
                        Tag_w:abstractNum w:abstractNumId $id {
                            set levelnr 0
                            foreach level $levelData {
                                array unset optsknown
                                OptVal $level "option"
                                Tag_w:lvl w:ilvl $levelnr {
                                    set start [my EatOption -start]
                                    if {$start ne ""} {
                                        ST_DecimalNumber $start
                                    } else {
                                        set start 1
                                    }
                                    Tag_w:start w:val $start
                                    my Create $properties(abstractNumStyle)
                                    my PPr
                                    my RPr
                                }
                                if {[catch {my CheckRemainingOpts} errMsg]} {
                                    error "level definition $levelnr: $errMsg"
                                }
                                incr levelnr
                            }
                        }
                    } [$numbering selectNodes {w:num[1]}]
                    $numbering appendFromScript {
                        Tag_w:num w:numId $id {
                            Tag_w:abstractNumId w:val $id
                        }
                    }
                } errMsg]} {
                    return -code error $errMsg
                }
            }
            "abstractNumIds" {
                return [lsort -integer \
                            [$numbering selectNodes -list {
                                w:abstractNum string(@w:abstractNumId)
                            }]]
            }
            "delete" {
                if {[llength $args] != 2} {
                    return -code error "wrong number of arguments\
                               for the subcommand \"delete\", expected\
                               the numbering type and the ID"
                }
                lassign $args type id
                if {$type ni {abstractNum num}} {
                    return -code error "unknown numbering type \"$type\""
                }
                switch $type {
                    "abstractNum" {
                        foreach node [$numbering selectNodes {
                            w:abstractNum[@w:abstractNumId=$id]
                        }] {
                            $node delete
                        }
                    }
                    "num" {
                        foreach node [$numbering selectNodes {
                            w:num[@w:numId=$id]
                        }] {
                            $node delete
                        }
                    }
                }
            }
            default {
                return -code error "invalid subcommand \"$cmd\""
            }
        }
    }

    method pagebreak {} {
        my variable body

        $body appendFromScript {
            Tag_w:p {
                Tag_w:r {
                    Tag_w:br w:type "page"
                }
            }
        }
    }

    # WordprocessingML has no concept of a page. Rather it groups
    # paragraphs (the main building block) into "sections". The page
    # setup of a section is located within the w:p subtree of the last
    # paragraph of the section. And that settings reach out back over
    # all w:p subtrees without page setup.
    method pagesetup {args} {
        my variable body
        my variable setuproot
        my variable pagesetup

        if {[catch {
            OptVal $args
            $setuproot appendFromScript {
                my SectionCommon
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
        [$setuproot lastChild] delete
        set pagesetup $args
    }

    method paragraph {text args} {
        my variable body

        if {[catch {
            OptVal $args "text"
            $body appendFromScript {
                Tag_w:p {
                    my PPr
                    Tag_w:r {
                        my RPr
                        my Wt $text
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method readpart {what file} {
        my variable docs

        set fd [::tdom::xmlOpenFile $file]
        if {[catch {set doc [dom parse -channel $fd]} errMsg]} {
            close $fd
            error $errMsg
        }
        close $fd
        if {[info exists docs($what)]} {
            $docs($what) delete
        }
        $doc selectNodesNamespaces $::ooxml::docx::prefixnslist
        set docs($what) $doc
    }

    method replace {from to {where ""}} {
        my variable docs

        if {$where eq ""} {
            set parts [array names docs]
        } else {
            set parts ""
            foreach this $where {
                if {[info exists docs($this)]} {
                    lappend parts $this
                } else {
                    lappend parts {*}[array names docs $this]
                }
            }
        }
        foreach part $parts {
            set doc $docs($part)
            foreach text [$doc selectNodes {//w:t/text()[contains(.,$from)]}] {
                $text nodeValue [string map [list $from $to] [$text nodeValue]]
            }
        }
    }

    method sectionend {} {
        my variable body
        my variable sectionsetup

        if {$sectionsetup eq ""} {
            return -code error "no section started"
        }
        if {[catch {
            OptVal $sectionsetup
            $body appendFromScript {
                Tag_w:p {
                    Tag_w:pPr {
                        my SectionCommon
                    }
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }
        set sectionsetup ""
    }

    method sectionstart {args} {
        my variable body
        my variable pagesetup
        my variable sectionsetup
        my variable setuproot


        if {[catch {
            # This way in any case a (maybe empty) w:sectPr tag (via
            # SectionCommon) will be created which helps to keep things
            # sane in case there was no pagesetup and no other section
            # before.
            if {$sectionsetup ne ""} {
                OptVal $sectionsetup
            } elseif {$pagesetup ne ""} {
                OptVal $pagesetup
            } else {
                OptVal ""
            }
            $body appendFromScript {
                Tag_w:p {
                    Tag_w:pPr {
                        my SectionCommon
                    }
                }
            }
            # No need for CheckRemainingOpts here because only the
            # already validated options stored in
            # pagesetup/sectionsetup are used.
        } errMsg]} {
            return -code error $errMsg
        }
        # Instead of just storing the arguments actually "evaluate" it
        # her for test to have the error message pointing to the line
        # with the error and not at the place the given arguments are
        # evaluated.
        if {[catch {
            OptVal $args
            $setuproot appendFromScript {
                my SectionCommon
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
        foreach child [$setuproot childNodes] {
            $child delete
        }
        set sectionsetup $args
    }

    method selectNodes {xpath {part ""} args} {
        my variable docs

        if {$part eq ""} {
            set doc $docs(word/document.xml)
        } else {
            if {[info exists docs($part)]} {
                set doc $docs($part)
            } elseif {[info exists docs(word/$part)]} {
                set doc $docs(word/$part)
            } elseif {[info exists docs(word/$part.xml)]} {
                set doc $docs(word/$part.xml)
            } else {
                return -code error "unknown docx part '$part'"
            }
        }
        $doc documentElement root
        if {[catch {
            set result [$root selectNodes {*}$args $xpath]
        } errMsg]} {
            return -code error $errMsg
        }
        return $result
    }

    method settings {args} {
        my variable docs
        variable ::ooxml::docx::properties
        variable ::ooxml::docx::xmlns

        if {[catch {
            OptVal $args
            set create 0
            set reset [my EatOption -reset CT_OnOff 0]
            if {![info exists docs(word/settings.xml)]} {
                my Add2Relationships settings settings.xml
                set create 1
            } else {
                if {$reset} {
                    $docs(word/settings.xml) delete
                    set create 1
                }
            }
            if {$create} {
                set docs(word/settings.xml) [dom parse {
                    <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                        xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"/>
                }]
                $docs(word/settings.xml) selectNodesNamespaces $::ooxml::docx::prefixnslist
            }
            set settings [$docs(word/settings.xml) documentElement]
            $settings lastChild lastchild
            $settings appendFromScript {
                my Create $properties(settings)
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
        # We only have to rearrange the settings children if there has
        # been childs and if there are new childs.
        # nc = new child
        if {$lastchild ne "" && [$lastchild nextSibling nc] ne ""} {
            foreach {opt optdata} $properties(settings) {
                set tag [lindex $optdata 0]
                lassign [split [lindex $optdata 0] :] prefix localname
                set tags($xmlns($prefix):$localname) [incr 1]
            }
            # cc = current child
            $settings firstChild cc
            # Find the first known child
            while {$cc ne $nc} {
                set fqcc [$cc namespaceURI]:[$cc localName]
                if {[info exists tags($fqcc)]} {
                    set ccind $tags($fqcc)
                    break
                }
                $cc nextSibling cc
            }
            if {$cc eq $nc} return
            while {$nc ne ""} {
                $nc nextSibling nextnc
                set ncind $tags([$nc namespaceURI]:[$nc localName])
                while {$ccind < $ncind && $cc ne $nc} {
                    $cc nextSibling cc
                    set fqcc [$cc namespaceURI]:[$cc localName]
                    if {[info exists tags($fqcc)]} {
                        set cclastind $ccind
                        set ccind $tags($fqcc)
                        if {$ccind < $cclastind} {
                            return -code error "invalid word/settings.xml:\
                                                children are not in order"
                        }
                    }
                }
                if {$cc eq $nc} return
                if {$ccind == $ncind} {
                    # TODO this doesn't handle the cases of multiple
                    # child elements nicely (activeWritingStyle,
                    # attachedSchema and smartTagType)
                    $settings replaceChild $nc $cc
                } else {
                    $settings insertBefore $nc $cc
                }
                set cc $nc
                set ccind $ncind
                set nc $nextnc
            }
        }
    }

    method simplecomment {text args} {
        if {[catch {
            OptVal $args "text"
            lassign [my CreateComment] comment id
            $comment appendFromScript {
                Tag_w:p {
                    my PPr
                    Tag_w:r {
                        my RPr
                        my Wt $text
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
        # Add the comment mark to the document
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:r {
                Tag_w:commentReference w:id $id
            }
        }
    }

    method simplecommentrangeend {id text args} {
        my variable commentranges

        if {[catch {
            if {![info exists commentranges($id)]} {
                error "no open comment range with the id '$id'"
            }
            OptVal $args "text"
            lassign [my CreateComment $id] comment
            $comment appendFromScript {
                Tag_w:p {
                    my PPr
                    Tag_w:r {
                        my RPr
                        my Wt $text
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
        # Add the comment mark to the document
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:commentRangeEnd w:id $id
            Tag_w:r {
                Tag_w:commentReference w:id $id
            }
        }
        unset commentranges($id)
    }

    method simpletable {tabledata args} {
        my variable body

        if {[catch {
            OptVal $args "tabledata"
            set firstStyle [my EatOption -firstStyle]
            set lastStyle [my EatOption -lastStyle]
            $body appendFromScript {
                Tag_w:tbl {
                    my TblPr
                    Tag_w:tblGrid {
                        foreach width [my EatOption -columnwidths] {
                            Tag_w:gridCol w:w [ST_TwipsMeasure $width]
                        }
                    }
                    set lastrow [expr {[llength $tabledata] - 1}]
                    for {set i 0} {$i <= $lastrow} {incr i} {
                        set row [lindex $tabledata $i]
                        Tag_w:tr {
                            foreach cell $row {
                                Tag_w:tc {
                                    Tag_w:p {
                                        if {$i == 0 && $firstStyle ne ""} {
                                            Tag_w:pPr {
                                                Tag_w:pStyle w:val $firstStyle
                                            }
                                        } elseif {$i == $lastrow && $lastStyle ne ""} {
                                            Tag_w:pPr {
                                                Tag_w:pStyle w:val $lastStyle
                                            }
                                        }
                                        Tag_w:r {my Wt $cell}
                                    }
                                }
                            }
                        }
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method style {cmd args} {
        my variable docs

        set result ""
        if {[catch {
            set styles [$docs(word/styles.xml) documentElement]
            switch $cmd {
                "paragraphdefault" {
                    set docDefaults [my GetDocDefault $styles]
                    set pdefault [$docDefaults selectNodes {
                        w:pPrDefault | w:rPrDefault
                    }]
                    foreach this $pdefault {
                        $this delete
                    }
                    # docDefaults has two children in the order:
                    # rPrDefault pPrDefault
                    OptVal $args "paragraphdefault"
                    $docDefaults appendFromScript {
                        Tag_w:rPrDefault {
                            my RPr
                        }
                        Tag_w:pPrDefault {
                            my PPr
                        }
                    }
                    my CheckRemainingOpts
                }
                "characterdefault" {
                    set docDefaults [my GetDocDefault $styles]
                    set rdefault [$docDefaults selectNodes w:rPrDefault]
                    foreach this $rdefault {
                        $this delete
                    }
                    # docDefaults has two children in the order:
                    # rPrDefault pPrDefault
                    OptVal $args "characterdefault"
                    $docDefaults insertBeforeFromScript {
                        Tag_w:rPrDefault {
                            my RPr
                        }
                    } [$docDefaults firstChild]
                    my CheckRemainingOpts
                }
                "character" -
                "paragraph" -
                "table" {
                    if {![llength $args]} {
                        error "missing the style name argument"
                    }
                    set name [lindex $args 0]
                    OptVal [lrange $args 1 end] $cmd
                    # Check for duplicate by display name
                    set style [$styles selectNodes {
                        w:style[@w:type=$cmd][w:name[@w:val=$name]]
                    }]
                    if {$style ne ""} {
                        error "$cmd style \"$name\" already exists"
                    }
                    # Derive styleId: use -styleid if given, otherwise
                    # strip non-alphanumeric characters from the name
                    set styleid [my EatOption -styleid]
                    if {$styleid eq ""} {
                        set styleid [regsub -all {[^[:alnum:]]} $name ""]
                    }
                    if {$styleid eq ""} {
                        error "cannot derive a non-empty styleId from\
                               \"$name\"; use the -styleid option"
                    }
                    # Check for styleId collision
                    set clash [$styles selectNodes {
                        w:style[@w:type=$cmd and @w:styleId=$styleid]
                    }]
                    if {$clash ne ""} {
                        error "$cmd style with styleId \"$styleid\" already\
                               exists; use the -styleid option to specify\
                               a unique ID"
                    }
                    # Resolve -basedon to a valid parent styleId
                    set basedon [my EatOption -basedon]
                    if {$basedon ne ""} {
                        set basedon [my StyleCheck $cmd $basedon]
                    }
                    $styles appendFromScript {
                        Tag_w:style w:type $cmd w:styleId $styleid {
                            Tag_w:name w:val $name
                            if {$basedon ne ""} {
                                Tag_w:basedOn w:val $basedon
                            }
                            if {$cmd in {paragraph table}} {
                                my PPr
                            }
                            my RPr
                            if {$cmd eq "table"} {
                                my TblPr
                                my TrPr
                                my TcPr
                                my TblStylePr [my EatOption -conditional]
                            }
                        }
                    }
                    my CheckRemainingOpts
                }
                "ids" {
                    if {[llength $args] != 1} {
                        error "wrong number of arguments, expected the style type"
                    }
                    set type [lindex $args 0]
                    set result [$styles selectNodes -list {
                        w:style[@w:type=$type]
                        string(@w:styleId)
                    }]
                }
                "names" {
                    if {[llength $args] != 1} {
                        error "wrong number of arguments, expected the style type"
                    }
                    set type [lindex $args 0]
                    set result [$styles selectNodes -list {
                        w:style[@w:type=$type]
                        string(w:name/@w:val)
                    }]
                }
                "deleteByName" -
                "delete" {
                    if {[llength $args] == 1]} {
                        lassign $args type
                        if {$type ni {paragraphdefault characterdefault}} {
                            error "unknown default style type \"$type\""
                        }
                    } elseif {[llength $args] == 2} {
                        lassign $args type name
                        if {$type ni {paragraph character table}} {
                            error "unknown style type \"type\""
                        }
                    } else {
                        error "wrong number of arguments for the subcommand \"delete\", expected\
                           the style type and - if not a default type - the style ID"
                    }
                    switch $type {
                        "paragraphdefault" {
                            foreach node [$styles selectNodes w:docDefaults/w:pPrDefault] {
                                $node delete
                            }
                        }
                        "characterdefault" {
                            foreach node [$styles selectNodes w:docDefaults/w:rPrDefault] {
                                $node delete
                            }
                        }
                        default {
                            if {$cmd eq "delete"} {
                                set nodes [$styles selectNodes {
                                    w:style[@w:type=$type and @w:styleId=$name]
                                }]
                            } else {
                                set nodes [$styles selectNodes {
                                    w:style[@w:type=$type][w:name[@w:val=$name]]
                                }]
                            }         
                            foreach node $nodes {
                                $node delete
                            }
                        }
                    }
                }
                default {
                    error "invalid subcommand \"$cmd\""
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }
        return $result
    }

    method tab {{n 1}} {
        my ControlChar tab $n
    }

    method table {args} {
        my variable body
        my variable tablecontext

        set script [lindex $args end]
        set tablecontext "table"
        if {[catch {
            OptVal [lrange $args 0 end-1] "table" "script"
            $body appendFromScript {
                Tag_w:tbl {
                    my TblPr
                    Tag_w:tblGrid {
                        foreach width [my EatOption -columnwidths] {
                            Tag_w:gridCol w:w [ST_TwipsMeasure $width]
                        }
                    }
                    uplevel [list eval $script]
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            set tablecontext ""
            return -code error $errMsg
        }
        set tablecontext ""
    }

    method tablecell {args} {
        my variable body
        my variable tablecontext

        if {$tablecontext ne "row"} {
            error "method tablecell called outside of table row script context"
        }
        set script [lindex $args end]
        OptVal [lrange $args 0 end-1] "tablecell" "script"
        set tablecontext ""
        if {[catch {
            Tag_w:tc {
                my TcPr
                my CheckRemainingOpts
                set savedbody $body
                set body [dom fromScriptContext]
                if {[catch {
                    uplevel [list eval $script]
                } errMsg errVals]} {
                    set body $savedbody
                    set tablecontext "row"
                    my ProcessErrorinfo "tablecell"
                    error $errMsg
                }
                set body $savedbody
            }
        } errMsg]} {
            set tablecontext "row"
            return -code error $errMsg
        }
        set tablecontext "row"
    }

    method tablerow {args} {
        my variable tablecontext

        if {$tablecontext ne "table"} {
            return -code error "method tablerow called outside\
                                of table script context"
        }
        set tablecontext "row"
        set script [lindex $args end]
        OptVal [lrange $args 0 end-1] "tablerow" "script"
        if {[catch {
            Tag_w:tr {
                my TrPr
                my CheckRemainingOpts
                if {[catch {
                    uplevel [list eval $script]
                } errMsg errVals]} {
                    my ProcessErrorinfo "tablerow"
                    error $errMsg
                }
            }
        } errMsg]} {
            set tablecontext "table"
            return -code error $errMsg
        }
        set tablecontext "table"
    }

    method textbox {args} {
        my variable body

        set p [my LastParagraph 1]
        set script [lindex $args end]
        if {[catch {
            OptVal [lrange $args 0 end-1]
            set name [my EatOption -name]
            if {$name eq ""} {
                set name "Textbox [my NextId textboxes]"
            }
            $p appendFromScript {
                Tag_w:r {
                    Tag_w:drawing {
                        set anchor [my Anchor $name]
                    }
                }
            }
            $anchor appendFromScript {
                Tag_a:graphic {
                    Tag_a:graphicData uri "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" {
                        Tag_wps:wsp {
                            Tag_wps:cNvSpPr txBox 1
                            Tag_wps:spPr {
                                my SpPr_Content
                            }
                            # Optional:
                            # wps:style
                            # wps:extLst
                            Tag_wps:txbx {
                                Tag_w:txbxContent {
                                    set savedbody $body
                                    set body [dom fromScriptContext]
                                    # The nested catch is needed to ensure body is set back
                                    if {[catch {uplevel [list eval $script]} errMsg errVals]} {
                                        set body $savedbody
                                        my ProcessErrorinfo "textbox"
                                        error $errMsg
                                    }
                                    set body $savedbody
                                }
                            }
                            # Defaults
                            array set bodyPrAtts {
                                rot 0
                                spcFirstLastPara 0
                                vertOverflow "overflow"
                                horzOverflow "overflow"
                                vert "horz"
                                wrap "square"
                                lIns 0
                                tIns 0
                                rIns 0
                                bIns 0
                                numCol 1
                                spcCol 0
                                rtlCol 0
                                fromWordArt 0
                                anchor "t"
                                anchorCtr 0
                                forceAA 0
                                compatLnSpc "1"
                            }
                            array set bodyPrAtts [my CheckedAttlist [my EatOption -bodyAtts] {
                                -rot ST_DecimalNumber
                                -spcFirstLastPara CT_Boolean
                                -vertOverflow ST_TextVertOverflowType
                                -horzOverflow ST_TextHorzOverflowType
                                -vert ST_TextVerticalType
                                -wrap ST_TextWrappingType
                                -lIns ST_Emu
                                -tIns ST_Emu
                                -rIns ST_Emu
                                -bIns ST_Emu
                                -numCol ST_DecimalNumber
                                -spcCol ST_Emu
                                -rtlCol CT_Boolean
                                -fromWordArt CT_Boolean
                                -anchor ST_TextAnchoringType
                                -anchorCtr CT_Boolean
                                -forceAA CT_Boolean
                                -compatLnSpc CT_Boolean
                            } -bodyAtts]
                            Tag_wps:bodyPr {*}[array get bodyPrAtts]
                        }
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method url {text url args} {
        set rId [my Add2Relationships hyperlink $url]
        set p [my LastParagraph 1]
        if {[catch {
            OptVal $args "text url"
            $p appendFromScript {
                Tag_w:hyperlink r:id $rId {
                    Tag_w:r {
                        my RPr
                        my Wt $text
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }

    method write {file} {
        my variable body
        my variable docs
        my variable binparts
        my variable pagesetup
        my variable sectionsetup
        my variable impPagesetup
        variable ::ooxml::docx::xmlns

        # Finalize section and do pagesetup
        if {$sectionsetup ne ""} {
            OptVal $sectionsetup
            if {[catch {
                $body appendFromScript {
                    Tag_w:p {
                        Tag_w:pPr {
                            my SectionCommon
                        }
                    }
                }
            } errMsg]} {
                return -code error $errMsg
            }
        }
        set sectionsetup ""
        if {$pagesetup ne ""} {
            OptVal $pagesetup
            if {[catch {
                $body appendFromScript {
                    my SectionCommon
                }
            } errMsg]} {
                return -code error $errMsg
            }
            set appendedPageSetup [$body lastChild]
        } else {
            if {$impPagesetup ne ""} {
                $body appendChild $impPagesetup
            } else {
                # This keeps things sane in case there was no pagesetup
                # but sections inbetween.
                $body appendFromScript {
                    Tag_w:sectPr
                }
            }
            set appendedPageSetup [$body lastChild]
        }

        # Initialize zip file
        set file [string trim $file]
        if {$file eq {}} {
            set file {document.docx}
        }
        if {[file extension $file] ne {.docx}} {
            append file .docx
        }
        if {[catch {open $file wb} zf]} {
            error "cannot open file $file for writing"
        }

        foreach part [array names docs] {
            ::ooxml::Dom2zip $zf $docs($part) $part cd count
        }
        foreach part [array names binparts] {
            append cd [::ooxml::add_str_to_archive $zf $part $binparts($part) "" 1]
            incr count
        }
        # Cleanup the appendedPageSetup node in case that after the
        # document was written more content will be added and than
        # written again.
        if {$pagesetup eq "" && $impPagesetup ne ""} {
            # Detach but preserve for potential re-write
            $body removeChild $appendedPageSetup
        } else {
            $appendedPageSetup delete
        }

        # Finalize zip.
        set cdoffset [tell $zf]
        set endrec [binary format a4ssssiis PK\05\06 0 0 $count $count [string length $cd] $cdoffset 0]
        puts -nonewline $zf $cd
        puts -nonewline $zf $endrec
        close $zf
        return
    }

    method writepart {what file} {
        my variable docs

        if {![info exists docs($what)]} {
            error "unknown part $what"
        }
        set fd [open $file w+]
        fconfigure $fd -encoding utf-8
        $docs($what) asXML -channel $fd
        close $fd
    }

    method xmlparts {{pattern ""}} {
        my variable docs

        if {$pattern eq ""} {
            set pattern *
        }
        return [lsort [array names docs $pattern]]
    }

    method xpath {xpath {part ""} args} {
        if {[catch {
            set result [my selectNodes $xpath $part {*}$args]
        } errMsg]} {
            return -code error $errMsg
        }
        return $result
    }
}

# Include the OOML related methods.
source [file join [file dir [info script]] ooxml-docx-math.tcl]

package provide ooxml::docx 0.6
