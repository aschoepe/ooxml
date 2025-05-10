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
package require ::ooxml::docx::lib

namespace eval ::ooxml::docx {

    namespace export docx OptVal NoCheck CT_* ST_* W3CDTF
    
    variable xmlns

    array set xmlns {
        a http://schemas.openxmlformats.org/drawingml/2006/main
        ct http://schemas.openxmlformats.org/package/2006/content-types
        mc http://schemas.openxmlformats.org/markup-compatibility/2006
        o urn:schemas-microsoft-com:office:office
        pic http://schemas.openxmlformats.org/drawingml/2006/picture
        r http://schemas.openxmlformats.org/officeDocument/2006/relationships
        rel http://schemas.openxmlformats.org/package/2006/relationships
        v urn:schemas-microsoft-com:vml
        w http://schemas.openxmlformats.org/wordprocessingml/2006/main
        w10 urn:schemas-microsoft-com:office:word
        w14 http://schemas.microsoft.com/office/word/2010/wordml
        wp http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
        wp14 http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
        wpg http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
        wps http://schemas.microsoft.com/office/word/2010/wordprocessingShape
    }

    # Most WordprocessingML elements have a sequence content model. So
    # the basic rule is to keep the order of the options below as is
    # (and to care that new options are inserted at the right place).
    # Exceptions are locally noted.

    set properties(abstractNumStyle) {
        -numberFormat {w:numFmt ST_NumberFormat}
        -levelText {w:lvlText NoCheck}
        -align {w:lvlJc ST_Jc}
    }
    
    set properties(cell) {
        -cellWidth {w:tcW {
            type ST_TblWidth
            {value w} ST_MeasurementOrPercent}}
        -span {w:gridSpan ST_DecimalNumber}
        -hspan {w:hMerge ST_Merge}
        -vspan {w:vMerge ST_Merge}
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
            anchorLock ST_OnOff
        }}
        +w:numPr {
            -level {w:ilvl ST_DecimalNumber}
            -numberingStyle {w:numId ST_DecimalNumber}
        }
    }
        
    set properties(paragraph2) {
        -pstyle {w:pStyle PStyle}
        -spacing {w:spacing {
            after ST_TwipsMeasure
            before ST_TwipsMeasure
            line ST_TwipsMeasure}}
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
        -cantSpit {w:cantSplit CT_OnOff}
        -rowHeight {w:trHeight {
            {value val} ST_TwipsMeasure
            hRule ST_HeightRule}}
    }
    
    # Unspecified order
    set properties(run) {
        -bold {{w:b w:bCs} CT_OnOff}
        -color {w:color ST_HexColor}
        -dstrike {w:dstrike CT_OnOff}
        -font {w:rFonts NoCheck RFonts}
        -fontsize {{w:sz w:szCs} ST_TwipsMeasure}
        -highlight {w:highlight ST_HighlightColor}
        -italic {{w:i w:iCs} CT_OnOff}
        -strike {w:strike CT_OnOff}
        -cstyle {w:rStyle RStyle}
        -underline {w:u ST_Underline}
    }

    set properties(sectionsetup1) {
        -sizeAndOrientaion {w:pgSz {
            {height h} ST_TwipsMeasure
            {orientation orient} ST_PageOrientation
            {width w} ST_TwipsMeasure}}
        -margins {w:pgMar {
            bottom ST_TwipsMeasure
            footer ST_TwipsMeasure
            gutter ST_TwipsMeasure
            header ST_TwipsMeasure
            left ST_TwipsMeasure
            right ST_TwipsMeasure
            top ST_TwipsMeasure}}
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
                <w:evenAndOddHeaders/>
                <w:compat>
                    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
                </w:compat>
            </w:settings>
        }
        word/styles.xml {
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"/>
        }
    } {
        set ::ooxml::docx::staticDocx($name) $xml
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
        w:emboss w:enabled w:encoding w:end w:endnote w:endnotePr
        w:endnoteRef w:endnoteReference w:endnotes w:entryMacro
        w:equation w:evenAndOddHeaders w:exitMacro w:family w:ffData
        w:fHdr w:fieldMapData w:fitText w:flatBorders w:fldChar
        w:fldData w:fldSimple w:font w:fonts w:footerReference
        w:footnote w:footnoteLayoutLikeWW8 w:footnotePr w:footnoteRef
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
        w:tblHeader w:tblInd w:tblLayout w:tblLook w:tblOverlap
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
        w:tabs w:pPr w:rPr w:tblCellMar w:tblBorders w:tblPr w:tcBorders
        w:tcMar w:tcPr w:trPr w:numPr w:pBdr
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(w) -notempty elementNode Tag_$tag
    }
    
    foreach tag {
        Relationships Relationship
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(rel) elementNode Tag_$tag
    }
    
    foreach tag {
        wp:align wp:anchor wp:docPr wp:effectExtent wp:extent
        wp:inline wp:positionH wp:positionV wp:posOffset wp:simplePos
        wp:wrapNone wp:wrapSquare wp:wrapTopAndBottom
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(wp) elementNode Tag_$tag
    }
    
    foreach tag {
        a:avLst a:blip a:ext a:fillRect a:graphic a:graphicData a:off
        a:picLocks a:prstGeom a:stretch a:xfrm
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(a) elementNode Tag_$tag
    }
    
    foreach tag {
        pic:pic pic:blipFill pic:nvPicPr pic:cNvPr pic:cNvPicPr pic:spPr
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(pic) elementNode Tag_$tag
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
        my variable media
        my variable setuproot
        my variable pagesetup
        my variable sectionsetup
        my variable tablecontext
        
        variable ::ooxml::docx::xmlns
        variable ::ooxml::docx::staticDocx

        namespace import ::ooxml::docx::lib::*
        namespace import ::ooxml::docx::Tag_*  ::ooxml::docx::Text

        foreach auxFile [array names staticDocx] {
            set docs($auxFile) [dom parse $staticDocx($auxFile)]
        }
        set document [dom createDocumentNS $xmlns(w) w:document]
        set docs(word/document.xml) $document
        $document documentElement root
        foreach ns {o r v w10 wp wps wpg mc wp14 w14 } {
            $root setAttributeNS "" xmlns:$ns $xmlns($ns)
        }
        $root setAttributeNS $xmlns(mc) mc:Ignorable "w14 wp14"
        $root appendFromScript Tag_w:body
        set body [$root firstChild]
        set media ""

        # Since the "general page setup" (WordprocessingML does not
        # really have a concept for that) is a child of w:body after
        # the last paragraph it seems handy to not realy insert that
        # into the tree at the moment the user calls pagesetup but
        # just before serializing - as the user consecutive add
        # content we can just append to the w:body without looking at
        # every place if there is already a page setup child and we
        # have to insert new content before that element. The current
        # pagesetup/sectionsetup user definition is storend in
        # according variables to be applied later. To give the user
        # error feedback at the place he provides commands an
        # auxiliary document is used and the definition is "tested"
        # with that.
        set setupdoc [dom createDocumentNS $xmlns(w) w:umbrella]
        set setuproot [$setupdoc documentElement]
        foreach ns {o r v w w10 wp wps wpg mc wp14 w14 } {
            $setuproot setAttributeNS "" xmlns:$ns $xmlns($ns)
        }
        $setuproot setAttributeNS $xmlns(mc) mc:Ignorable "w14 wp14"
        set pagesetup ""
        set sectionsetup ""
        set tablecontext ""
        
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
        if {[$ctRoot selectNodes -namespaces [list ct $xmlns(ct)] {
            count(ct:Override[@Partname=$file])
        }] > 0} {
            return
        }
        if {[string range $file 0 11] eq "/word/media/"} {
            set suffix [file extension $file]
            if {[$ctRoot selectNodes -namespaces [list ct $xmlns(ct)] {
                count(ct:Default[@Extenxtion=$suffix])
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
                set suffix [string range [file extension $file] 1 end]
                switch $suffix {
                    "jpeg" -
                    "jpg" {
                        set ContentType "image/jpeg"
                    }
                    "png" {
                        set ContentType "image/png"
                    }
                }
                $ctRoot insertBeforeFromScript {
                    Tag_Default Extension $suffix ContentType $ContentType
                } [$ctRoot selectNodes -namespaces [list ct $xmlns(ct)] {
                    ct:Override[1]
                }]
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
                "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
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
        variable ::ooxml::docx::xmlns

        set relsRoot [$docs(word/_rels/document.xml.rels) documentElement]

        # If there is already an entry for this type and target just
        # return the rId
        set fqtype http://schemas.openxmlformats.org/officeDocument/2006/relationships/$type
        set result [$relsRoot selectNodes -namespaces "rel $xmlns(rel)" {
            string(rel:Relationship[@Type=$fqtype and @Target=$target]/@Id)
        }]
        if {$result ne ""} {
            return $result
        }

        # At least for documents we build up from scratch this should
        # work reliable enough
        set lastchild [$relsRoot lastChild]
        set nr [string range [$lastchild @Id] 3 end]
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
        my Add2Content_Types "/word/$target"
        return rId$nr
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
                error "$errtext: $errMsg"
            }
        }
        return $ooxmlvalue
    }

    method CheckedAttlist {optionValue attdefs option} {
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
        foreach {attdata type} $attdefs {
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
            if {![info exists atts($key)]} {
                continue
            }
            set ooxmlvalue [my CallType $type $atts($key) \
                                "the argument \"$optionValue\" given to the
                                \"$option\" option is invalid: the value\
                                given to the key \"$key\" in the argument\
                                is invalid"]
            if {[string index $attname 0] eq "-"} {
                lappend attlist [string range $attname 1 end] $ooxmlvalue
            } else {
                lappend attlist w:$attname $ooxmlvalue
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
    
    method Create switchActionList {
        upvar opts opts
        
        foreach {opt optdata} $switchActionList {
            if {[string index $opt 0] eq "+"} {
                Tag_[string range $opt 1 end] {
                    my Create $optdata
                }
                continue
            }
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
                                            "the value \"$value\" given to the \"$opt\"\
                                              option is invalid"]
                        foreach tag $tags {
                            Tag_$tag w:val $ooxmlvalue
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
                            set attname w:$attname
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
    
    method CheckRemainingOpts {} {
        upvar opts opts

        set nrRemainigOpts [llength [array names opts]]
        if {$nrRemainigOpts == 0} return
        if {$nrRemainigOpts == 1} {
            set text "unknown option: [lindex [array names opts] 0]"
        } else {
            set text "unknown options: [AllowedValues [array names opts] and]"
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

    method PeekOption {option {type ""}} {
        upvar opts opts
        return [my EatOption $option $type 0]
    }
        
    method EatOption {option {type ""} {deleteOption 1}} {
        upvar opts opts
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
        return ""
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
        variable ::ooxml::docx::xmlns

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
        set docs(word/$what$nr.xml) $document
        $document documentElement root
        foreach ns {o r v w10 wp wps wpg mc wp14 w14 } {
            $root setAttributeNS {} xmlns:$ns $xmlns($ns)
        }
        $root setAttributeNS $xmlns(mc) mc:Ignorable "w14 wp14"
        set savedbody $body
        set body $root
        if {[catch {uplevel 2 [list eval $script]} errMsg]} {
            set body $savedbody
            return -code error $errMsg
        }
        set body $savedbody
        return $rId
    }

    method Image_anchor {rId file} {
        my variable media
        variable ::ooxml::docx::properties
        upvar opts opts

        array set anchorAtts {
            behindDoc 0
            distT 0
            distB 0
            distL 0
            distR 0
            hidden 0
            simplePos 0
            locked 0
            layoutInCell 0
            allowOverlap 1
            relativeHeight 1
        }
        Tag_wp:anchor [array get anchorAtts] {
            Tag_wp:simplePos x 0 y 0
            Tag_wp:positionH [my Option -positionH relativeFrom ST_RelFromH "column"] {
                set alignH [my EatOption -alignH ST_AlignH]
                set posOffsetH [my EatOption -posOffsetH ST_Emu]
                if {$alignH ne "" && $posOffsetH ne ""} {
                    error "the options -alignH and -posOffsetH are mutually exclusive"
                }
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
                set alignV [my EatOption -alignV ST_AlignV]
                set posOffsetV [my EatOption -posOffsetV ST_Emu]
                if {$alignV ne "" && $posOffsetV ne ""} {
                    error "the options -alignV and -posOffsetV are mutually exclusive"
                }
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
            Tag_wp:extent {*}$attlist
            switch [my EatOption -wrapMode] {
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
            }
            Tag_wp:docPr id [llength $media] name [file tail $file]
            Tag_a:graphic {
                Tag_a:graphicData uri "http://schemas.openxmlformats.org/drawingml/2006/picture" {
                    Tag_pic:pic {
                        Tag_pic:nvPicPr {
                            Tag_pic:cNvPr id [llength $media] name [file rootname [file tail $file]]
                            Tag_pic:cNvPicPr {
                                Tag_a:picLocks  noChangeAspect 1 noChangeArrowheads 1
                            }
                        }
                        Tag_pic:blipFill {
                            Tag_a:blip r:embed $rId
                        }
                        Tag_pic:spPr {*}[my Option -bwMode bwMode ST_BlackWhiteMode "auto"] {
                            Tag_a:xfrm {
                                Tag_a:off x 0 y 0
                                my Create $properties(xfrm)
                            }
                            Tag_a:prstGeom prst "rect" {
                                Tag_a:avLst
                            }
                        }
                    }
                }
            }
        }
    }
    
    method Image_inline {rId file} {
        my variable media
        variable ::ooxml::docx::properties
        upvar opts opts

        Tag_wp:inline {
            set thisOptionValue [my PeekOption -dimension]
            set attlist [my CheckedAttlist $thisOptionValue {
                {width -cx} ST_Emu
                {height -cy} ST_Emu
            } -dimension]
            Tag_wp:extent {*}$attlist
            Tag_wp:docPr id [llength $media] name [file tail $file]
            Tag_a:graphic {
                Tag_a:graphicData uri "http://schemas.openxmlformats.org/drawingml/2006/picture" {
                    Tag_pic:pic {
                        Tag_pic:nvPicPr {
                            Tag_pic:cNvPr id [llength $media] name [file rootname [file tail $file]]
                            Tag_pic:cNvPicPr {
                                Tag_a:picLocks  noChangeAspect 1 noChangeArrowheads 1
                            }
                        }
                        Tag_pic:blipFill {
                            Tag_a:blip r:embed $rId
                        }
                        Tag_pic:spPr {*}[my Option -bwMode bwMode ST_BlackWhiteMode "auto"] {
                            Tag_a:xfrm {
                                Tag_a:off x 0 y 0
                                my Create $properties(xfrm)
                            }
                            Tag_a:prstGeom prst "rect" {
                                Tag_a:avLst
                            }
                        }
                    }
                }
            }
        }
    }
    
    method Image_inline {rId file} {
        my variable media
        variable ::ooxml::docx::properties
        upvar opts opts

    }
    
    # For now only a debuging and developing helper method, not
    # exposed.
    method initWord {docxfile} {
        my variable docs

        package require Tcl 9.0-
        package require fileutil::traverse
        
        set base [file join [zipfs root] InitWord]
        zipfs mount $docxfile $base
        ::fileutil::traverse filesInDocx $base
        #puts [filesInDocx files]
        set prefixlen [string length "$base/word/"]
        set pathstart [string length "$base/"]
        set files [list]
        filesInDocx foreach file {
            if {[string range $file 0 $prefixlen-1] ne "$base/word/"} {
                continue
            }
            if {[file isdirectory $file]} {
                continue
            }
            lappend files [string range $file $pathstart end]
        }
        zipfs unmount $base
        ::ooxml::ZipOpen $docxfile
        foreach file $files {
            if {$file eq "word/document.xml"} {
                continue
            }
            if {[info exists docs($file)]} {
                $docs($file) delete
            }
            set docs($file) [::ooxml::ZipReadParse $file]
        }
        ::ooxml::ZipClose
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

    method TblPr {} {
        variable ::ooxml::docx::properties
        upvar opts opts

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

    method TcPr {} {
        variable ::ooxml::docx::properties
        upvar opts opts

        Tag_w:tcPr {
            my Create $properties(cell)
            Tag_w:tcBorders {
                my Create $properties(cellBorders)
            }
            Tag_w:tcMar {
                my Create $properties(cellMargins)
            }
        }
    }

    method TrPr {} {
        variable ::ooxml::docx::properties
        upvar opts opts

        Tag_w:trPr {
            my Create $properties(row)
        }
    }
        
    method PStyle {value} {
        return [my StyleCheck paragraph $value]
    }

    method Option {option attname type {default ""}} {
        upvar opts opts
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

    method ParagraphStyle {} {
        variable ::ooxml::docx::properties
        upvar opts opts

        Tag_w:pPr {
            my Create $properties(paragraph1)
            Tag_w:pBdr {
                my Create $properties(paragraphBorders)
            }
            Tag_w:tabs {
                my Tabs [my EatOption -tabs]
            }
            my Create $properties(paragraph2)
        }
    }
    
    method RFonts {value} {
        Tag_w:rFonts \
            w:ascii $value \
            w:hAnsi $value \
            w:eastAsia $value \
            w:cs $value
    }
    
    method RStyle {value} {
        return [my StyleCheck character $value]
    }

    method SectionCommon {} {
        variable ::ooxml::docx::properties
        upvar opts opts

        Tag_w:sectPr {
            foreach what {Header Footer} {
                foreach type {even default first} {
                    set value [my EatOption -$type$what]
                    if {$value eq ""} {
                        continue
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

    method StyleCheck {type value} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {[$styles selectNodes {w:style[@w:type=$type and @w:styleId=$value]}] eq ""} {
            error "unknown $type style \"$value\""
        }
        return $value
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
            my CheckRemainingOpts
            unset opts
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
            if {[string length $part]} {
                set atts ""
                if {[string index $part 0] eq " " || [string index $part end] eq " "} {
                    lappend atts xml:space preserve
                }
                Tag_w:t $atts {
                    Text [dom clearString -replace $part]
                }
            }
            if {$pos < $end} {
                switch [string index $text $pos] {
                    "\n" Tag_w:br
                    "\r" Tag_w:cr
                    "\t" Tag_w:tab
                    "\f" {Tag_w:br w:type "page"}
                }
            }
            incr pos
        }
    }

    method append {text args} {
        OptVal $args "text"
        set p [my LastParagraph]
        if {[catch {
            $p appendFromScript {
                Tag_w:r {
                    Tag_w:rPr {
                        my Create $::ooxml::docx::properties(run)
                    }
                    my Wt $text
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
    }
    
    method configure {args} {
        my variable docs

        OptVal $args
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
    }

    method field {field} {
        my variable docs

        # TODO: field formating / switches
        set nfield [string toupper $field]
        # libreoffice doesn't seem to support
        # AUTHOR
        # FILESIZE
        # SAVEDATE
        # SECTION
        # although the spec say it should.
        # TODO: check what word does
        set values {
            AUTHOR
            CREATEDATE
            DATE
            FILESIZE
            PAGE
            SAVEDATE
            SECTION
            TIME
            TITLE
            USERNAME
        }
        if {$nfield ni $values} {
            return -code error "Unknown field type '$field', expected one\
                                out of [AllowedValues $values]"
        }
        set p [my LastParagraph 1]
        $p appendFromScript {
            Tag_w:r {
                Tag_w:fldChar w:fldCharType "begin"
            }
            Tag_w:r {
                Tag_w:instrText {
                    Text $nfield
                }
            }
            Tag_w:r {
                Tag_w:fldChar w:fldCharType "separate"
            }
            Tag_w:r {
                Tag_w:fldChar w:fldCharType "end"
            }
        }
    }

    method footer {script {returnvar ""}} {
        if {$returnvar ne ""} {
            upvar $returnvar result
        }
        set result [my HeaderFooter footer $script]
    }
    
    method header {script {returnvar ""}} {
        if {$returnvar ne ""} {
            upvar $returnvar result
        }
        set result [my HeaderFooter header $script]
    }
    
    method image {file type args} {
        my variable media
        variable ::ooxml::docx::properties

        if {![file isfile $file] || ![file readable $file]} {
            error "çannot read file \"$file\""
        }
        if {$type ni {inline anchor}} {
            error "invalid type \"$type\" (expected \"inline\" or \"anchor\")"
        }
        if {[catch {
            OptVal $args "file"
            set file [file normalize $file]
            set ind [lsearch -exact $media $file]
            if {$ind < 0} {
                lappend media $file
                set ind [llength $media]
            } else {
                # Tcl indexes count from 0, the image numbering from 1
                incr ind
            }
            set imagename image${ind}[file extension $file]
            set rId [my Add2Relationships image media/$imagename]
            set p [my LastParagraph]
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

    method numbering {cmd args} {
        my variable docs
        variable ::ooxml::docx::properties

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
        }
        set numbering [$docs(word/numbering.xml) documentElement]
        switch $cmd {
            "abstractNum" {
                if {[llength $args] != 2} {
                    error "wrong # of argumentes, expecting abstractNumId <list with each element a level description>"
                }
                lassign $args id levelData
                set style [$numbering selectNodes {
                    w:abstractNum[@w:abstractNumId=$id]
                }]
                if {$style ne ""} {
                    error "abstractNum style id $id already exists"
                }
                if {[catch {
                    $numbering appendFromScript {
                        Tag_w:abstractNum w:abstractNumId $id {
                            set levelnr 0
                            foreach level $levelData {
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
                                    my ParagraphStyle
                                    Tag_w:rPr {
                                        my Create $properties(run)
                                    }
                                }
                                if {[catch {my CheckRemainingOpts} errMsg]} {
                                    error "level definition $levelnr: $errMsg"
                                }
                                incr levelnr
                            }
                        }
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
                    error "wrong number of arguments for the subcommand \"delete\", expected\
                               the style type and the style ID"
                }
                lassign $args type id
                if {$type ni {abstractNum num}} {
                    error "unknown style type \"type\""
                }
                switch $type {
                    "abstractNum" {
                        foreach node [$numbering selectNodes {
                            w:abstractNum[@w:abstractNumId=$id]
                        }] {
                            $node delete
                        }
                    }
                }
            }
            default {
                error "invalid subcommand \"$cmd\""
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

        OptVal $args
        if {[catch {
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
        variable ::ooxml::docx::properties
        
        if {[catch {
            OptVal $args "text"
            $body appendFromScript {
                Tag_w:p {
                    my ParagraphStyle
                    Tag_w:r {
                        Tag_w:rPr {
                            my Create $properties(run)
                        }
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
        set docs($what) $doc
    }

    method sectionend {} {
        my variable body
        my variable sectionsetup
        
        if {$sectionsetup eq ""} {
            error "no section started"
        }
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
        set sectionsetup ""
    }

    method sectionstart {args} {
        my variable body
        my variable pagesetup
        my variable sectionsetup
        my variable setuproot

        if {$sectionsetup ne "" || $pagesetup ne ""} {
            if {$sectionsetup ne ""} {
                OptVal $sectionsetup
            } else {
                OptVal $pagesetup
            }
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
        OptVal $args
        if {[catch {
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

    method simpletable {tabledata args} {
        my variable body
        variable ::ooxml::docx::properties

        OptVal $args "tabledata"
        if {[catch {
            set firstStyle [my EatOption -firstStyle]
            set lastStyle [my EatOption -lastStyle]
            $body appendFromScript {
                Tag_w:tbl {
                    my TblPr
                    set widths [my EatOption -columnwidths]
                    if {[llength $widths]} {
                        Tag_w:tblGrid {
                            foreach width $widths {
                                Tag_w:gridCol w:w [ST_TwipsMeasure $width]
                            }
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
        variable ::ooxml::docx::properties

        if {[catch {
            set styles [$docs(word/styles.xml) documentElement]
            switch $cmd {
                "paragraphdefault" {
                    set docDefaults [my GetDocDefault $styles]
                    set pdefault [$docDefaults selectNodes w:pPrDefault]
                    foreach this $pdefault {
                        $pdefault delete
                    }
                    # docDefaults has two childs in the order:
                    # rPrDefault pPrDefault
                    OptVal $args "paragraphdefault"
                    $docDefaults appendFromScript {
                        Tag_w:rPrDefault {
                            Tag_w:rPr {
                                my Create $properties(run)
                            }
                        }
                        Tag_w:pPrDefault {
                            my ParagraphStyle
                        }
                    }
                    my CheckRemainingOpts
                }
                "characterdefault" {
                    set docDefaults [my GetDocDefault $styles]
                    set rdefault [$docDefaults selectNodes w:rPrDefault]
                    foreach this $rdefault {
                        $rdefault delete
                    }
                    # docDefaults has two childs in the order:
                    # rPrDefault pPrDefault
                    OptVal $args "characterdefault"
                    $docDefaults insertBeforeFromScript {
                        Tag_w:rPrDefault {
                            Tag_w:rPr {
                                my Create $properties(run)
                            }
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
                    set style [$styles selectNodes {
                        w:style[@w:type=$cmd and @w:styleId=$name]
                    }]
                    if {$style ne ""} {
                        error "$cmd style \"$name\" already exists"
                    }
                    set basedon [my EatOption -basedon]
                    $styles appendFromScript {
                        Tag_w:style w:type $cmd w:styleId $name {
                            Tag_w:name w:val $name
                            if {$basedon ne ""} {
                                Tag_w:basedOn w:val $basedon
                            }
                            if {$cmd in {paragraph table}} {
                                my ParagraphStyle
                            }
                            Tag_w:rPr {
                                my Create $properties(run)
                            }
                            if {$cmd eq "table"} {
                                my TblPr
                                my TrPr
                                my TcPr
                            }
                        }
                    }
                    my CheckRemainingOpts
                }
                "ids" {
                    if {[llength $args] != 1} {
                        error "wrong number of arguments, expectecd the style type"
                    }
                    set type [lindex $args 0]
                    return [$styles selectNodes -list {w:style[@w:type=$type] string(@w:styleId)}]
                }
                "delete" {
                    if {[llength $args] != 2} {
                        error "wrong number of arguments for the subcommand \"delete\", expected\
                           the style type and the style ID"
                    }
                    lassign $args type name
                    if {$type ni {paragraph paragraphdefault character characterdefault}} {
                        error "unknown style type \"type\""
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
                            foreach node [$styles selectNodes {
                                w:style[@w:type=$type and @w:styleId=$name]
                            }] {
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
    }

    method table {args} {
        my variable body
        my variable tablecontext
        variable ::ooxml::docx::properties

        set script [lindex $args end]
        OptVal [lrange $args 0 end-1]
        set tablecontext "table"
        if {[catch {
            $body appendFromScript {
                Tag_w:tbl {
                    my TblPr
                    set widths [my EatOption -columnwidths]
                    if {[llength $widths]} {
                        Tag_w:tblGrid {
                            foreach width $widths {
                                Tag_w:gridCol w:w [ST_TwipsMeasure $width]
                            }
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
        variable ::ooxml::docx::properties

        if {$tablecontext ne "row"} {
            error "method tablecell called outside of table row script context"
        }
        set script [lindex $args end]
        OptVal [lrange $args 0 end-1]
        set tablecontext ""
        try {
            Tag_w:tc {
                my TcPr
                set savedbody $body
                set body [dom fromScriptContext]
                try {
                    uplevel [list eval $script]
                } finally {
                    set body $savedbody
                }
            }
        } finally {
            set tablecontext "row"
        }
    }
    
    method tablerow {args} {
        my variable body
        my variable tablecontext
        variable ::ooxml::docx::properties

        if {$tablecontext ne "table"} {
            return -code error "method tablerow called outside of table script context"
        }
        set tablecontext "row"
        set script [lindex $args end]
        OptVal [lrange $args 0 end-1]
        try {
            Tag_w:tr {
                my TrPr
                uplevel [list eval $script]
            }
        } finally {
            set tablecontext "table"
        }
    }
    
    method url {text url args} {
        my variable docs
        variable ::ooxml::docx::xmlns

        OptVal $args "text url"
        set rId [my Add2Relationships hyperlink $url]
        set p [my LastParagraph 1]
        if {[catch {
            $p appendFromScript {
                Tag_w:hyperlink r:id $rId {
                    Tag_w:r {
                        Tag_w:rPr {
                            my Create $::ooxml::docx::properties(run)
                        }
                        my Wt $text
                    }
                }
            }
            my CheckRemainingOpts
        } errMsg]} {
            return -code error $errMsg
        }
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
            
    method write {file} {
        my variable body
        my variable docs
        my variable media
        my variable pagesetup
        my variable sectionsetup
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
        set appendedPageSetup ""
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
        }
        
        # Initialize zip file
        set file [string trim $file]
        if {$file eq {}} {
            set file {document.docx}
        }
        if {[file extension $file] ne {.docx}} {
            append file .docx
        }
        if {[catch {open $file w} zf]} {
            error "cannot open file $file for writing"
        }
        fconfigure $zf -translation binary -eofchar {}

        foreach part [array names docs] {
            ::ooxml::Dom2zip $zf $docs($part) $part cd count
        }
        set nr 1
        foreach this $media {
            set medianame word/media/image${nr}[file extension $this]
            append cd [::ooxml::add_file_with_path_to_archive $zf $medianame $this]
            incr count
            incr nr
        }
        # Cleanup the appendedPageSetup node in case that after the
        # document was written more content will be added and than
        # written again.
        if {$appendedPageSetup ne ""} {
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
}

package provide ::ooxml::docx 1.8.1
