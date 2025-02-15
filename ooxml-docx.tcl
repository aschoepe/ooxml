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

namespace eval ::ooxml::docx {

    namespace export docx OptVal NoCheck CT_* ST_* W3CDTF
    
    variable xmlns

    array set xmlns {
        a http://schemas.openxmlformats.org/drawingml/2006/main
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

    set properties(stylerun) {
        -bold {{w:b w:bCs} CT_OnOff}
        -color {w:color ST_HexColor}
        -dstrike {w:dstrike CT_OnOff}
        -font {w:rFonts NoCheck RFonts}
        -fontsize {{w:sz w:szCs} ST_TwipsMeasure}
        -italic {{w:i w:iCs} CT_OnOff}
        -strict {w:strike CT_OnOff}
        -underline {w:u ST_Underline}
    }
    set properties(run) [concat $properties(stylerun) {-style {w:rStyle RStyle}}]
    set properties(styleparagraph) {
        -align {w:jc ST_Jc}
        -indentation {w:ind {
            end ST_SignedTwipsMeasure
            endChars ST_DecimalNumber
            firstLine ST_SignedTwipsMeasure
            firstLineChars ST_DecimalNumber
            hanging ST_SignedTwipsMeasure
            hangingChars ST_DecimalNumber
            start ST_SignedTwipsMeasure
            startChars ST_DecimalNumber}}
        -spacing {w:spacing {
            after ST_TwipsMeasure
            before ST_TwipsMeasure
            line ST_TwipsMeasure}}
    }
    set properties(paragraph) [concat $properties(styleparagraph) {-style {w:pStyle PStyle}}]

    # The content model of sectPr is sequence; keep the options in
    # order (and insert new options at the right place)
    set properties(sectionsetup) {
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
    set properties(xfrm) {
        -dimension {a:ext {
            {width -cx} ST_Emu
            {height -cy} ST_Emu
        }}
    }
    # The content model of tblPr is sequence; keep the options in
    # order (and insert new options at the right place)
    set desclist [list]
    foreach tblBordersOpt {top start left bottom end right insideH insideV} {
        lappend desclist -${tblBordersOpt}border [list w:$tblBordersOpt {
            {type val} ST_Border
            color ST_HexColor
            {borderwidth sz} ST_EighthPointMeasure
            space ST_PointMeasure
        }]
    }
    set properties(tblBorders) $desclist
    
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
        w:numPicBullet w:numPr w:numRestart w:numStart w:numStyleLink
        w:object w:objectEmbed w:objectLink w:odso w:oMath
        w:optimizeForBrowser w:outline w:outlineLvl w:overflowPunct
        w:p w:pageBreakBefore w:panose1 w:paperSrc w:pBdr w:permEnd
        w:permStart w:personal w:personalCompose w:personalReply
        w:pgBorders w:pgMar w:pgNum w:pgNumType w:pgSz w:pict
        w:picture w:pitch w:pixelsPerInch w:placeholder w:pos
        w:position w:pPr w:pPrChange w:pPrDefault
        w:printBodyTextBeforeHeader w:printColBlack w:printerSettings
        w:printFormsData w:printFractionalCharacterWidth
        w:printPostScriptOverText w:printTwoOnOne w:proofErr
        w:proofState w:pStyle w:ptab w:qFormat w:query w:r
        w:readModeInkLockDown w:recipientData w:recipients w:relyOnVML
        w:removeDateAndTime w:removePersonalInformation w:result
        w:revisionView w:rFonts w:richText w:right w:rPr w:rPrChange
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
        w:tabIndex w:table w:tabs w:tag w:targetScreenSz w:tbl
        w:tblBorders w:tblCaption w:tblCellMar w:tblCellSpacing
        w:tblDescription w:tblGrid w:tblGridChange w:tblHeader
        w:tblInd w:tblLayout w:tblLook w:tblOverlap w:tblpPr w:tblPr
        w:tblPrChange w:tblPrEx w:tblPrExChange w:tblStyle
        w:tblStyleColBandSize w:tblStylePr w:tblStyleRowBandSize
        w:tblW w:tc w:tcBorders w:tcFitText w:tcMar w:tcPr
        w:tcPrChange w:tcW w:temporary w:text w:textAlignment
        w:textboxTightWrap w:textDirection w:textInput w:themeFontLang
        w:title w:titlePg w:tl2br w:tmpl w:top w:topLinePunct w:tr
        w:tr2bl w:trackRevisions w:trHeight w:trPr w:trPrChange
        w:truncateFontHeightsLikeWP6 w:txbxContent w:type w:types w:u
        w:udl w:uiPriority w:ulTrailSpace w:underlineTabInNumList
        w:unhideWhenUsed w:uniqueTag w:updateFields
        w:useAltKinsokuLineBreakRules w:useAnsiKerningPairs
        w:useFELayout w:useNormalStyleForList w:usePrinterMetrics
        w:useSingleBorderforContiguousCells w:useWord97LineBreakRules
        w:useWord2002TableStyleRules w:useXSLTWhenSaving w:vAlign
        w:vanish w:vertAlign w:view w:viewMergedData w:vMerge w:w
        w:wAfter w:wBefore w:webHidden w:webSettings w:widowControl
        w:wordWrap w:wpJustification w:wpSpaceWidth w:wrapTrailSpaces
        w:writeProtection w:yearLong w:yearShort w:zoom
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(w) elementNode Tag_$tag
    }
    foreach tag {
        Relationships Relationship
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(rel) elementNode Tag_$tag
    }
    foreach tag {
        wp:anchor
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(wp) elementNode Tag_$tag
    }
    foreach tag {
        a:graphic a:graphicData a:blip a:stretch a:fillRect a:xfrm a:ext a:off
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(a) elementNode Tag_$tag
    }        
    foreach tag {
        pic:pic pic:blipFill pic:nvPicPr pic:cNvPr pic:cNvPicPr pic:spPr
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(pic) elementNode Tag_$tag
    }        
    dom createNodeCmd textNode Text
    namespace export Tag_* Text

    # Method option handling helper procs and value checks follow
    proc OptVal {arglist {prefix ""}} {
        if {[llength $arglist] % 2 != 0} {
            if {$prefix ne ""} {append prefix " "}
            error "invalid arguments: expectecd ${prefix}?-option value ?-option value? .."
        }
        uplevel "array set opts [list $arglist]"
    }

    proc NoCheck {value} {
        return $value
    }

    proc ST_Emu {value} {
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

    proc ST_Border {value} {
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
            apples
            archedScallops
            babyPacifier
            babyRattle
            balloons3Colors
            balloonsHotAir
            basicBlackDashes
            basicBlackDots
            basicBlackSquares
            basicThinLines
            basicWhiteDashes
            basicWhiteDots
            basicWhiteSquares
            basicWideInline
            basicWideMidline
            basicWideOutline
            bats
            birds
            birdsFlight
            cabins
            cakeSlice
            candyCorn
            celticKnotwork
            certificateBanner
            chainLink
            champagneBottle
            checkedBarBlack
            checkedBarColor
            checkered
            christmasTree
            circlesLines
            circlesRectangles
            classicalWave
            clocks
            compass
            confetti
            confettiGrays
            confettiOutline
            confettiStreamers
            confettiWhite
            cornerTriangles
            couponCutoutDashes
            couponCutoutDots
            crazyMaze
            creaturesButterfly
            creaturesFish
            creaturesInsects
            creaturesLadyBug
            crossStitch
            cup
            decoArch
            decoArchColor
            decoBlocks
            diamondsGray
            doubleD
            doubleDiamonds
            earth1
            earth2
            earth3
            eclipsingSquares1
            eclipsingSquares2
            eggsBlack
            fans
            film
            firecrackers
            flowersBlockPrint
            flowersDaisies
            flowersModern1
            flowersModern2
            flowersPansy
            flowersRedRose
            flowersRoses
            flowersTeacup
            flowersTiny
            gems
            gingerbreadMan
            gradient
            handmade1
            handmade2
            heartBalloon
            heartGray
            hearts
            heebieJeebies
            holly
            houseFunky
            hypnotic
            iceCreamCones
            lightBulb
            lightning1
            lightning2
            mapPins
            mapleLeaf
            mapleMuffins
            marquee
            marqueeToothed
            moons
            mosaic
            musicNotes
            northwest
            ovals
            packages
            palmsBlack
            palmsColor
            paperClips
            papyrus
            partyFavor
            partyGlass
            pencils
            people
            peopleWaving
            peopleHats
            poinsettias
            postageStamp
            pumpkin1
            pushPinNote2
            pushPinNote1
            pyramids
            pyramidsAbove
            quadrants
            rings
            safari
            sawtooth
            sawtoothGray
            scaredCat
            seattle
            shadowedSquares
            sharksTeeth
            shorebirdTracks
            skyrocket
            snowflakeFancy
            snowflakes
            sombrero
            southwest
            stars
            starsTop
            stars3d
            starsBlack
            starsShadowed
            sun
            swirligig
            tornPaper
            tornPaperBlack
            trees
            triangleParty
            triangles
            triangle1
            triangle2
            triangleCircle1
            triangleCircle2
            shapes1
            shapes2
            twistedLines1
            twistedLines2
            vine
            waveline
            weavingAngles
            weavingBraid
            weavingRibbon
            weavingStrips
            whiteFlowers
            woodwork
            xIllusions
            zanyTriangles
            zigZag
            zigZagStitch
        }
        if {$value in $values} {
            return $value
        }
        error "unknown border type value \"$value\", expected one of:\
               [my AllowedValues $values]"
    }

    
    proc CT_OnOff {value} {
        if {![string is boolean -strict $value]} {
            error "expected a Tcl boolean value"
        }
        if {$value} {
            return "on"
        } else {
            return "off"
        }
    }

    proc ST_DecimalNumber {value} {
        if {![string is integer -strict $value]} {
            error "expected integer but got \"$value\""
        }
        return $value
    }

    # ST_HpsMeasure accepts exactly the same value as ST_TwipsMeasure.
    # The difference is only the interpretation of the integer (only)
    # values. For ST_HpsMeasure the number specifies half points
    # (1/144 of an inch), for ST_TwipsMeasure the number specifies
    # twentieths of a point (equivalent to 1/1440th of an inch).
    proc ST_TwipsMeasure {value} {
        if {[string is integer -strict $value] && $value >= 0} {
            return $value
        }
        if {![regexp {[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
            error "\"$value\" is not a valid measure value - value must match\
               the regular expression \[0-9\]+(\.\[0-9\]+)?(mm|cm|in|pt|pc|pi)"
        }
        return $value
    }

    proc ST_SignedTwipsMeasure {value} {
        if {[string is integer -strict $value]} {
            return $value
        }
        if {![regexp -- {-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
            error "\"$value\" is not a valid measure value - value must match\
               the regular expression \[0-9\]+(\.\[0-9\]+)?(mm|cm|in|pt|pc|pi)"
        }
        return $value
    }

    proc ST_PageOrientation {value} {
        if {$value ni {landscape portrait}} {
            error "unknown symbol \"$value\", expected \"landscape\"\
                   or \"portrait\""
        }
        return $value
    }

    proc ST_Jc {value} {
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
               [my AllowedValues $values]"
    }
    
    proc ST_HexColor {value} {
        if {$value ne "auto"} {
            return $value
        }
        if {[string length $value] != 6 || ![string is xdigit]} {
            error "unknown color value \"$value\", should be \"auto\" or a hex value in\
                   RRGGBB format."
        }
        return $value
    }

    proc ST_EighthPointMeasure {value} {
        if {[string is integer -strict $value] && $value >= 0} {
            return $value
        }
        error "\"$value\" is not a valid measure value - value must be an\
               integer"
    }

    proc ST_PointMeasure {value} {
        if {[string is integer -strict $value] && $value >= 0} {
            return $value
        }
        error "\"$value\" is not a valid measure value - value must be an\
               integer"
    }
    
    proc ST_Underline {value} {
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
                  [my AllowedValues $values]"
        }
        return $value
    }

    proc ST_BlackWhiteMode {value} {
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
            [my AllowedValues $values]"
    }

    proc W3CDTF {value} {
        if {[catch {
            set value [clock format [clock scan $value] -format %Y-%m-%dT%H:%M:%SZ -gmt 1]
        }]} {
            error "invalid datetime value \"$value\", expected everything which is\
                accepted by \[clock scan\]"
        }
        return $value
    }
}

oo::class create ooxml::docx::docx {

    constructor { args } {
        my variable docs
        my variable body
        my variable media
        my variable setuproot
        my variable pagesetup
        my variable sectionsetup
        
        variable ::ooxml::docx::xmlns
        variable ::ooxml::docx::staticDocx

        namespace import ::ooxml::docx::*
        # namespace import ::ooxml::docx::Tag_*  ::ooxml::docx::Text
        # namespace import ::ooxml::docx::

        foreach auxFile [array names staticDocx] {
            set docs($auxFile) [dom parse $staticDocx($auxFile)]
        }
        set document [dom createDocument w:document]
        set docs(word/document.xml) $document
        $document documentElement root
        foreach ns {o r v w w10 wp wps wpg mc wp14 w14 } {
            $root setAttributeNS "" xmlns:$ns $xmlns($ns)
        }
        $root setAttributeNS $xmlns(mc) mc:Ignorable "w14 wp14"
        $root appendFromScript Tag_w:body
        set body [$root firstChild]
        set media ""

        # Since the "general page setup" (WordprocessingML does not
        # really have a concept for that) is a child of w:body after
        # the last paragraph it seemed handy to not realy insert that
        # into the tree at the moment the user calls pagesetup but
        # just before serializing - as the user consecutive add
        # content we can just append to the w:body without looking at
        # every place if there is already a page setup child and we
        # have to insert new content before that element. Long story
        # short - and yes it smells like a hack, there should be a
        # more simpler and natural way, perhaps there is a reason for
        # this overlong comment - wie need an umbrella node with the
        # necessary XML namespaces to that wie do not have unnecessary
        # XML namespaces declarations in the serialized docx.
        set setuproot [$document createElementNS $xmlns(w) w:umbrella]
        set pagesetup ""
        set sectionsetup ""

        my configure {*}$args
    }

    destructor {
        my variable docs

        foreach part [array names docs] {
            $docs($part) delete
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
    
    method import {what docxfile} {

    }

    method readpart {what file} {
        my variable docs
        variable ::ooxml::docx::xmlns

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
        # If pagesetup or sectionsetup are not empty they stored a node
        # out of the fragment list of document and they are now gone
        # with the deletion of document
        if {$what eq "word/document.xml"} {
            my variable setuproot
            my variable pagesetup
            my variable sectionsetup
            set setuproot [$docs(word/document.xml) createElementNS $xmlns(w) w:umbrella]
            set pagesetup ""
            set sectionsetup ""
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


    method PStyle {value} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {[$styles selectNodes {w:style[@w:type="paragraph" and @w:styleId=$value]}] eq ""} {
            error "unknown paragraph style \"$value\""
        }
        return $value
    }

    method RStyle {value} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {[$styles selectNodes {w:style[@w:type="character" and @w:styleId=$value]}] eq ""} {
            error "unknown character style \"$value\""
        }
        return $value
    }

    method RFonts {value} {
        Tag_w:rFonts \
            [my Watt ascii] $value \
            [my Watt hAnsi] $value \
            [my Watt eastAsia] $value \
            [my Watt cs] $value
    }

    method AllowedValues {values {word "or"}} {
        return "[join [lrange $values 0 end-1] ", "] $word [lindex $values end]"
    }
    
    method CreateWorker {value opt optdata} {
        upvar opts opts
        
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
                    return -code continue
                }
                # If we stumble about a tag with just one
                # attribute to set and that attribute is not w:val
                # this case has be handled here.
                #
                # For now the code assumes always several atts
                # and therefore the value to the option is always
                # handled as a key value pairs list.
                set attlist ""
                array unset atts
                if {[catch {array set atts $value}]} {
                    set keys ""
                    foreach {attdata type} $attdefs {
                        lappend keys [lindex $attdata 0]
                    }
                    error "the value given to the \"$opt\" option is\
                           invalid, expected ist a key value pairs\
                           list with keys out of\
                           [my AllowedValues $keys]"
                }
                foreach {attdata type} $attdefs {
                    if {[llength $attdata] == 2} {
                        lassign $attdata key attname
                    } else {
                        set key $attdata
                        set attname $key
                    }
                    if {![info exists atts($key)]} {
                        continue
                    }
                    set ooxmlvalue [my CallType $type $atts($key) \
                                        "the argument \"$value\" given to the \"$opt\"\
                             option is invalid: the value given to the key\
                             \"$key\" in the argument is invalid"]
                    if {[string index $attname 0] eq "-"} {
                        lappend attlist [string range $attname 1 end] $ooxmlvalue
                    } else {
                        lappend attlist [my Watt $attname] $ooxmlvalue
                    }
                    unset atts($key)
                }
                foreach tag $tags {
                    Tag_$tag {*}$attlist {}
                }
                unset opts($opt)                    
                # Check if there are unknown keys left in the key
                # values list
                set remainigKeys [array names atts]
                if {[llength $remainigKeys] == 0} {
                    return -code continue
                }
                set keys ""
                foreach {attdata type} $attdefs {
                    lappend keys [lindex $attdata 0]
                }
                if {[llength $remainigKeys] == 1} {
                    error "unknown key \"[lindex $remainigKeys 0]\" in\
                               the value \"$value\" of the option \"$opt\",\
                               the expected keys are [my AllowedValues $keys]"
                } else {
                    error "unknown keys [my AllowedValues $remainigKeys] in\
                               the value \"$value\" of the option \"$opt\",\
                               the expected keys are [my AllowedValues $keys]"
                }
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
    
    method Create switchActionList {
        upvar opts opts
        array set switchesData $switchActionList

        set switches [array names switchesData]
        foreach {opt value} [array get opts] {
            if {$opt ni $switches} {
                continue
            }
            set optdata $switchesData($opt)
            my CreateWorker $value $opt $optdata
        }
    }

    method CreateOrder switchActionList {
        upvar opts opts

        foreach {opt optdata} $switchActionList {
            if {![info exists opts($opt)]} {
                continue
            }
            set value $opts($opt)
            my CreateWorker $value $opt $optdata
        }
    }
    
    method CheckRemainingOpts {} {
        upvar opts opts

        set nrRemainigOpts [llength [array names opts]]
        if {$nrRemainigOpts == 0} return
        if {$nrRemainigOpts == 1} {
            set text "unknown option: [lindex [array names opts] 0]"
        } else {
            set text "unknown options: [my AllowedValues [array names opts] and]"
        }
        uplevel [list error $text]
    }
    
    method Wt {text} {
        # Just not exactly that easy.
        # Handle at least \n \r \t and \f special
        set pos 0
        set end [string length $text]
        foreach part [split $text "\n\r\t\f"] {
            if {$pos == [incr pos [string length $part]]} {
                incr pos
                continue
            }
            set atts ""
            if {[string index $part 0] eq " " || [string index $part end] eq " "} {
                lappend atts xml:space preserve
            }
            Tag_w:t $atts {
                Text [dom clearString -replace $part]
            }
            if {$pos < $end} {
                switch [string index $text $pos] {
                    "\n" Tag_w:br
                    "\r" Tag_w:cr
                    "\t" Tag_w:tab
                    "\f" {Tag_w:br [my Watt type] "page" {}}
                }
            }
            incr pos
        }
    }

    method Watt {attname} {
        list http://schemas.openxmlformats.org/wordprocessingml/2006/main w:$attname
    }

    method Ratt {attname} {
        list http://schemas.openxmlformats.org/officeDocument/2006/relationships r:$attname
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

    method LastParagraph {{returnEmpty 0}} {
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
        if {$p eq "" && !$returnEmpty} {
            # Or create a new one?
            error "no paragraph to append to in the document"
        }
        return $p
    }
    
    method paragraph {text args} {
        my variable body

        OptVal $args "text"
        $body appendFromScript {
            Tag_w:p {
                Tag_w:pPr {
                    my Create $::ooxml::docx::properties(paragraph)
                }
                Tag_w:r {
                    my Wt $text
                }
            }
        }
        my CheckRemainingOpts
    }
    
    method CallType {type value errtext} {
        # A few value checks need access to the docx object internal
        # data and therefor are implemented as object methods. The
        # bulk of the type checks are "static" procs in the
        # ooxml::docx namespace.
        set error 0
        if {[catch {set ooxmlvalue [$type $value]} errMsg]} {
            if {![llength [info procs ::ooxml::docx::$type]]} {
                if {[catch {set ooxmlvalue [my $type $value]} errMsg]} {set error 1}
            } else {set error 1}
            if {$error} {
                error "$errtext: $errMsg"
            }
        }
        return $ooxmlvalue
    }
    
    method Option {option attname type {default ""}} {
        upvar opts opts
        if {[info exists opts($option)]} {
            set ooxmlvalue [my CallType $type $opts($option) \
                                "the value given to the \"$option\" option\
                                 is invalid"]
            unset opts($option)
        } else {
            return ""
        }
        return [list $attname $ooxmlvalue]
    }
    
    method image {file args} {
        my variable media
        variable ::ooxml::docx::properties
        
        OptVal $args "file"
        lappend media $file
        set rId [my Add2Relationships image media/$file]
        set p [my LastParagraph]
        $p appendFromScript {
            Tag_w:r {
                Tag_w:drawing {
                    Tag_wp:anchor {
                        Tag_a:graphic {
                            Tag_a:graphicData uri "http://schemas.openxmlformats.org/drawingml/2006/picture" {
                                Tag_pic:pic {
                                    Tag_pic:blipFill {
                                        Tag_a:blip [my Ratt embed] rId$rId {}
                                    }
                                    Tag_pic:spPr {*}[my Option -bwMode bwMode ST_BlackWhiteMode "auto"] {
                                        Tag_a:xfrm {
                                            my Create $properties(xfrm)
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        my CheckRemainingOpts
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
        } errMsg]} {
            uplevel error [list $errMsg]
        }
        my CheckRemainingOpts
    }

    method style {cmd args} {
        my variable docs
        
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
                    Tag_w:pPrDefault {
                        Tag_w:pPr {
                            my Create $::ooxml::docx::properties(styleparagraph)
                        }
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
                            my Create $::ooxml::docx::properties(stylerun)
                        }
                    }
                } [$docDefaults firstChild]
                my CheckRemainingOpts
            }
            "character" -
            "paragraph" {
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
                $styles appendFromScript {
                    Tag_w:style [my Watt type] $cmd [my Watt styleId] $name {
                        Tag_w:name [my Watt val] $name {}
                        if {$cmd eq "paragraph"} {
                            Tag_w:pPr {
                                my Create $::ooxml::docx::properties(styleparagraph)
                            }
                        }
                        Tag_w:rPr {
                            my Create $::ooxml::docx::properties(stylerun)
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
                          " the style type and the style ID"
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
    }

    method EatOption {option} {
        upvar opts opts
        if {[info exists opts($option)]} {
            set value $opts($option)
            unset opts($option)
            return $value
        }
        return ""
    }

    method simpletable {tabledata args} {
        my variable body
        variable ::ooxml::docx::properties

        OptVal $args "tabledata"
        $body appendFromScript {
            Tag_w:tbl {
                Tag_w:tblPr {
                    Tag_w:tblBorders {
                        my CreateOrder $properties(tblBorders)
                    }
                }
                set widths [my EatOption -columnwidths]
                if {[llength $widths]} {
                    Tag_w:tblGrid {
                        foreach width $widths {
                            Tag_w:gridCol [my Watt w] [ST_TwipsMeasure $width] {}
                        }
                    }
                }
                foreach row $tabledata {
                    Tag_w:tr {
                        foreach cell $row {
                            Tag_w:tc {
                                # Tag_w:tcPr {
                                #     Tag_w:tcW w:w 200 w:type "dxa"
                                # }
                                Tag_w:p {
                                    Tag_w:r {my Wt $cell}
                                }
                            }
                        }
                    }
                }
            }
        }
        my CheckRemainingOpts
    }

    # WordprocessingML has no concept of a page. Rather it groups
    # paragraphs (the main building block) into "sections". The page
    # setup of a section is located within the w:p subtree of the last
    # paragraph of the section. And that settings reach out _back_
    # over all w:p subtrees without page setup.
    method pagesetup {args} {
        my variable docs
        my variable body
        my variable setuproot
        my variable pagesetup
        variable ::ooxml::docx::xmlns

        if {$pagesetup ne ""} {
            $pagesetup delete
        }
        $setuproot appendFromScript Tag_w:sectPr
        set pagesetup [$setuproot lastChild]
        array set opts $args
        $pagesetup appendFromScript {
            my CreateOrder $::ooxml::docx::properties(sectionsetup)
        }
        my CheckRemainingOpts
    }

    method sectionstart {args} {
        my variable docs
        my variable body
        my variable pagesetup
        my variable sectionsetup
        variable ::ooxml::docx::xmlns

        $body appendFromScript {
            Tag_w:p {
                Tag_w:pPr {
                    if {$sectionsetup ne ""} {
                        ::tdom::fsinsertNode $sectionsetup
                    } elseif {$pagesetup ne ""} {
                        ::tdom::fsinsertNode [$pagesetup clone -deep]
                    }
                }
            }
        }
        $docs(word/document.xml) createElementNS $xmlns(w) w:sectPr sectionsetup
        array set opts $args
        $sectionsetup appendFromScript {
            my CreateOrder $::ooxml::docx::properties(sectionsetup)
        }
        my CheckRemainingOpts
    }

    method sectionend {} {
        my variable body
        my variable pagesetup
        my variable sectionsetup
        
        if {$sectionsetup eq ""} {
            error "no section started"
        }
        $body appendFromScript {
            Tag_w:p {
                Tag_w:pPr {
                    ::tdom::fsinsertNode $sectionsetup
                }
            }
        }
        set sectionsetup ""
    }

    method Add2Relationships {type target} {
        my variable docs
        variable ::ooxml::docx::xmlns

        set relsRoot [$docs(word/_rels/document.xml.rels) documentElement]
        # The following is perhaps over-complicated:
        # set relsns http://schemas.openxmlformats.org/package/2006/relationships
        # set ids [$relsRoot selectNodes -namespaces [list r $relsns] \
        #              -list {r:Relationship string(@Id)}]
        # set rId 1
        # foreach id $ids {
        #     set nr [string range $id 3 end]
        #     if {[string range $id 0 2] eq "rId"
        #         && [string is integer -strict $nr]
        #     } {
        #         if {$nr > $rId} {
        #             set rId $nr
        #         }
        #     }
        # }
        # At least for documents we build up from scratch this should
        # work reliable enough (and work faster)
        set lastchild [$relsRoot lastChild]
        set rId [string range [$lastchild @Id] 3 end]
        incr rId
        set attlist [list \
             Id rId$rId \
             Type http://schemas.openxmlformats.org/officeDocument/2006/relationships/$type \
             Target $target]
        if {$type eq "hyperlink"} {
            lappend attlist TargetMode External
        }
        $relsRoot appendFromScript {
            Tag_Relationship {*}$attlist 
        }
        return $rId
    }
    
    method url {text url args} {
        my variable docs
        variable ::ooxml::docx::xmlns

        OptVal $args "text url"
        set rId [my Add2Relationships hyperlink $url]
        set p [my LastParagraph]
        if {[catch {
            $p appendFromScript {
                Tag_w:hyperlink [list $xmlns(r) r:id] rId$rId {
                    Tag_w:r {
                        Tag_w:rPr {
                            my Create $::ooxml::docx::properties(run)
                        }
                        my Wt $text
                    }
                }
            }
        } errMsg]} {
            uplevel error [list $errMsg]
        }
        my CheckRemainingOpts
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
            set p [my LastParagraph 1]
            if {$p ne ""} {
                $p appendChild $sectionsetup
                set sectionsetup ""
            } else {
                $sectionsetup delete
            }
        }
        set appendedPageSetup ""
        if {$pagesetup ne ""} {
            $body appendChild [$pagesetup clone -deep]
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
        foreach this $media {
            append cd [::ooxml::add_file_with_path_to_archive $zf word/media/[file tail $this] $this]
            incr count
        }
        # Cleanup the appendedPageSetup node in case that after the
        # document was written more content will be added and than
        # writen again.
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
