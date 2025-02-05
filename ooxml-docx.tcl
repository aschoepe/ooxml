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

    set properties(stylerun) {
        -bold {CT_OnOff {w:b w:bCs}}
        -dstrike {CT_OnOff w:dstrike}
        -font {noCheck w:rFonts rFonts}
        -italic {CT_OnOff {w:i w:iCs}}
        -strict {CT_OnOff w:strike}
    }
    set properties(run) [concat $properties(stylerun) {-style {RStyle w:style}}]
    set properties(styleparagraph) {
        -spacing {noCheck w:spacing spacing}
    }
    set properties(paragraph) [concat $properties(styleparagraph) {-style {PStyle w:pStyle}}]
}

proc ooxml::InitDocxNodeCommands {} {
    variable xmlns
    
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
    dom createNodeCmd textNode Text
    namespace export Tag_* Text
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
                            <w:lang w:val="en" w:eastAsia="zh-CN" w:bidi="hi-IN"/>
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
                        <w:lang w:val="en" w:eastAsia="zh-CN" w:bidi="hi-IN"/>
                    </w:rPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Heading">
                    <w:name w:val="Heading"/>
                    <w:basedOn w:val="Normal"/>
                    <w:next w:val="TextBody"/>
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
                <w:style w:type="paragraph" w:styleId="TextBody">
                    <w:name w:val="Body Text"/>
                    <w:basedOn w:val="Normal"/>
                    <w:pPr>
                        <w:spacing w:lineRule="auto" w:line="276" w:before="0" w:after="140"/>
                    </w:pPr>
                    <w:rPr/>
                </w:style>
                <w:style w:type="paragraph" w:styleId="List">
                    <w:name w:val="List"/>
                    <w:basedOn w:val="TextBody"/>
                    <w:pPr/>
                    <w:rPr>
                        <w:rFonts w:cs="FreeSans"/>
                    </w:rPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Caption">
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
                <w:style w:type="paragraph" w:styleId="Index">
                    <w:name w:val="Index"/>
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
        set ::ooxml::staticDocx($name) $xml
    }
}

::ooxml::InitStaticDocx

oo::class create ooxml::docx_write {

    constructor { args } {
        my variable document
        my variable body
        my variable docs
        variable ::ooxml::xmlns
        variable ::ooxml::staticDocx

        namespace import ::ooxml::Tag_*  ::ooxml::Text

        foreach auxFile [array names staticDocx] {
            set docs($auxFile) [dom parse $staticDocx($auxFile)]
        }
        dom createDocument w:document document
        $document documentElement root
        foreach ns {o r v w w10 wp wps wpg mc wp14 w14 } {
            $root setAttribute xmlns:$ns $xmlns($ns)
        }
        $root setAttribute mc:Ignorable "w14 wp14"
        $root appendFromScript Tag_w:body
        set body [$root firstChild]
    }

    destructor {
        my variable document
        my variable docs

        $document delete
        foreach auxDoc [array names docs] {
            $docs($auxDoc) delete
        }
    }

    method import {what docxfile} {

    }

    method read {what file} {

    }

    method write {what file} {

    }

    method noCheck {value} {
        return $value
    }

    method CT_OnOff {value} {
        if {![string is boolean -strict $value]} {
            error "expected a Tcl boolean value"
        }
        if {$value} {
            return "on"
        } else {
            return "off"
        }
    }

    method ST_TwipsMeasure {value} {
        if {[string is integer -strict $value] && $value >= 0} {
            return $value
        }
        if {![regexp {[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)} $value]} {
            error "\"$value\" is not a valid measure value - value must match\
               the regular expression \[0-9\]+(\.\[0-9\]+)?(mm|cm|in|pt|pc|pi)"
        }
        return $value
    }

    method PStyle {value} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {[$styles selectNodes {w:style[@w:styleId=$value]}] eq ""} {
            error "unknown style \"$value\""
        }
        return $value
    }

    method rFonts {value} {
        Tag_w:rFonts \
            [my watt ascii] $value \
            [my watt hAnsi] $value \
            [my watt eastAsia] $value \
            [my watt cs] $value
    }

    method spacing {value} {
        set errMsg "the value given to the -spacing attribute is invalid,\
                    expected a name value list with name any combination of\
                    after, before or line" 
        if {[catch {array set atts $value}]} {
            error $errMsg
        }
        set attlist [list]
        foreach key [array names atts] {
            if {$key ni {after before line}} {
                error $errMsg
            }
            my ST_TwipsMeasure $atts($key)
            lappend attlist [my watt $key] $atts($key)
        }
        Tag_w:spacing {*}$attlist {}
        
    }
    
    method Create switchActionList {
        upvar opts opts
        array set switchesData $switchActionList

        set switches [array names switchesData]
        foreach {opt value} [array get opts] {
            if {$opt in $switches} {
                lassign $switchesData($opt) check action createCmd
                if {[catch {set ooxmlvalue [my $check $value]} errMsg]} {
                    error "the value \"$value\" given to\
                       the \"$opt\" option is invalid:\
                       $errMsg"
                }
                if {$createCmd eq ""} {
                    foreach tag $action {
                        Tag_$tag w:val $ooxmlvalue
                    }
                } else {
                    my $createCmd $ooxmlvalue
                }
                unset opts($opt)
            }
        }
    }

    method checkRemainingOpts {} {
        upvar opts opts

        if {[llength [array names opts]]} {
            uplevel error "unknown: [array names opts]"
        }
    }
    
    method paragraph {text args} {
        my variable body

        if {[llength $args] % 2 != 0} {
            error "invalid arguments: expectecd -option value pairs"
        }
        array set opts $args
        $body appendFromScript {
            Tag_w:p {
                Tag_w:pPr {
                    my Create $::ooxml::properties(paragraph)
                }
                Tag_w:r {
                    my Wt $text
                }
            }
        }
        my checkRemainingOpts
    }

    method append {text args} {
        my variable body

        if {[llength $args] % 2 != 0} {
            error "invalid arguments: expectecd -option value pairs"
        }
        array set opts $args
        # Identify the last paragraph
        set p [$body lastChild]
        while {$p ne ""} {
            if {[$p nodeType] ne "ELEMENT_NODE"} {
                set p [$p previousSibling]
                continue
            }
            if {[$p nodeName] ne "w:p"} {
                # Perhaps it is wise to raise error in this situation
                # or create a new one?
                set p [$p previousSibling]
                continue
            }
            break
        }
        if {$p eq ""} {
            # Or create a new one?
            error "no paragraph to append to in the document"
        }
        if {[catch {
            $p appendFromScript {
                Tag_w:r {
                    Tag_w:rPr {
                        my Create $::ooxml::properties(run)
                    }
                    my Wt $text
                }
            }
        } errMsg]} {
            uplevel error $errMsg
        }
        my checkRemainingOpts
    }

    method Wt {text} {
        # Just not exactly that easy.
        # Handle at least \n \r \t special
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
                    "\f" {Tag_w:br [my watt type] "page" {}}
                }
            }
            incr pos
        }
    }

    method watt {attname} {
        list http://schemas.openxmlformats.org/wordprocessingml/2006/main w:$attname
    }
    
    method GetDocDefault {styles} {
        set docDefaults [$styles selectNodes -cache 1 {w:docDefaults[1]}]
        if {$node eq ""} {
            set nextNode [$styles selectNodes -cache 1 {*[1]}]
            $styles insertBeforeFromScript Tag_w:docDefaults $nextNode
            set docDefaults [$styles selectNodes -cache 1 {*[1]}]
        }
        return $docDefaults
    }
            
    method style {cmd args} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {$cmd eq "ids"} {
            return [$styles selectNodes -list {w:style string(@w:styleId)}]
        }
        switch $cmd {
            "defaultparagraph" {
                set docDefaults [my GetDocDefault $styles]
            }
            "defaultcharacter" {
                set docDefaults [my GetDocDefault $styles]
            }
            "paragraph" {
                if {![llength $args]} {
                    error "missing paragraph name argument"
                }
                set name [lindex $args 0]
                array set opts [lrange $args 1 end]
                set style [$styles selectNodes {
                    w:style[@w:type="paragraph" and @w:styleId=$name]
                }]
                if {$style ne ""} {
                    error "style \"$name\" already exists"
                }
                $styles appendFromScript {
                    Tag_w:style [my watt type] "paragraph" [my watt styleId] $name {
                        Tag_w:name [my watt val] $name {}
                        Tag_w:pPr {
                            my Create $::ooxml::properties(styleparagraph)
                        }
                        Tag_w:rPr {
                            my Create $::ooxml::properties(stylerun)
                        }
                    }
                }
                my checkRemainingOpts
            }
            "character" {
                if {![llength $args]} {
                    error "missing paragraph name argument"
                }
                set name [lindex $args 0]
                set style [$styles selectNodes {
                    w:style[@w:type="character" and @w:styleId=$name ]
                }]
            }
            default {
                error "invalid subcommand \"$cmd\""
            }
        }
    }

    method simpletable {tabledata} {
        my variable body

        $body appendFromScript {
            Tag_w:tbl {
                Tag_w:tblPr
                # Tag_w:tblGrid {
                #     Tag_w:gridCol w:w 200
                #     Tag_w:gridCol w:w 200
                #     Tag_w:gridCol w:w 200
                # }
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
    }

    method write {file} {
        my variable document
        my variable body
        my variable docs
        variable ::ooxml::xmlns
        variable ::ooxml::staticDocx

        namespace import ::ooxml::Tag_* ::ooxml::Text

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

        foreach auxFile [array names docs] {
            ::ooxml::Dom2zip $zf $docs($auxFile) $auxFile cd count
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
