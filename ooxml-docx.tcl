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

    # Unspecified order
    set properties(stylerun) {
        -bold {{w:b w:bCs} CT_OnOff}
        -color {w:color ST_HexColor}
        -dstrike {w:dstrike CT_OnOff}
        -font {w:rFonts NoCheck RFonts}
        -fontsize {{w:sz w:szCs} ST_TwipsMeasure}
        -italic {{w:i w:iCs} CT_OnOff}
        -strike {w:strike CT_OnOff}
        -underline {w:u ST_Underline}
    }
    set properties(run) [concat $properties(stylerun) {-style {w:rStyle RStyle}}]

    set properties(styleparagraph) {
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
    set properties(paragraph) [concat {-style {w:pStyle PStyle}} $properties(styleparagraph)]

    set properties(xfrm) {
        -dimension {a:ext {
            {width -cx} ST_Emu
            {height -cy} ST_Emu
        }}
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
    } {
        foreach option $borderOptions {
            lappend properties($property) \
                -${option}Border [list w:$option $BorderOpts]
        }
    }

    set properties(sectionsetup2) {
        -pageNumbering {w:pgNumType {
            chapSep ST_ChapterSep
            chapStyle ST_DecimalNumber
            fmt ST_NumberFormat
            start ST_DecimalNumber
        }}
    }
    
    set properties(numbering) {
        -level {w:ilvl ST_DecimalNumber}
        -numberingStyle {w:numId ST_DecimalNumber}
    }

    set properties(abstractNumStyle) {
        -numberFormat {w:numFmt ST_NumberFormat}
        -levelText {w:lvlText NoCheck}
        -align {w:lvlJc ST_Jc}
    }
    
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
                <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
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
        word/numbering.xml {
            <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
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
        w:numPicBullet w:numRestart w:numStart w:numStyleLink
        w:object w:objectEmbed w:objectLink w:odso w:oMath
        w:optimizeForBrowser w:outline w:outlineLvl w:overflowPunct
        w:p w:pageBreakBefore w:panose1 w:paperSrc w:permEnd
        w:permStart w:personal w:personalCompose w:personalReply
        w:pgBorders w:pgMar w:pgNum w:pgNumType w:pgSz w:pict
        w:picture w:pitch w:pixelsPerInch w:placeholder w:pos
        w:position w:pPrChange w:pPrDefault
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
        w:tabIndex w:table w:tag w:targetScreenSz w:tbl
        w:tblCaption w:tblCellMar w:tblCellSpacing
        w:tblDescription w:tblGrid w:tblGridChange w:tblHeader
        w:tblInd w:tblLayout w:tblLook w:tblOverlap w:tblpPr
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
        w:tabs w:pPr w:rPr w:tblBorders w:tblPr w:numPr w:pBdr
    } {
        dom createNodeCmd -tagName $tag -namespace $xmlns(w) -notempty elementNode Tag_$tag
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
    
    method Create switchActionList {
        upvar opts opts
        
        foreach {opt optdata} $switchActionList {
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
                    # If we stumble about a tag with just one
                    # attribute to set and that attribute is not w:val
                    # this case has to be handled here.
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
                           [AllowedValues $keys]"
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
                            lappend attlist w:$attname $ooxmlvalue
                        }
                        unset atts($key)
                    }
                    foreach tag $tags {
                        Tag_$tag {*}$attlist
                    }
                    unset opts($opt)                    
                    # Check if there are unknown keys left in the key
                    # values list
                    set remainigKeys [array names atts]
                    if {[llength $remainigKeys] == 0} {
                        continue
                    }
                    set keys ""
                    foreach {attdata type} $attdefs {
                        lappend keys [lindex $attdata 0]
                    }
                    if {[llength $remainigKeys] == 1} {
                        error "unknown key \"[lindex $remainigKeys 0]\" in\
                               the value \"$value\" of the option \"$opt\",\
                               the expected keys are [AllowedValues $keys]"
                    } else {
                        error "unknown keys [AllowedValues $remainigKeys] in\
                               the value \"$value\" of the option \"$opt\",\
                               the expected keys are [AllowedValues $keys]"
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

    method EatOption {option} {
        upvar opts opts
        if {[info exists opts($option)]} {
            set value $opts($option)
            unset opts($option)
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

        set footer [lsort -dictionary [array names docs word/footer*]]
        if {![llength $footer]} {
            set nr 1
        } else {
            set nr [string range $footer 11 end]
            incr nr
        }
        set rId [my Add2Relationships image word/footer$nr]
        set document [dom createDocumentNS $xmlns(w) w:ftr]
        set docs(word/footer$nr) $document
        $document documentElement root
        foreach ns {o r v w10 wp wps wpg mc wp14 w14 } {
            $root setAttributeNS {} xmlns:$ns $xmlns($ns)
        }
        $root setAttributeNS $xmlns(mc) mc:Ignorable "w14 wp14"
        set savedbody $body
        set body $root
        $body appendFromScript {
            Tag_w:p r:rId foo
        }
        puts [$document asXML]
        if {[catch {uplevel 2 [list eval $script]} errMsg]} {
            set body $savedbody
            return -code error $errMsg
        }
        set body $savedbody
        return $rId
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

    method PStyle {value} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {[$styles selectNodes {w:style[@w:type="paragraph" and @w:styleId=$value]}] eq ""} {
            error "unknown paragraph style \"$value\""
        }
        return $value
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
    
    method RFonts {value} {
        Tag_w:rFonts \
            w:ascii $value \
            w:hAnsi $value \
            w:eastAsia $value \
            w:cs $value
    }
    
    method RStyle {value} {
        my variable docs
        
        set styles [$docs(word/styles.xml) documentElement]
        if {[$styles selectNodes {w:style[@w:type="character" and @w:styleId=$value]}] eq ""} {
            error "unknown character style \"$value\""
        }
        return $value
    }

    method SectionCommon {} {
        variable ::ooxml::docx::properties
        upvar opts opts

        Tag_w:sectPr {
            foreach what {Header Footer} {
                foreach type {even default first} {
                    puts "EatOption -$type$what"
                    set value [my EatOption -$type$what]
                    if {$value eq ""} {
                        continue
                    }
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
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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

    method footer {script} {
        my HeaderFooter footer $script
    }
    
    method header {script} {
        my HeaderFooter header $script
    }
    
    method image {file args} {
        my variable media
        variable ::ooxml::docx::properties
        
        if {[catch {
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
                                            Tag_a:blip r:embed rId$rId
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
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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
                                    Tag_w:pPr {
                                        my Create $properties(styleparagraph)
                                    }
                                    Tag_w:rPr {
                                        my Create $properties(stylerun)
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
    # paragraph of the section. And that settings reach out _back_
    # over all w:p subtrees without page setup.
    method pagesetup {args} {
        my variable body
        my variable setuproot
        my variable pagesetup

        OptVal $args
        if {[catch {
            $setuproot appendFromScript {
                my SectionCommon
            }
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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
                    Tag_w:pPr {
                        Tag_w:numPr {
                            my Create $properties(numbering)
                        }
                        Tag_w:pBdr {
                            my Create $properties(paragraphBorders)
                        }
                        Tag_w:tabs {
                            my Tabs [my EatOption -tabs]
                        }
                        my Create $properties(paragraph)
                    }
                    Tag_w:r {
                        my Wt $text
                    }
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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
            set style [my EatOption -style]
            $body appendFromScript {
                Tag_w:tbl {
                    Tag_w:tblPr {
                        if {$style ne ""} {
                            Tag_w:tblStyle w:val $style
                        }
                        Tag_w:tblBorders {
                            my Create $properties(tableBorders)
                        }
                    }
                    set widths [my EatOption -columnwidths]
                    if {[llength $widths]} {
                        Tag_w:tblGrid {
                            foreach width $widths {
                                Tag_w:gridCol w:w [ST_TwipsMeasure $width]
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
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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
                        Tag_w:pPrDefault {
                            Tag_w:pPr {
                                Tag_w:pBdr {
                                    my Create $properties(paragraphBorders)
                                }
                                my Create $properties(styleparagraph)
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
                                my Create $properties(stylerun)
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
                            Tag_w:basedOn w:val $basedon
                            if {$cmd eq "table"} {
                                Tag_w:tblPr {
                                    Tag_w:tblBorders {
                                        my Create $properties(tableBorders)
                                    }
                                }
                            } else {
                                if {$cmd eq "paragraph"} {
                                    Tag_w:pPr {
                                        Tag_w:pBdr {
                                            my Create $properties(paragraphBorders)
                                        }
                                        my Create $properties(styleparagraph)
                                    }
                                }
                                Tag_w:rPr {
                                    my Create $properties(stylerun)
                                }
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

    method url {text url args} {
        my variable docs
        variable ::ooxml::docx::xmlns

        OptVal $args "text url"
        set rId [my Add2Relationships hyperlink $url]
        set p [my LastParagraph 1]
        if {[catch {
            $p appendFromScript {
                Tag_w:hyperlink r:id rId$rId {
                    Tag_w:r {
                        Tag_w:rPr {
                            my Create $::ooxml::docx::properties(run)
                        }
                        my Wt $text
                    }
                }
            }
        } errMsg]} {
            return -code error $errMsg
        }
        my CheckRemainingOpts
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
        foreach this $media {
            append cd [::ooxml::add_file_with_path_to_archive $zf word/media/[file tail $this] $this]
            incr count
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
