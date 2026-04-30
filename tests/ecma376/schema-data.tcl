# ecma376/schema-data.tcl --
#
#   Machine-generated from ECMA-376 5th Edition RNC schemas:
#     wml.rnc (WordprocessingML)
#     shared-commonSimpleTypes.rnc (shared simple types)
#
#   DO NOT EDIT — regenerate from RNC sources.

namespace eval ecma376::schema {

    # ── Enumeration types ({name} → {allowed values}) ────────
    variable enums
    array set enums {
        s_ST_AlgClass {hash custom}
        s_ST_AlgType {typeAny custom}
        s_ST_CalendarType {gregorian gregorianUs gregorianMeFrench gregorianArabic hijri hebrew taiwan japan thai korea saka gregorianXlitEnglish gregorianXlitFrench none}
        s_ST_ConformanceClass {strict transitional}
        s_ST_CryptProv {rsaAES rsaFull custom}
        s_ST_OnOff1 {on off}
        s_ST_TrueFalse {t f true false}
        s_ST_TrueFalseBlank {t f true false  True False}
        s_ST_VerticalAlignRun {baseline superscript subscript}
        s_ST_XAlign {left center right inside outside}
        s_ST_YAlign {inline top center bottom inside outside}
        w_ST_AnnotationVMerge {cont rest}
        w_ST_Border {nil none single thick double dotted dashed dotDash dotDotDash triple thinThickSmallGap thickThinSmallGap thinThickThinSmallGap thinThickMediumGap thickThinMediumGap thinThickThinMediumGap thinThickLargeGap thickThinLargeGap thinThickThinLargeGap wave doubleWave dashSmallGap dashDotStroked threeDEmboss threeDEngrave outset inset apples archedScallops babyPacifier babyRattle balloons3Colors balloonsHotAir basicBlackDashes basicBlackDots basicBlackSquares basicThinLines basicWhiteDashes basicWhiteDots basicWhiteSquares basicWideInline basicWideMidline basicWideOutline bats birds birdsFlight cabins cakeSlice candyCorn celticKnotwork certificateBanner chainLink champagneBottle checkedBarBlack checkedBarColor checkered christmasTree circlesLines circlesRectangles classicalWave clocks compass confetti confettiGrays confettiOutline confettiStreamers confettiWhite cornerTriangles couponCutoutDashes couponCutoutDots crazyMaze creaturesButterfly creaturesFish creaturesInsects creaturesLadyBug crossStitch cup decoArch decoArchColor decoBlocks diamondsGray doubleD doubleDiamonds earth1 earth2 earth3 eclipsingSquares1 eclipsingSquares2 eggsBlack fans film firecrackers flowersBlockPrint flowersDaisies flowersModern1 flowersModern2 flowersPansy flowersRedRose flowersRoses flowersTeacup flowersTiny gems gingerbreadMan gradient handmade1 handmade2 heartBalloon heartGray hearts heebieJeebies holly houseFunky hypnotic iceCreamCones lightBulb lightning1 lightning2 mapPins mapleLeaf mapleMuffins marquee marqueeToothed moons mosaic musicNotes northwest ovals packages palmsBlack palmsColor paperClips papyrus partyFavor partyGlass pencils people peopleWaving peopleHats poinsettias postageStamp pumpkin1 pushPinNote2 pushPinNote1 pyramids pyramidsAbove quadrants rings safari sawtooth sawtoothGray scaredCat seattle shadowedSquares sharksTeeth shorebirdTracks skyrocket snowflakeFancy snowflakes sombrero southwest stars starsTop stars3d starsBlack starsShadowed sun swirligig tornPaper tornPaperBlack trees triangleParty triangles triangle1 triangle2 triangleCircle1 triangleCircle2 shapes1 shapes2 twistedLines1 twistedLines2 vine waveline weavingAngles weavingBraid weavingRibbon weavingStrips whiteFlowers woodwork xIllusions zanyTriangles zigZag zigZagStitch custom}
        w_ST_BrClear {none left right all}
        w_ST_BrType {page column textWrapping}
        w_ST_CaptionPos {above below left right}
        w_ST_ChapterSep {hyphen period colon emDash enDash}
        w_ST_CharacterSpacing {doNotCompress compressPunctuation compressPunctuationAndJapaneseKana}
        w_ST_CombineBrackets {none round square angle curly}
        w_ST_Direction {ltr rtl}
        w_ST_DisplacedByCustomXml {next prev}
        w_ST_DocGrid {default lines linesAndChars snapToChars}
        w_ST_DocPartBehavior {content p pg}
        w_ST_DocPartGallery {placeholder any default docParts coverPg eq ftrs hdrs pgNum tbls watermarks autoTxt txtBox pgNumT pgNumB pgNumMargins tblOfContents bib custQuickParts custCoverPg custEq custFtrs custHdrs custPgNum custTbls custWatermarks custAutoTxt custTxtBox custPgNumT custPgNumB custPgNumMargins custTblOfContents custBib custom1 custom2 custom3 custom4 custom5}
        w_ST_DocPartType {none normal autoExp toolbar speller formFld bbPlcHdr}
        w_ST_DocProtect {none readOnly comments trackedChanges forms}
        w_ST_DropCap {none drop margin}
        w_ST_EdGrp {none everyone administrators contributors editors owners current}
        w_ST_EdnPos {sectEnd docEnd}
        w_ST_Em {none dot comma circle underDot}
        w_ST_FFTextType {regular number date currentTime currentDate calculated}
        w_ST_FldCharType {begin separate end}
        w_ST_FontFamily {decorative modern roman script swiss auto}
        w_ST_FrameLayout {rows cols none}
        w_ST_FrameScrollbar {on off auto}
        w_ST_FtnEdn {normal separator continuationSeparator continuationNotice}
        w_ST_FtnPos {pageBottom beneathText sectEnd docEnd}
        w_ST_HAnchor {text margin page}
        w_ST_HdrFtr {even default first}
        w_ST_HeightRule {auto exact atLeast}
        w_ST_HighlightColor {black blue cyan green magenta red yellow white darkBlue darkCyan darkGreen darkMagenta darkRed darkYellow darkGray lightGray none}
        w_ST_Hint {default eastAsia}
        w_ST_InfoTextType {text autoText}
        w_ST_Jc {start center end both mediumKashida distribute numTab highKashida lowKashida thaiDistribute left right}
        w_ST_JcTable {center end left right start}
        w_ST_LevelSuffix {tab space nothing}
        w_ST_LineNumberRestart {newPage newSection continuous}
        w_ST_LineSpacingRule {auto exact atLeast}
        w_ST_Lock {sdtLocked contentLocked unlocked sdtContentLocked}
        w_ST_MailMergeDest {newDocument printer email fax}
        w_ST_MailMergeDocType {catalog envelopes mailingLabels formLetters email fax}
        w_ST_MailMergeOdsoFMDFieldType {null dbColumn}
        w_ST_MailMergeSourceType {database addressBook document1 document2 text email native legacy master}
        w_ST_Merge {continue restart}
        w_ST_MultiLevelType {singleLevel multilevel hybridMultilevel}
        w_ST_NumberFormat {decimal upperRoman lowerRoman upperLetter lowerLetter ordinal cardinalText ordinalText hex chicago ideographDigital japaneseCounting aiueo iroha decimalFullWidth decimalHalfWidth japaneseLegal japaneseDigitalTenThousand decimalEnclosedCircle decimalFullWidth2 aiueoFullWidth irohaFullWidth decimalZero bullet ganada chosung decimalEnclosedFullstop decimalEnclosedParen decimalEnclosedCircleChinese ideographEnclosedCircle ideographTraditional ideographZodiac ideographZodiacTraditional taiwaneseCounting ideographLegalTraditional taiwaneseCountingThousand taiwaneseDigital chineseCounting chineseLegalSimplified chineseCountingThousand koreanDigital koreanCounting koreanLegal koreanDigital2 vietnameseCounting russianLower russianUpper none numberInDash hebrew1 hebrew2 arabicAlpha arabicAbjad hindiVowels hindiConsonants hindiNumbers hindiCounting thaiLetters thaiNumbers thaiCounting bahtText dollarText custom}
        w_ST_ObjectDrawAspect {content icon}
        w_ST_ObjectUpdateMode {always onCall}
        w_ST_PTabAlignment {left center right}
        w_ST_PTabLeader {none dot hyphen underscore middleDot}
        w_ST_PTabRelativeTo {margin indent}
        w_ST_PageBorderDisplay {allPages firstPage notFirstPage}
        w_ST_PageBorderOffset {page text}
        w_ST_PageBorderZOrder {front back}
        w_ST_PageOrientation {portrait landscape}
        w_ST_Pitch {fixed variable default}
        w_ST_Proof {clean dirty}
        w_ST_ProofErr {spellStart spellEnd gramStart gramEnd}
        w_ST_RestartNumber {continuous eachSect eachPage}
        w_ST_RubyAlign {center distributeLetter distributeSpace left right rightVertical}
        w_ST_SdtDateMappingType {text date dateTime}
        w_ST_SectionMark {nextPage nextColumn continuous evenPage oddPage}
        w_ST_Shd {nil clear solid horzStripe vertStripe reverseDiagStripe diagStripe horzCross diagCross thinHorzStripe thinVertStripe thinReverseDiagStripe thinDiagStripe thinHorzCross thinDiagCross pct5 pct10 pct12 pct15 pct20 pct25 pct30 pct35 pct37 pct40 pct45 pct50 pct55 pct60 pct62 pct65 pct70 pct75 pct80 pct85 pct87 pct90 pct95}
        w_ST_StyleSort {name priority default font basedOn type 0000 0001 0002 0003 0004 0005}
        w_ST_StyleType {paragraph character table numbering}
        w_ST_TabJc {clear start center end decimal bar num left right}
        w_ST_TabTlc {none dot hyphen underscore heavy middleDot}
        w_ST_TargetScreenSz {544x376 640x480 720x512 800x600 1024x768 1152x882 1152x900 1280x1024 1600x1200 1800x1440 1920x1200}
        w_ST_TblLayoutType {fixed autofit}
        w_ST_TblOverlap {never overlap}
        w_ST_TblStyleOverrideType {wholeTable firstRow lastRow firstCol lastCol band1Vert band2Vert band1Horz band2Horz neCell nwCell seCell swCell}
        w_ST_TblWidth {nil pct dxa auto}
        w_ST_TextAlignment {top center baseline bottom auto}
        w_ST_TextDirection {tb rl lr tbV rlV lrV btLr lrTb lrTbV tbLrV tbRl tbRlV}
        w_ST_TextEffect {blinkBackground lights antsBlack antsRed shimmer sparkle none}
        w_ST_TextboxTightWrap {none allLines firstAndLastLine firstLineOnly lastLineOnly}
        w_ST_Theme {majorEastAsia majorBidi majorAscii majorHAnsi minorEastAsia minorBidi minorAscii minorHAnsi}
        w_ST_ThemeColor {dark1 light1 dark2 light2 accent1 accent2 accent3 accent4 accent5 accent6 hyperlink followedHyperlink none background1 text1 background2 text2}
        w_ST_Underline {single words double thick dotted dottedHeavy dash dashedHeavy dashLong dashLongHeavy dotDash dashDotHeavy dotDotDash dashDotDotHeavy wave wavyHeavy wavyDouble none}
        w_ST_VAnchor {text margin page}
        w_ST_VerticalJc {top center both bottom}
        w_ST_View {none print outline masterPages normal web}
        w_ST_WmlColorSchemeIndex {dark1 light1 dark2 light2 accent1 accent2 accent3 accent4 accent5 accent6 hyperlink followedHyperlink}
        w_ST_Wrap {auto notBeside around tight through none}
        w_ST_Zoom {none fullPage bestFit textFit}
    }

    # ── Element child ordering (ECMA-376 sequence constraints) ──
    variable childOrder
    array set childOrder {
        pPr {pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle}
        rPr {rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath}
        sectPrContents {footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings}
        tblPr {tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription}
        settings {writeProtection view zoom removePersonalInformation removeDateAndTime doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge revisionView trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing drawingGridVerticalSpacing displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars rsids attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator}
    }

    # ── Attribute → enum type mapping ──────────────────────────
    # Maps w:attrName → enum type name for attributes that
    # take enumerated values.
    variable attrEnum
    array set attrEnum {
        accent1 w_ST_WmlColorSchemeIndex
        accent2 w_ST_WmlColorSchemeIndex
        accent3 w_ST_WmlColorSchemeIndex
        accent4 w_ST_WmlColorSchemeIndex
        accent5 w_ST_WmlColorSchemeIndex
        accent6 w_ST_WmlColorSchemeIndex
        alignment w_ST_PTabAlignment
        asciiTheme w_ST_Theme
        bg1 w_ST_WmlColorSchemeIndex
        bg2 w_ST_WmlColorSchemeIndex
        chapSep w_ST_ChapterSep
        clear w_ST_BrClear
        combineBrackets w_ST_CombineBrackets
        conformance s_ST_ConformanceClass
        cryptAlgorithmClass s_ST_AlgClass
        cryptAlgorithmType s_ST_AlgType
        cryptProviderType s_ST_CryptProv
        cstheme w_ST_Theme
        displacedByCustomXml w_ST_DisplacedByCustomXml
        display w_ST_PageBorderDisplay
        drawAspect w_ST_ObjectDrawAspect
        dropCap w_ST_DropCap
        eastAsiaTheme w_ST_Theme
        edGrp w_ST_EdGrp
        edit w_ST_DocProtect
        fldCharType w_ST_FldCharType
        fmt w_ST_NumberFormat
        followedHyperlink w_ST_WmlColorSchemeIndex
        grammar w_ST_Proof
        hAnchor w_ST_HAnchor
        hAnsiTheme w_ST_Theme
        hRule w_ST_HeightRule
        hint w_ST_Hint
        horzAnchor w_ST_HAnchor
        hyperlink w_ST_WmlColorSchemeIndex
        leader w_ST_TabTlc
        lineRule w_ST_LineSpacingRule
        numFmt w_ST_NumberFormat
        offsetFrom w_ST_PageBorderOffset
        orient w_ST_PageOrientation
        pos w_ST_CaptionPos
        relativeTo w_ST_PTabRelativeTo
        restart w_ST_LineNumberRestart
        sep w_ST_ChapterSep
        spelling w_ST_Proof
        t1 w_ST_WmlColorSchemeIndex
        t2 w_ST_WmlColorSchemeIndex
        tblpXSpec s_ST_XAlign
        tblpYSpec s_ST_YAlign
        themeColor w_ST_ThemeColor
        themeFill w_ST_ThemeColor
        type w_ST_TblStyleOverrideType
        updateMode w_ST_ObjectUpdateMode
        vAnchor w_ST_VAnchor
        vMerge w_ST_AnnotationVMerge
        vMergeOrig w_ST_AnnotationVMerge
        val w_ST_Border
        vertAnchor w_ST_VAnchor
        wrap w_ST_Wrap
        xAlign s_ST_XAlign
        yAlign s_ST_YAlign
        zOrder w_ST_PageBorderZOrder
    }
}
