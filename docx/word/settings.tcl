package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName m:brkBin elementNode Tag_m:brkBin
dom createNodeCmd -tagName m:brkBinSub elementNode Tag_m:brkBinSub
dom createNodeCmd -tagName m:defJc elementNode Tag_m:defJc
dom createNodeCmd -tagName m:dispDef elementNode Tag_m:dispDef
dom createNodeCmd -tagName m:intLim elementNode Tag_m:intLim
dom createNodeCmd -tagName m:lMargin elementNode Tag_m:lMargin
dom createNodeCmd -tagName m:mathFont elementNode Tag_m:mathFont
dom createNodeCmd -tagName m:mathPr elementNode Tag_m:mathPr
dom createNodeCmd -tagName m:naryLim elementNode Tag_m:naryLim
dom createNodeCmd -tagName m:rMargin elementNode Tag_m:rMargin
dom createNodeCmd -tagName m:smallFrac elementNode Tag_m:smallFrac
dom createNodeCmd -tagName m:wrapIndent elementNode Tag_m:wrapIndent
dom createNodeCmd -tagName w14:defaultImageDpi elementNode Tag_w14:defaultImageDpi
dom createNodeCmd -tagName w14:docId elementNode Tag_w14:docId
dom createNodeCmd -tagName w15:chartTrackingRefBased elementNode Tag_w15:chartTrackingRefBased
dom createNodeCmd -tagName w15:docId elementNode Tag_w15:docId
dom createNodeCmd -tagName w:characterSpacingControl elementNode Tag_w:characterSpacingControl
dom createNodeCmd -tagName w:clrSchemeMapping elementNode Tag_w:clrSchemeMapping
dom createNodeCmd -tagName w:compat elementNode Tag_w:compat
dom createNodeCmd -tagName w:compatSetting elementNode Tag_w:compatSetting
dom createNodeCmd -tagName w:decimalSymbol elementNode Tag_w:decimalSymbol
dom createNodeCmd -tagName w:defaultTabStop elementNode Tag_w:defaultTabStop
dom createNodeCmd -tagName w:hyphenationZone elementNode Tag_w:hyphenationZone
dom createNodeCmd -tagName w:listSeparator elementNode Tag_w:listSeparator
dom createNodeCmd -tagName w:rsid elementNode Tag_w:rsid
dom createNodeCmd -tagName w:rsidRoot elementNode Tag_w:rsidRoot
dom createNodeCmd -tagName w:rsids elementNode Tag_w:rsids
dom createNodeCmd -tagName w:themeFontLang elementNode Tag_w:themeFontLang
dom createNodeCmd -tagName w:zoom elementNode Tag_w:zoom
set doc [dom createDocument w:settings]
set root [$doc documentElement]
$root setAttribute xmlns:mc http://schemas.openxmlformats.org/markup-compatibility/2006
$root setAttribute xmlns:o urn:schemas-microsoft-com:office:office
$root setAttribute xmlns:r http://schemas.openxmlformats.org/officeDocument/2006/relationships
$root setAttribute xmlns:m http://schemas.openxmlformats.org/officeDocument/2006/math
$root setAttribute xmlns:v urn:schemas-microsoft-com:vml
$root setAttribute xmlns:w10 urn:schemas-microsoft-com:office:word
$root setAttribute xmlns:w http://schemas.openxmlformats.org/wordprocessingml/2006/main
$root setAttribute xmlns:w14 http://schemas.microsoft.com/office/word/2010/wordml
$root setAttribute xmlns:w15 http://schemas.microsoft.com/office/word/2012/wordml
$root setAttribute xmlns:w16cid http://schemas.microsoft.com/office/word/2016/wordml/cid
$root setAttribute xmlns:w16se http://schemas.microsoft.com/office/word/2015/wordml/symex
$root setAttribute xmlns:sl http://schemas.openxmlformats.org/schemaLibrary/2006/main
$root setAttribute mc:Ignorable {w14 w15 w16se w16cid}
$root appendFromScript {
  Tag_w:zoom w:percent 130 {
  }
  Tag_w:defaultTabStop w:val 708 {
  }
  Tag_w:hyphenationZone w:val 425 {
  }
  Tag_w:characterSpacingControl w:val doNotCompress {
  }
  Tag_w:compat {
    Tag_w:compatSetting w:name compatibilityMode w:uri http://schemas.microsoft.com/office/word w:val 15 {
    }
    Tag_w:compatSetting w:name overrideTableStyleFontSizeAndJustification w:uri http://schemas.microsoft.com/office/word w:val 1 {
    }
    Tag_w:compatSetting w:name enableOpenTypeFeatures w:uri http://schemas.microsoft.com/office/word w:val 1 {
    }
    Tag_w:compatSetting w:name doNotFlipMirrorIndents w:uri http://schemas.microsoft.com/office/word w:val 1 {
    }
    Tag_w:compatSetting w:name differentiateMultirowTableHeaders w:uri http://schemas.microsoft.com/office/word w:val 1 {
    }
    Tag_w:compatSetting w:name useWord2013TrackBottomHyphenation w:uri http://schemas.microsoft.com/office/word w:val 0 {
    }
  }
  Tag_w:rsids {
    Tag_w:rsidRoot w:val 00D76278 {
    }
    Tag_w:rsid w:val 001B1FDE {
    }
    Tag_w:rsid w:val 00CB4A40 {
    }
    Tag_w:rsid w:val 00D76278 {
    }
  }
  Tag_m:mathPr {
    Tag_m:mathFont m:val {Cambria Math} {
    }
    Tag_m:brkBin m:val before {
    }
    Tag_m:brkBinSub m:val -- {
    }
    Tag_m:smallFrac m:val 0 {
    }
    Tag_m:dispDef {
    }
    Tag_m:lMargin m:val 0 {
    }
    Tag_m:rMargin m:val 0 {
    }
    Tag_m:defJc m:val centerGroup {
    }
    Tag_m:wrapIndent m:val 1440 {
    }
    Tag_m:intLim m:val subSup {
    }
    Tag_m:naryLim m:val undOvr {
    }
  }
  Tag_w:themeFontLang w:val de-DE {
  }
  Tag_w:clrSchemeMapping w:bg1 light1 w:t1 dark1 w:bg2 light2 w:t2 dark2 w:accent1 accent1 w:accent2 accent2 w:accent3 accent3 w:accent4 accent4 w:accent5 accent5 w:accent6 accent6 w:hyperlink hyperlink w:followedHyperlink followedHyperlink {
  }
  Tag_w:decimalSymbol w:val , {
  }
  Tag_w:listSeparator w:val {;} {
  }
  Tag_w14:docId w14:val 7A816370 {
  }
  Tag_w14:defaultImageDpi w14:val 32767 {
  }
  Tag_w15:chartTrackingRefBased {
  }
  Tag_w15:docId w15:val {{22CAD71F-778E-A14F-AEC6-E0D033380262}} {
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
