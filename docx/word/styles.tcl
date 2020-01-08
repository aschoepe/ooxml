package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName w:bottom elementNode Tag_w:bottom
dom createNodeCmd -tagName w:docDefaults elementNode Tag_w:docDefaults
dom createNodeCmd -tagName w:lang elementNode Tag_w:lang
dom createNodeCmd -tagName w:latentStyles elementNode Tag_w:latentStyles
dom createNodeCmd -tagName w:left elementNode Tag_w:left
dom createNodeCmd -tagName w:lsdException elementNode Tag_w:lsdException
dom createNodeCmd -tagName w:name elementNode Tag_w:name
dom createNodeCmd -tagName w:pPrDefault elementNode Tag_w:pPrDefault
dom createNodeCmd -tagName w:qFormat elementNode Tag_w:qFormat
dom createNodeCmd -tagName w:rFonts elementNode Tag_w:rFonts
dom createNodeCmd -tagName w:rPr elementNode Tag_w:rPr
dom createNodeCmd -tagName w:rPrDefault elementNode Tag_w:rPrDefault
dom createNodeCmd -tagName w:right elementNode Tag_w:right
dom createNodeCmd -tagName w:semiHidden elementNode Tag_w:semiHidden
dom createNodeCmd -tagName w:style elementNode Tag_w:style
dom createNodeCmd -tagName w:sz elementNode Tag_w:sz
dom createNodeCmd -tagName w:szCs elementNode Tag_w:szCs
dom createNodeCmd -tagName w:tblCellMar elementNode Tag_w:tblCellMar
dom createNodeCmd -tagName w:tblInd elementNode Tag_w:tblInd
dom createNodeCmd -tagName w:tblPr elementNode Tag_w:tblPr
dom createNodeCmd -tagName w:top elementNode Tag_w:top
dom createNodeCmd -tagName w:uiPriority elementNode Tag_w:uiPriority
dom createNodeCmd -tagName w:unhideWhenUsed elementNode Tag_w:unhideWhenUsed
set doc [dom createDocument w:styles]
set root [$doc documentElement]
$root setAttribute xmlns:mc http://schemas.openxmlformats.org/markup-compatibility/2006
$root setAttribute xmlns:r http://schemas.openxmlformats.org/officeDocument/2006/relationships
$root setAttribute xmlns:w http://schemas.openxmlformats.org/wordprocessingml/2006/main
$root setAttribute xmlns:w14 http://schemas.microsoft.com/office/word/2010/wordml
$root setAttribute xmlns:w15 http://schemas.microsoft.com/office/word/2012/wordml
$root setAttribute xmlns:w16cid http://schemas.microsoft.com/office/word/2016/wordml/cid
$root setAttribute xmlns:w16se http://schemas.microsoft.com/office/word/2015/wordml/symex
$root setAttribute mc:Ignorable {w14 w15 w16se w16cid}
$root appendFromScript {
  Tag_w:docDefaults {
    Tag_w:rPrDefault {
      Tag_w:rPr {
        Tag_w:rFonts w:asciiTheme minorHAnsi w:eastAsiaTheme minorHAnsi w:hAnsiTheme minorHAnsi w:cstheme minorBidi {
        }
        Tag_w:sz w:val 24 {
        }
        Tag_w:szCs w:val 24 {
        }
        Tag_w:lang w:val de-DE w:eastAsia en-US w:bidi ar-SA {
        }
      }
    }
    Tag_w:pPrDefault {
    }
  }
  Tag_w:latentStyles w:defLockedState 0 w:defUIPriority 99 w:defSemiHidden 0 w:defUnhideWhenUsed 0 w:defQFormat 0 w:count 375 {
    Tag_w:lsdException w:name Normal w:uiPriority 0 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 1} w:uiPriority 9 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 2} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 3} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 4} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 5} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 6} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 7} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 8} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {heading 9} w:semiHidden 1 w:uiPriority 9 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {index 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 6} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 7} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 8} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index 9} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 1} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 2} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 3} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 4} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 5} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 6} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 7} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 8} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toc 9} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Normal Indent} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {footnote text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {annotation text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name header w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name footer w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {index heading} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name caption w:semiHidden 1 w:uiPriority 35 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {table of figures} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {envelope address} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {envelope return} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {footnote reference} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {annotation reference} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {line number} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {page number} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {endnote reference} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {endnote text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {table of authorities} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name macro w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {toa heading} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name List w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Bullet} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Number} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Bullet 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Bullet 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Bullet 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Bullet 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Number 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Number 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Number 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Number 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name Title w:uiPriority 10 w:qFormat 1 {
    }
    Tag_w:lsdException w:name Closing w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name Signature w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Default Paragraph Font} w:semiHidden 1 w:uiPriority 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text Indent} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Continue} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Continue 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Continue 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Continue 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {List Continue 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Message Header} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name Subtitle w:uiPriority 11 w:qFormat 1 {
    }
    Tag_w:lsdException w:name Salutation w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name Date w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text First Indent} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text First Indent 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Note Heading} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text Indent 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Body Text Indent 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Block Text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name Hyperlink w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name FollowedHyperlink w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name Strong w:uiPriority 22 w:qFormat 1 {
    }
    Tag_w:lsdException w:name Emphasis w:uiPriority 20 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Document Map} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Plain Text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {E-mail Signature} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Top of Form} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Bottom of Form} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Normal (Web)} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Acronym} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Address} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Cite} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Code} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Definition} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Keyboard} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Preformatted} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Sample} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Typewriter} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {HTML Variable} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Normal Table} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {annotation subject} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {No List} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Outline List 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Outline List 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Outline List 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Simple 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Simple 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Simple 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Classic 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Classic 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Classic 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Classic 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Colorful 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Colorful 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Colorful 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Columns 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Columns 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Columns 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Columns 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Columns 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 6} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 7} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid 8} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 4} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 5} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 6} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 7} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table List 8} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table 3D effects 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table 3D effects 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table 3D effects 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Contemporary} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Elegant} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Professional} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Subtle 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Subtle 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Web 1} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Web 2} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Web 3} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Balloon Text} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Table Grid} w:uiPriority 39 {
    }
    Tag_w:lsdException w:name {Table Theme} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Placeholder Text} w:semiHidden 1 {
    }
    Tag_w:lsdException w:name {No Spacing} w:uiPriority 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Light Shading} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name {Medium List 2} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Light Shading Accent 1} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List Accent 1} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid Accent 1} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1 Accent 1} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2 Accent 1} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1 Accent 1} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name Revision w:semiHidden 1 {
    }
    Tag_w:lsdException w:name {List Paragraph} w:uiPriority 34 w:qFormat 1 {
    }
    Tag_w:lsdException w:name Quote w:uiPriority 29 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Intense Quote} w:uiPriority 30 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Medium List 2 Accent 1} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1 Accent 1} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2 Accent 1} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3 Accent 1} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List Accent 1} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading Accent 1} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List Accent 1} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid Accent 1} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Light Shading Accent 2} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List Accent 2} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid Accent 2} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1 Accent 2} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2 Accent 2} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1 Accent 2} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name {Medium List 2 Accent 2} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1 Accent 2} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2 Accent 2} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3 Accent 2} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List Accent 2} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading Accent 2} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List Accent 2} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid Accent 2} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Light Shading Accent 3} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List Accent 3} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid Accent 3} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1 Accent 3} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2 Accent 3} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1 Accent 3} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name {Medium List 2 Accent 3} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1 Accent 3} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2 Accent 3} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3 Accent 3} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List Accent 3} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading Accent 3} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List Accent 3} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid Accent 3} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Light Shading Accent 4} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List Accent 4} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid Accent 4} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1 Accent 4} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2 Accent 4} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1 Accent 4} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name {Medium List 2 Accent 4} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1 Accent 4} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2 Accent 4} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3 Accent 4} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List Accent 4} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading Accent 4} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List Accent 4} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid Accent 4} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Light Shading Accent 5} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List Accent 5} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid Accent 5} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1 Accent 5} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2 Accent 5} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1 Accent 5} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name {Medium List 2 Accent 5} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1 Accent 5} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2 Accent 5} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3 Accent 5} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List Accent 5} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading Accent 5} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List Accent 5} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid Accent 5} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Light Shading Accent 6} w:uiPriority 60 {
    }
    Tag_w:lsdException w:name {Light List Accent 6} w:uiPriority 61 {
    }
    Tag_w:lsdException w:name {Light Grid Accent 6} w:uiPriority 62 {
    }
    Tag_w:lsdException w:name {Medium Shading 1 Accent 6} w:uiPriority 63 {
    }
    Tag_w:lsdException w:name {Medium Shading 2 Accent 6} w:uiPriority 64 {
    }
    Tag_w:lsdException w:name {Medium List 1 Accent 6} w:uiPriority 65 {
    }
    Tag_w:lsdException w:name {Medium List 2 Accent 6} w:uiPriority 66 {
    }
    Tag_w:lsdException w:name {Medium Grid 1 Accent 6} w:uiPriority 67 {
    }
    Tag_w:lsdException w:name {Medium Grid 2 Accent 6} w:uiPriority 68 {
    }
    Tag_w:lsdException w:name {Medium Grid 3 Accent 6} w:uiPriority 69 {
    }
    Tag_w:lsdException w:name {Dark List Accent 6} w:uiPriority 70 {
    }
    Tag_w:lsdException w:name {Colorful Shading Accent 6} w:uiPriority 71 {
    }
    Tag_w:lsdException w:name {Colorful List Accent 6} w:uiPriority 72 {
    }
    Tag_w:lsdException w:name {Colorful Grid Accent 6} w:uiPriority 73 {
    }
    Tag_w:lsdException w:name {Subtle Emphasis} w:uiPriority 19 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Intense Emphasis} w:uiPriority 21 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Subtle Reference} w:uiPriority 31 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Intense Reference} w:uiPriority 32 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Book Title} w:uiPriority 33 w:qFormat 1 {
    }
    Tag_w:lsdException w:name Bibliography w:semiHidden 1 w:uiPriority 37 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {TOC Heading} w:semiHidden 1 w:uiPriority 39 w:unhideWhenUsed 1 w:qFormat 1 {
    }
    Tag_w:lsdException w:name {Plain Table 1} w:uiPriority 41 {
    }
    Tag_w:lsdException w:name {Plain Table 2} w:uiPriority 42 {
    }
    Tag_w:lsdException w:name {Plain Table 3} w:uiPriority 43 {
    }
    Tag_w:lsdException w:name {Plain Table 4} w:uiPriority 44 {
    }
    Tag_w:lsdException w:name {Plain Table 5} w:uiPriority 45 {
    }
    Tag_w:lsdException w:name {Grid Table Light} w:uiPriority 40 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light Accent 1} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2 Accent 1} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3 Accent 1} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4 Accent 1} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark Accent 1} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful Accent 1} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful Accent 1} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light Accent 2} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2 Accent 2} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3 Accent 2} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4 Accent 2} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark Accent 2} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful Accent 2} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful Accent 2} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light Accent 3} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2 Accent 3} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3 Accent 3} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4 Accent 3} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark Accent 3} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful Accent 3} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful Accent 3} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light Accent 4} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2 Accent 4} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3 Accent 4} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4 Accent 4} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark Accent 4} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful Accent 4} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful Accent 4} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light Accent 5} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2 Accent 5} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3 Accent 5} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4 Accent 5} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark Accent 5} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful Accent 5} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful Accent 5} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {Grid Table 1 Light Accent 6} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {Grid Table 2 Accent 6} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {Grid Table 3 Accent 6} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {Grid Table 4 Accent 6} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {Grid Table 5 Dark Accent 6} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {Grid Table 6 Colorful Accent 6} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {Grid Table 7 Colorful Accent 6} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light Accent 1} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2 Accent 1} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3 Accent 1} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4 Accent 1} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark Accent 1} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful Accent 1} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful Accent 1} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light Accent 2} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2 Accent 2} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3 Accent 2} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4 Accent 2} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark Accent 2} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful Accent 2} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful Accent 2} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light Accent 3} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2 Accent 3} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3 Accent 3} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4 Accent 3} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark Accent 3} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful Accent 3} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful Accent 3} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light Accent 4} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2 Accent 4} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3 Accent 4} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4 Accent 4} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark Accent 4} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful Accent 4} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful Accent 4} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light Accent 5} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2 Accent 5} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3 Accent 5} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4 Accent 5} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark Accent 5} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful Accent 5} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful Accent 5} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name {List Table 1 Light Accent 6} w:uiPriority 46 {
    }
    Tag_w:lsdException w:name {List Table 2 Accent 6} w:uiPriority 47 {
    }
    Tag_w:lsdException w:name {List Table 3 Accent 6} w:uiPriority 48 {
    }
    Tag_w:lsdException w:name {List Table 4 Accent 6} w:uiPriority 49 {
    }
    Tag_w:lsdException w:name {List Table 5 Dark Accent 6} w:uiPriority 50 {
    }
    Tag_w:lsdException w:name {List Table 6 Colorful Accent 6} w:uiPriority 51 {
    }
    Tag_w:lsdException w:name {List Table 7 Colorful Accent 6} w:uiPriority 52 {
    }
    Tag_w:lsdException w:name Mention w:semiHidden 1 w:unhideWhenUsed 1 {
    }
    Tag_w:lsdException w:name {Smart Hyperlink} w:semiHidden 1 w:unhideWhenUsed 1 {
    }
  }
  Tag_w:style w:type paragraph w:default 1 w:styleId Standard {
    Tag_w:name w:val Normal {
    }
    Tag_w:qFormat {
    }
  }
  Tag_w:style w:type character w:default 1 w:styleId Absatz-Standardschriftart {
    Tag_w:name w:val {Default Paragraph Font} {
    }
    Tag_w:uiPriority w:val 1 {
    }
    Tag_w:semiHidden {
    }
    Tag_w:unhideWhenUsed {
    }
  }
  Tag_w:style w:type table w:default 1 w:styleId NormaleTabelle {
    Tag_w:name w:val {Normal Table} {
    }
    Tag_w:uiPriority w:val 99 {
    }
    Tag_w:semiHidden {
    }
    Tag_w:unhideWhenUsed {
    }
    Tag_w:tblPr {
      Tag_w:tblInd w:w 0 w:type dxa {
      }
      Tag_w:tblCellMar {
        Tag_w:top w:w 0 w:type dxa {
        }
        Tag_w:left w:w 108 w:type dxa {
        }
        Tag_w:bottom w:w 0 w:type dxa {
        }
        Tag_w:right w:w 108 w:type dxa {
        }
      }
    }
  }
  Tag_w:style w:type numbering w:default 1 w:styleId KeineListe {
    Tag_w:name w:val {No List} {
    }
    Tag_w:uiPriority w:val 99 {
    }
    Tag_w:semiHidden {
    }
    Tag_w:unhideWhenUsed {
    }
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
