package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName w:body elementNode Tag_w:body
dom createNodeCmd -tagName w:bookmarkEnd elementNode Tag_w:bookmarkEnd
dom createNodeCmd -tagName w:bookmarkStart elementNode Tag_w:bookmarkStart
dom createNodeCmd -tagName w:cols elementNode Tag_w:cols
dom createNodeCmd -tagName w:docGrid elementNode Tag_w:docGrid
dom createNodeCmd -tagName w:p elementNode Tag_w:p
dom createNodeCmd -tagName w:pgMar elementNode Tag_w:pgMar
dom createNodeCmd -tagName w:pgSz elementNode Tag_w:pgSz
dom createNodeCmd -tagName w:sectPr elementNode Tag_w:sectPr
set doc [dom createDocument w:document]
set root [$doc documentElement]
$root setAttribute xmlns:wpc http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas
$root setAttribute xmlns:cx http://schemas.microsoft.com/office/drawing/2014/chartex
$root setAttribute xmlns:cx1 http://schemas.microsoft.com/office/drawing/2015/9/8/chartex
$root setAttribute xmlns:cx2 http://schemas.microsoft.com/office/drawing/2015/10/21/chartex
$root setAttribute xmlns:cx3 http://schemas.microsoft.com/office/drawing/2016/5/9/chartex
$root setAttribute xmlns:cx4 http://schemas.microsoft.com/office/drawing/2016/5/10/chartex
$root setAttribute xmlns:cx5 http://schemas.microsoft.com/office/drawing/2016/5/11/chartex
$root setAttribute xmlns:cx6 http://schemas.microsoft.com/office/drawing/2016/5/12/chartex
$root setAttribute xmlns:cx7 http://schemas.microsoft.com/office/drawing/2016/5/13/chartex
$root setAttribute xmlns:cx8 http://schemas.microsoft.com/office/drawing/2016/5/14/chartex
$root setAttribute xmlns:mc http://schemas.openxmlformats.org/markup-compatibility/2006
$root setAttribute xmlns:aink http://schemas.microsoft.com/office/drawing/2016/ink
$root setAttribute xmlns:am3d http://schemas.microsoft.com/office/drawing/2017/model3d
$root setAttribute xmlns:o urn:schemas-microsoft-com:office:office
$root setAttribute xmlns:r http://schemas.openxmlformats.org/officeDocument/2006/relationships
$root setAttribute xmlns:m http://schemas.openxmlformats.org/officeDocument/2006/math
$root setAttribute xmlns:v urn:schemas-microsoft-com:vml
$root setAttribute xmlns:wp14 http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
$root setAttribute xmlns:wp http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
$root setAttribute xmlns:w10 urn:schemas-microsoft-com:office:word
$root setAttribute xmlns:w http://schemas.openxmlformats.org/wordprocessingml/2006/main
$root setAttribute xmlns:w14 http://schemas.microsoft.com/office/word/2010/wordml
$root setAttribute xmlns:w15 http://schemas.microsoft.com/office/word/2012/wordml
$root setAttribute xmlns:w16cid http://schemas.microsoft.com/office/word/2016/wordml/cid
$root setAttribute xmlns:w16se http://schemas.microsoft.com/office/word/2015/wordml/symex
$root setAttribute xmlns:wpg http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
$root setAttribute xmlns:wpi http://schemas.microsoft.com/office/word/2010/wordprocessingInk
$root setAttribute xmlns:wne http://schemas.microsoft.com/office/word/2006/wordml
$root setAttribute xmlns:wps http://schemas.microsoft.com/office/word/2010/wordprocessingShape
$root setAttribute mc:Ignorable {w14 w15 w16se w16cid wp14}
$root appendFromScript {
  Tag_w:body {
    Tag_w:p w:rsidR 001B1FDE w:rsidRDefault 001B1FDE {
      Tag_w:bookmarkStart w:id 0 w:name _GoBack {
      }
      Tag_w:bookmarkEnd w:id 0 {
      }
    }
    Tag_w:sectPr w:rsidR 001B1FDE w:rsidSect 00CB4A40 {
      Tag_w:pgSz w:w 11900 w:h 16840 {
      }
      Tag_w:pgMar w:top 1417 w:right 1417 w:bottom 1134 w:left 1417 w:header 708 w:footer 708 w:gutter 0 {
      }
      Tag_w:cols w:space 708 {
      }
      Tag_w:docGrid w:linePitch 360 {
      }
    }
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
