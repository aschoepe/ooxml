package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName w:allowPNG elementNode Tag_w:allowPNG
dom createNodeCmd -tagName w:optimizeForBrowser elementNode Tag_w:optimizeForBrowser
set doc [dom createDocument w:webSettings]
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
  Tag_w:optimizeForBrowser {
  }
  Tag_w:allowPNG {
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
