package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd textNode Text
dom createNodeCmd -tagName AppVersion elementNode Tag_AppVersion
dom createNodeCmd -tagName Application elementNode Tag_Application
dom createNodeCmd -tagName Characters elementNode Tag_Characters
dom createNodeCmd -tagName CharactersWithSpaces elementNode Tag_CharactersWithSpaces
dom createNodeCmd -tagName Company elementNode Tag_Company
dom createNodeCmd -tagName DocSecurity elementNode Tag_DocSecurity
dom createNodeCmd -tagName HyperlinksChanged elementNode Tag_HyperlinksChanged
dom createNodeCmd -tagName Lines elementNode Tag_Lines
dom createNodeCmd -tagName LinksUpToDate elementNode Tag_LinksUpToDate
dom createNodeCmd -tagName Pages elementNode Tag_Pages
dom createNodeCmd -tagName Paragraphs elementNode Tag_Paragraphs
dom createNodeCmd -tagName ScaleCrop elementNode Tag_ScaleCrop
dom createNodeCmd -tagName SharedDoc elementNode Tag_SharedDoc
dom createNodeCmd -tagName Template elementNode Tag_Template
dom createNodeCmd -tagName TotalTime elementNode Tag_TotalTime
dom createNodeCmd -tagName Words elementNode Tag_Words
set doc [dom createDocument Properties]
set root [$doc documentElement]
$root setAttribute xmlns http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
$root setAttribute xmlns:vt http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes
$root appendFromScript {
  Tag_Template {
    Text Normal.dotm
  }
  Tag_TotalTime {
    Text 0
  }
  Tag_Pages {
    Text 1
  }
  Tag_Words {
    Text 0
  }
  Tag_Characters {
    Text 0
  }
  Tag_Application {
    Text Microsoft Office Word
  }
  Tag_DocSecurity {
    Text 0
  }
  Tag_Lines {
    Text 0
  }
  Tag_Paragraphs {
    Text 0
  }
  Tag_ScaleCrop {
    Text false
  }
  Tag_Company {
  }
  Tag_LinksUpToDate {
    Text false
  }
  Tag_CharactersWithSpaces {
    Text 0
  }
  Tag_SharedDoc {
    Text false
  }
  Tag_HyperlinksChanged {
    Text false
  }
  Tag_AppVersion {
    Text 16.0000
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
