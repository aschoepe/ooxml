package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName w:charset elementNode Tag_w:charset
dom createNodeCmd -tagName w:family elementNode Tag_w:family
dom createNodeCmd -tagName w:font elementNode Tag_w:font
dom createNodeCmd -tagName w:panose1 elementNode Tag_w:panose1
dom createNodeCmd -tagName w:pitch elementNode Tag_w:pitch
dom createNodeCmd -tagName w:sig elementNode Tag_w:sig
set doc [dom createDocument w:fonts]
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
  Tag_w:font w:name Calibri {
    Tag_w:panose1 w:val 020F0502020204030204 {
    }
    Tag_w:charset w:val 00 {
    }
    Tag_w:family w:val swiss {
    }
    Tag_w:pitch w:val variable {
    }
    Tag_w:sig w:usb0 E0002AFF w:usb1 C000247B w:usb2 00000009 w:usb3 00000000 w:csb0 000001FF w:csb1 00000000 {
    }
  }
  Tag_w:font w:name {Times New Roman} {
    Tag_w:panose1 w:val 02020603050405020304 {
    }
    Tag_w:charset w:val 00 {
    }
    Tag_w:family w:val roman {
    }
    Tag_w:pitch w:val variable {
    }
    Tag_w:sig w:usb0 E0002AFF w:usb1 C0007841 w:usb2 00000009 w:usb3 00000000 w:csb0 000001FF w:csb1 00000000 {
    }
  }
  Tag_w:font w:name {Calibri Light} {
    Tag_w:panose1 w:val 020F0302020204030204 {
    }
    Tag_w:charset w:val 00 {
    }
    Tag_w:family w:val swiss {
    }
    Tag_w:pitch w:val variable {
    }
    Tag_w:sig w:usb0 E0002AFF w:usb1 C000247B w:usb2 00000009 w:usb3 00000000 w:csb0 000001FF w:csb1 00000000 {
    }
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
