package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName a:accent1 elementNode Tag_a:accent1
dom createNodeCmd -tagName a:accent2 elementNode Tag_a:accent2
dom createNodeCmd -tagName a:accent3 elementNode Tag_a:accent3
dom createNodeCmd -tagName a:accent4 elementNode Tag_a:accent4
dom createNodeCmd -tagName a:accent5 elementNode Tag_a:accent5
dom createNodeCmd -tagName a:accent6 elementNode Tag_a:accent6
dom createNodeCmd -tagName a:alpha elementNode Tag_a:alpha
dom createNodeCmd -tagName a:bgFillStyleLst elementNode Tag_a:bgFillStyleLst
dom createNodeCmd -tagName a:clrScheme elementNode Tag_a:clrScheme
dom createNodeCmd -tagName a:cs elementNode Tag_a:cs
dom createNodeCmd -tagName a:dk1 elementNode Tag_a:dk1
dom createNodeCmd -tagName a:dk2 elementNode Tag_a:dk2
dom createNodeCmd -tagName a:ea elementNode Tag_a:ea
dom createNodeCmd -tagName a:effectLst elementNode Tag_a:effectLst
dom createNodeCmd -tagName a:effectStyle elementNode Tag_a:effectStyle
dom createNodeCmd -tagName a:effectStyleLst elementNode Tag_a:effectStyleLst
dom createNodeCmd -tagName a:ext elementNode Tag_a:ext
dom createNodeCmd -tagName a:extLst elementNode Tag_a:extLst
dom createNodeCmd -tagName a:extraClrSchemeLst elementNode Tag_a:extraClrSchemeLst
dom createNodeCmd -tagName a:fillStyleLst elementNode Tag_a:fillStyleLst
dom createNodeCmd -tagName a:fmtScheme elementNode Tag_a:fmtScheme
dom createNodeCmd -tagName a:folHlink elementNode Tag_a:folHlink
dom createNodeCmd -tagName a:font elementNode Tag_a:font
dom createNodeCmd -tagName a:fontScheme elementNode Tag_a:fontScheme
dom createNodeCmd -tagName a:gradFill elementNode Tag_a:gradFill
dom createNodeCmd -tagName a:gs elementNode Tag_a:gs
dom createNodeCmd -tagName a:gsLst elementNode Tag_a:gsLst
dom createNodeCmd -tagName a:hlink elementNode Tag_a:hlink
dom createNodeCmd -tagName a:latin elementNode Tag_a:latin
dom createNodeCmd -tagName a:lin elementNode Tag_a:lin
dom createNodeCmd -tagName a:ln elementNode Tag_a:ln
dom createNodeCmd -tagName a:lnStyleLst elementNode Tag_a:lnStyleLst
dom createNodeCmd -tagName a:lt1 elementNode Tag_a:lt1
dom createNodeCmd -tagName a:lt2 elementNode Tag_a:lt2
dom createNodeCmd -tagName a:lumMod elementNode Tag_a:lumMod
dom createNodeCmd -tagName a:majorFont elementNode Tag_a:majorFont
dom createNodeCmd -tagName a:minorFont elementNode Tag_a:minorFont
dom createNodeCmd -tagName a:miter elementNode Tag_a:miter
dom createNodeCmd -tagName a:objectDefaults elementNode Tag_a:objectDefaults
dom createNodeCmd -tagName a:outerShdw elementNode Tag_a:outerShdw
dom createNodeCmd -tagName a:prstDash elementNode Tag_a:prstDash
dom createNodeCmd -tagName a:satMod elementNode Tag_a:satMod
dom createNodeCmd -tagName a:schemeClr elementNode Tag_a:schemeClr
dom createNodeCmd -tagName a:shade elementNode Tag_a:shade
dom createNodeCmd -tagName a:solidFill elementNode Tag_a:solidFill
dom createNodeCmd -tagName a:srgbClr elementNode Tag_a:srgbClr
dom createNodeCmd -tagName a:sysClr elementNode Tag_a:sysClr
dom createNodeCmd -tagName a:themeElements elementNode Tag_a:themeElements
dom createNodeCmd -tagName a:tint elementNode Tag_a:tint
dom createNodeCmd -tagName thm15:themeFamily elementNode Tag_thm15:themeFamily
set doc [dom createDocument a:theme]
set root [$doc documentElement]
$root setAttribute xmlns:a http://schemas.openxmlformats.org/drawingml/2006/main
$root setAttribute name Office-Design
$root appendFromScript {
  Tag_a:themeElements {
    Tag_a:clrScheme name Office {
      Tag_a:dk1 {
        Tag_a:sysClr val windowText lastClr 000000 {
        }
      }
      Tag_a:lt1 {
        Tag_a:sysClr val window lastClr FFFFFF {
        }
      }
      Tag_a:dk2 {
        Tag_a:srgbClr val 44546A {
        }
      }
      Tag_a:lt2 {
        Tag_a:srgbClr val E7E6E6 {
        }
      }
      Tag_a:accent1 {
        Tag_a:srgbClr val 4472C4 {
        }
      }
      Tag_a:accent2 {
        Tag_a:srgbClr val ED7D31 {
        }
      }
      Tag_a:accent3 {
        Tag_a:srgbClr val A5A5A5 {
        }
      }
      Tag_a:accent4 {
        Tag_a:srgbClr val FFC000 {
        }
      }
      Tag_a:accent5 {
        Tag_a:srgbClr val 5B9BD5 {
        }
      }
      Tag_a:accent6 {
        Tag_a:srgbClr val 70AD47 {
        }
      }
      Tag_a:hlink {
        Tag_a:srgbClr val 0563C1 {
        }
      }
      Tag_a:folHlink {
        Tag_a:srgbClr val 954F72 {
        }
      }
    }
    Tag_a:fontScheme name Office {
      Tag_a:majorFont {
        Tag_a:latin typeface {Calibri Light} panose 020F0302020204030204 {
        }
        Tag_a:ea typeface {} {
        }
        Tag_a:cs typeface {} {
        }
        Tag_a:font script Jpan typeface {Yu Gothic Light} {
        }
        Tag_a:font script Hang typeface {맑은 고딕} {
        }
        Tag_a:font script Hans typeface {DengXian Light} {
        }
        Tag_a:font script Hant typeface 新細明體 {
        }
        Tag_a:font script Arab typeface {Times New Roman} {
        }
        Tag_a:font script Hebr typeface {Times New Roman} {
        }
        Tag_a:font script Thai typeface {Angsana New} {
        }
        Tag_a:font script Ethi typeface Nyala {
        }
        Tag_a:font script Beng typeface Vrinda {
        }
        Tag_a:font script Gujr typeface Shruti {
        }
        Tag_a:font script Khmr typeface MoolBoran {
        }
        Tag_a:font script Knda typeface Tunga {
        }
        Tag_a:font script Guru typeface Raavi {
        }
        Tag_a:font script Cans typeface Euphemia {
        }
        Tag_a:font script Cher typeface {Plantagenet Cherokee} {
        }
        Tag_a:font script Yiii typeface {Microsoft Yi Baiti} {
        }
        Tag_a:font script Tibt typeface {Microsoft Himalaya} {
        }
        Tag_a:font script Thaa typeface {MV Boli} {
        }
        Tag_a:font script Deva typeface Mangal {
        }
        Tag_a:font script Telu typeface Gautami {
        }
        Tag_a:font script Taml typeface Latha {
        }
        Tag_a:font script Syrc typeface {Estrangelo Edessa} {
        }
        Tag_a:font script Orya typeface Kalinga {
        }
        Tag_a:font script Mlym typeface Kartika {
        }
        Tag_a:font script Laoo typeface DokChampa {
        }
        Tag_a:font script Sinh typeface {Iskoola Pota} {
        }
        Tag_a:font script Mong typeface {Mongolian Baiti} {
        }
        Tag_a:font script Viet typeface {Times New Roman} {
        }
        Tag_a:font script Uigh typeface {Microsoft Uighur} {
        }
        Tag_a:font script Geor typeface Sylfaen {
        }
      }
      Tag_a:minorFont {
        Tag_a:latin typeface Calibri panose 020F0502020204030204 {
        }
        Tag_a:ea typeface {} {
        }
        Tag_a:cs typeface {} {
        }
        Tag_a:font script Jpan typeface {Yu Mincho} {
        }
        Tag_a:font script Hang typeface {맑은 고딕} {
        }
        Tag_a:font script Hans typeface DengXian {
        }
        Tag_a:font script Hant typeface 新細明體 {
        }
        Tag_a:font script Arab typeface Arial {
        }
        Tag_a:font script Hebr typeface Arial {
        }
        Tag_a:font script Thai typeface {Cordia New} {
        }
        Tag_a:font script Ethi typeface Nyala {
        }
        Tag_a:font script Beng typeface Vrinda {
        }
        Tag_a:font script Gujr typeface Shruti {
        }
        Tag_a:font script Khmr typeface DaunPenh {
        }
        Tag_a:font script Knda typeface Tunga {
        }
        Tag_a:font script Guru typeface Raavi {
        }
        Tag_a:font script Cans typeface Euphemia {
        }
        Tag_a:font script Cher typeface {Plantagenet Cherokee} {
        }
        Tag_a:font script Yiii typeface {Microsoft Yi Baiti} {
        }
        Tag_a:font script Tibt typeface {Microsoft Himalaya} {
        }
        Tag_a:font script Thaa typeface {MV Boli} {
        }
        Tag_a:font script Deva typeface Mangal {
        }
        Tag_a:font script Telu typeface Gautami {
        }
        Tag_a:font script Taml typeface Latha {
        }
        Tag_a:font script Syrc typeface {Estrangelo Edessa} {
        }
        Tag_a:font script Orya typeface Kalinga {
        }
        Tag_a:font script Mlym typeface Kartika {
        }
        Tag_a:font script Laoo typeface DokChampa {
        }
        Tag_a:font script Sinh typeface {Iskoola Pota} {
        }
        Tag_a:font script Mong typeface {Mongolian Baiti} {
        }
        Tag_a:font script Viet typeface Arial {
        }
        Tag_a:font script Uigh typeface {Microsoft Uighur} {
        }
        Tag_a:font script Geor typeface Sylfaen {
        }
      }
    }
    Tag_a:fmtScheme name Office {
      Tag_a:fillStyleLst {
        Tag_a:solidFill {
          Tag_a:schemeClr val phClr {
          }
        }
        Tag_a:gradFill rotWithShape 1 {
          Tag_a:gsLst {
            Tag_a:gs pos 0 {
              Tag_a:schemeClr val phClr {
                Tag_a:lumMod val 110000 {
                }
                Tag_a:satMod val 105000 {
                }
                Tag_a:tint val 67000 {
                }
              }
            }
            Tag_a:gs pos 50000 {
              Tag_a:schemeClr val phClr {
                Tag_a:lumMod val 105000 {
                }
                Tag_a:satMod val 103000 {
                }
                Tag_a:tint val 73000 {
                }
              }
            }
            Tag_a:gs pos 100000 {
              Tag_a:schemeClr val phClr {
                Tag_a:lumMod val 105000 {
                }
                Tag_a:satMod val 109000 {
                }
                Tag_a:tint val 81000 {
                }
              }
            }
          }
          Tag_a:lin ang 5400000 scaled 0 {
          }
        }
        Tag_a:gradFill rotWithShape 1 {
          Tag_a:gsLst {
            Tag_a:gs pos 0 {
              Tag_a:schemeClr val phClr {
                Tag_a:satMod val 103000 {
                }
                Tag_a:lumMod val 102000 {
                }
                Tag_a:tint val 94000 {
                }
              }
            }
            Tag_a:gs pos 50000 {
              Tag_a:schemeClr val phClr {
                Tag_a:satMod val 110000 {
                }
                Tag_a:lumMod val 100000 {
                }
                Tag_a:shade val 100000 {
                }
              }
            }
            Tag_a:gs pos 100000 {
              Tag_a:schemeClr val phClr {
                Tag_a:lumMod val 99000 {
                }
                Tag_a:satMod val 120000 {
                }
                Tag_a:shade val 78000 {
                }
              }
            }
          }
          Tag_a:lin ang 5400000 scaled 0 {
          }
        }
      }
      Tag_a:lnStyleLst {
        Tag_a:ln w 6350 cap flat cmpd sng algn ctr {
          Tag_a:solidFill {
            Tag_a:schemeClr val phClr {
            }
          }
          Tag_a:prstDash val solid {
          }
          Tag_a:miter lim 800000 {
          }
        }
        Tag_a:ln w 12700 cap flat cmpd sng algn ctr {
          Tag_a:solidFill {
            Tag_a:schemeClr val phClr {
            }
          }
          Tag_a:prstDash val solid {
          }
          Tag_a:miter lim 800000 {
          }
        }
        Tag_a:ln w 19050 cap flat cmpd sng algn ctr {
          Tag_a:solidFill {
            Tag_a:schemeClr val phClr {
            }
          }
          Tag_a:prstDash val solid {
          }
          Tag_a:miter lim 800000 {
          }
        }
      }
      Tag_a:effectStyleLst {
        Tag_a:effectStyle {
          Tag_a:effectLst {
          }
        }
        Tag_a:effectStyle {
          Tag_a:effectLst {
          }
        }
        Tag_a:effectStyle {
          Tag_a:effectLst {
            Tag_a:outerShdw blurRad 57150 dist 19050 dir 5400000 algn ctr rotWithShape 0 {
              Tag_a:srgbClr val 000000 {
                Tag_a:alpha val 63000 {
                }
              }
            }
          }
        }
      }
      Tag_a:bgFillStyleLst {
        Tag_a:solidFill {
          Tag_a:schemeClr val phClr {
          }
        }
        Tag_a:solidFill {
          Tag_a:schemeClr val phClr {
            Tag_a:tint val 95000 {
            }
            Tag_a:satMod val 170000 {
            }
          }
        }
        Tag_a:gradFill rotWithShape 1 {
          Tag_a:gsLst {
            Tag_a:gs pos 0 {
              Tag_a:schemeClr val phClr {
                Tag_a:tint val 93000 {
                }
                Tag_a:satMod val 150000 {
                }
                Tag_a:shade val 98000 {
                }
                Tag_a:lumMod val 102000 {
                }
              }
            }
            Tag_a:gs pos 50000 {
              Tag_a:schemeClr val phClr {
                Tag_a:tint val 98000 {
                }
                Tag_a:satMod val 130000 {
                }
                Tag_a:shade val 90000 {
                }
                Tag_a:lumMod val 103000 {
                }
              }
            }
            Tag_a:gs pos 100000 {
              Tag_a:schemeClr val phClr {
                Tag_a:shade val 63000 {
                }
                Tag_a:satMod val 120000 {
                }
              }
            }
          }
          Tag_a:lin ang 5400000 scaled 0 {
          }
        }
      }
    }
  }
  Tag_a:objectDefaults {
  }
  Tag_a:extraClrSchemeLst {
  }
  Tag_a:extLst {
    Tag_a:ext uri {{05A4C25C-085E-4340-85A3-A5531E510DB2}} {
      Tag_thm15:themeFamily xmlns:thm15 http://schemas.microsoft.com/office/thememl/2012/main name {Office Theme} id {{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}} vid {{4A3C46E8-61CC-4603-A589-7422A47A8E4A}} {
      }
    }
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
