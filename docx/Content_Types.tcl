package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd -tagName Default elementNode Tag_Default
dom createNodeCmd -tagName Override elementNode Tag_Override
set doc [dom createDocument Types]
set root [$doc documentElement]
$root setAttribute xmlns http://schemas.openxmlformats.org/package/2006/content-types
$root appendFromScript {
  Tag_Default Extension rels ContentType application/vnd.openxmlformats-package.relationships+xml {
  }
  Tag_Default Extension xml ContentType application/xml {
  }
  Tag_Override PartName /word/document.xml ContentType application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml {
  }
  Tag_Override PartName /word/styles.xml ContentType application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml {
  }
  Tag_Override PartName /word/settings.xml ContentType application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml {
  }
  Tag_Override PartName /word/webSettings.xml ContentType application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml {
  }
  Tag_Override PartName /word/fontTable.xml ContentType application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml {
  }
  Tag_Override PartName /word/theme/theme1.xml ContentType application/vnd.openxmlformats-officedocument.theme+xml {
  }
  Tag_Override PartName /docProps/core.xml ContentType application/vnd.openxmlformats-package.core-properties+xml {
  }
  Tag_Override PartName /docProps/app.xml ContentType application/vnd.openxmlformats-officedocument.extended-properties+xml {
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
