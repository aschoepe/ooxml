package require tdom 0.9.0-
set encoding utf-8
dom createNodeCmd textNode Text
dom createNodeCmd -tagName cp:keywords elementNode Tag_cp:keywords
dom createNodeCmd -tagName cp:lastModifiedBy elementNode Tag_cp:lastModifiedBy
dom createNodeCmd -tagName cp:revision elementNode Tag_cp:revision
dom createNodeCmd -tagName dc:creator elementNode Tag_dc:creator
dom createNodeCmd -tagName dc:description elementNode Tag_dc:description
dom createNodeCmd -tagName dc:subject elementNode Tag_dc:subject
dom createNodeCmd -tagName dc:title elementNode Tag_dc:title
dom createNodeCmd -tagName dcterms:created elementNode Tag_dcterms:created
dom createNodeCmd -tagName dcterms:modified elementNode Tag_dcterms:modified
set doc [dom createDocument cp:coreProperties]
set root [$doc documentElement]
$root setAttribute xmlns:cp http://schemas.openxmlformats.org/package/2006/metadata/core-properties
$root setAttribute xmlns:dc http://purl.org/dc/elements/1.1/
$root setAttribute xmlns:dcterms http://purl.org/dc/terms/
$root setAttribute xmlns:dcmitype http://purl.org/dc/dcmitype/
$root setAttribute xmlns:xsi http://www.w3.org/2001/XMLSchema-instance
$root appendFromScript {
  Tag_dc:title {
  }
  Tag_dc:subject {
  }
  Tag_dc:creator {
    Text Microsoft Office-Benutzer
  }
  Tag_cp:keywords {
  }
  Tag_dc:description {
  }
  Tag_cp:lastModifiedBy {
    Text Microsoft Office-Benutzer
  }
  Tag_cp:revision {
    Text 1
  }
  Tag_dcterms:created xsi:type dcterms:W3CDTF {
    Text 2020-01-07T22:45:00Z
  }
  Tag_dcterms:modified xsi:type dcterms:W3CDTF {
    Text 2020-01-07T22:45:00Z
  }
}
fconfigure stdout -encoding $encoding
puts [$root asXML -indent 2 -xmlDeclaration 1 -encString [string toupper $encoding]]
