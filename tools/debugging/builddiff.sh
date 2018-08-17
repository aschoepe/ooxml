#!/bin/sh

mkdir -p /Users/alex/src/ooxml/t
rm -fr /Users/alex/src/ooxml/t/*
cd /Users/alex/src/ooxml/t

unzip -q /Users/alex/src/ooxml/examples/ok/export${1}.xlsx
xmllint --format [Content_Types].xml > 0Content_Types.xml
xmllint --format xl/workbook.xml > 0workbook.xml
xmllint --format xl/_rels/workbook.xml.rels > 0workbook.xml.rels
xmllint --format docProps/core.xml > 0core.xml
xmllint --format xl/worksheets/sheet1.xml > 0sheet1.xml
xmllint --format xl/theme/theme1.xml > 0theme1.xml
xmllint --format _rels/.rels > 0rels
xmllint --format docProps/app.xml > 0app.xml
xmllint --format xl/sharedStrings.xml > 0sharedStrings.xml
rm -fr [Content_Types].xml _rels docProps xl

unzip -q /Users/alex/src/ooxml/examples/export${1}.xlsx
xmllint --format [Content_Types].xml > 1Content_Types.xml
xmllint --format xl/workbook.xml > 1workbook.xml
xmllint --format xl/_rels/workbook.xml.rels > 1workbook.xml.rels
xmllint --format docProps/core.xml > 1core.xml
xmllint --format xl/worksheets/sheet1.xml > 1sheet1.xml
xmllint --format xl/theme/theme1.xml > 1theme1.xml
xmllint --format _rels/.rels > 1rels
xmllint --format docProps/app.xml > 1app.xml
xmllint --format xl/sharedStrings.xml > 1sharedStrings.xml
rm -fr [Content_Types].xml _rels docProps xl

> diff.txt
for i in Content_Types.xml workbook.xml workbook.xml.rels core.xml sheet1.xml theme1.xml rels app.xml sharedStrings.xml
do
  echo === $i === >> diff.txt
  diff 0$i 1$i >> diff.txt
done
