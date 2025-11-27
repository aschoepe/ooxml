% ooxml(n) | ooxml user documentation
# NAME

::ooxml::docx - Create ECMA-376 Office Open XML writer documents (the .docx format)

# SYNOPSIS

    package require ooxml
    ::ooxml::docx new *?-option value ...?* ?docx-file?
    
# DESCRIPTION

The command *::ooxml::docx* creates a docx object. Without the
optional file argument it represents an almost minimal empty document.
If the last argument is an ECMA-376 Office Open XML WordprocessingML
document (vulgo a .docx file) the inital document is a copy of this
.docx. The object has methods to append and style content as text,
tables, pictures or textboxes. The .docx file created by the object
can be read by word processing programs as libreOffice or Word.

A simple "Hello, World" example is:

    set docx [::ooxml::docx new]
    $docx paragraph "Hello, World!"
    $docx write hello.docx
    $doc destroy
    
The style and appearance of the content is determined in order by
local setting, by referencing to a defined style or finaly by the
default settings.

The methods of a docx object typically expect, if ever, a few
arguments required in order and after that option value pairs. The
options may appear in any order. An unknown option is reported as
error. If an option is given more than one time then the last one
wins. 

The value of an option is - depending on the option - either a
single value or a key value list. In the second case the key value
pairs may be in any order. If a key is given more than one time then
the last one wins. An unknown key is reported as error. In almost all
cases the value given to the option or key will be type checked.

Content to the document is added top down.

The created docx object commands currently support the following
methods:

**append** *text* *?-option value ...?*

: Appends *text* to the last paragraph of the content of the document.
  The option/value pairs control the appearance of the *text* within
  the paragraph. See [CHARACTER OPTIONS](#character) for the valid
  options and the type of their value.

**comment** *?-option value ...?* *creatingScript*

This method creates a comment to the current point in the doucment
with comment content created by the *creatingScript*. Although the
standard allows much more rich content most word processor support
only basic text formating for comments.

The allowed options are

-author string

-date <xsd datetime>

-initals string

**commentrangeend** *id* *?-option value ...?* *creatingScript*

This method marks the end of a range of content in the document and
creates a comment to this content range with comment content created
by the *creatingScript*. The *id* argument must be the id returned by
the corresponding *commentrangestart* call. For the allowed options
see the method *comment*.

**commentrangestart** ?returnvar?

This method marks the start of a range of content in the document
for which a comment should be added. The method returns the id of the
range and stores the id additionally in the variable with the name
*resultvar*, if the optional argument is given.

**configure** *?-option value ...?*

Set certain document properties. The recogniced options are:

-category

-contentStatus

-erms:created

-creator

-description

-identifier

-keywords

-language

-lastModifiedBy

-lastPrinted

-erms:modified

-revision

-subject

-title

-version

**field** *field-type* ?*switches*? *?-option value ...?*

Appends the given field type to the last paragraph. Recognized
field types are

            AUTHOR
            CREATEDATE
            DATE
            FILESIZE
            PAGE
            NUMPAGES
            SAVEDATE
            SECTION
            TIME
            TITLE
            USERNAME

If the optional argument *switches* is given its value will be
appended as the field switches to the field-type name. The following
optional *-option value* pairs allows formating of the (whole)
inserted field value. See [CHARACTER OPTIONS](#character) for the valid
  options and the type of their value.

**footer** *creatingScript* ?returnvar?
**header** *creatingScript* ?returnvar?

Creates a new footer or header by evaluating the creating script.
The id of the creaded footer or header is returnd and, if given,
stored in the returnvar.

**image** *file* *anchor|inline* *?-option value ...?*

An *inline* image is part of the text, like a character (and therefore
may change the line hight). An *anchor* image is also anchored at an
exact place within the text (which determines the page at which the
image is shown) but can be freely placed at the page.

Both types of images share the following options:


**import** *part* *docx*

: Imports the given *part* from the file *docx* into the docx object,
  replacing what the object had for that part of the docx zip archive.
  in case.
  
  The shortcut "styles" may be used for word/styles.xml and
  "numbering" for word/numbering.xml.

**jumpto** *text* *markid* *?-option value ...?*

Appends the text *text* to the last paragraph and if the text is
clicked as link (typically Ctrl-Button-1) the text processor jumps to
the mark with the id *markid* inside the document. 

  The option/value pairs control locally the appearance of the *text*
  within the paragraph. See [CHARACTER OPTIONS](#character) for the
  valid options and the type of their argument.

**mark** *markid*

Sets the mark *markid* at the end of the last paragraph.

**numberring** *subcmd* *args*

The allowed sub commands are:

*abstractNum*

*abstractNumIds*

*delete*

**pagebreak**

Appends a page break to the document.

**pagesetup** *?-option value ...?*

The allowed options are:

-margins keyValueList

-paperSource keyValueList

-sizeAndOrientaion keyValueList

-defaultHeader id
-defaultFooter id
-firstHeader id
-firstFooter id
-evenHeader id
-evenFooter id

Set the respectively header or footer. The expeced id has to be returned
from a header oder footer method call.

-border

-pageNumbering

**paragraph** *text* *?-pstyle style? ?-cstyle style? ?-option value ...?*

: Appends *text* as paragraph to the content of the document. If the
  *-pstyle* option is given, the referenced paragraph style will be
  used. If the *-cstyle* option is given the referenced character
  style will be used. The values are first searched belong the
  respective style names and if not found searched belong the
  respective styles. The other options may locally overwrite a style
  stetting or add more properties. See [PARAGRAPH OPTIONS](#paragraph)
  and [CHARACTER OPTIONS](#character) for the valid options and the
  type of their argument.

**readpart** *part* *filename*

**replace** *from* *to* *?part_list?*

Replaces every *from* substring in any text with *to*, Only uniformly
styled text is replaced (if *from* spans over more than one text part
it will not be recogniced). Without the optional *part list* argument
all parts of the document (the main documentbody, footnotes, comments
pp.) will be processed. If the argument is given then only the named
parts will be processed. The elements of the part list may be a Tcl
glob expression. Part names given by the *part list* argument which
are currently not exists in the object will be silently ignored.

**sectionend**

Ends the last section page settings.

**sectionstart** *?-option value ...?*

Appends a new section to the document with new page settings. The
allowed options are the same as for the method pagesetup, see there.

**settings** *?-option value ...?*

**selectNodes** *xpath* ?*part*? *?selectNodes options?*

Returns the result of the *xpath* expression. If no other argument is
given the context node will be the document element node of the
word/document.xml part of the docx. If the *part* argument is given the
context node of the expression will be the document element of the
docx part identified by this argument. If *part* is a relative path
inside the docx zip (example: _rels/.rels) then this is used. All
parts of the docx inside the word directory may specified without the
directory part and the .xml suffix. The optional *selectNodes options*
are routed throw.

All documents accessible by this have the following prefix namespace
mappings predefined:

Prefix   Namespace
------   ---------
 a       http://schemas.openxmlformats.org/drawingml/2006/main|
 ct      http://schemas.openxmlformats.org/package/2006/content-types|
 m       http://schemas.openxmlformats.org/officeDocument/2006/math|
 mc      http://schemas.openxmlformats.org/markup-compatibility/2006|
 o       urn:schemas-microsoft-com:office:office|
 pic     http://schemas.openxmlformats.org/drawingml/2006/picture|
 r       http://schemas.openxmlformats.org/officeDocument/2006/relationships|
 rel     http://schemas.openxmlformats.org/package/2006/relationships|
 sl      http://schemas.openxmlformats.org/schemaLibrary/2006/main|
 v       urn:schemas-microsoft-com:vml|
 w       http://schemas.openxmlformats.org/wordprocessingml/2006/main|
 w10     urn:schemas-microsoft-com:office:word|
 w14     http://schemas.microsoft.com/office/word/2010/wordml|
 wp      http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing|
 wp14    http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing|
 wpg     http://schemas.microsoft.com/office/word/2010/wordprocessingGroup|
 wps     http://schemas.microsoft.com/office/word/2010/wordprocessingShape|


**simplecomment** *comment* *?-option value ...?*

This method creates a comment to the current point in the document.
The text content of the comment will be what is given by the argument
*comment*. For the allowed options see the method *comment*.
Additionally to this options all [PARAGRAPH OPTIONS](#paragraph) and
[CHARACTER OPTIONS](#character) are allowed. If given they are
applied to format the comment text.

**simplecommentrangeend** *id* *comment* *?-option value ...?*

This method marks the end of a range of content in the document and
creates a comment to this content range with comment content given by
the *comment* argument. The *id* argument must be the id returned by
the corresponding *commentrangestart* call. For the allowed options
see the method *comment*. Additionally to this options all
[PARAGRAPH OPTIONS](#paragraph) and [CHARACTER OPTIONS](#character) are
allowed. If given they are applied to format the comment text.


**simpletable** *args*

**style** *subcommand* *args*

: The supported subcommands are:

    : **paragraphdefault** *?-option value ...?*
    
    : **characterdefault** *?-option value ...?*
    
    : **paragraph** *styleid* *?-option value ...?*

    : **character** *styleid* *?-option value ...?*
    
    : **ids** *styletype*
    
    : **delete** *styletype* *styleid*

**table** *?-option value ...?* *creatingScript*

Creates a table by defining every row and cell individually by the
*creatingScript*. The recogniced options are

-columnwidths <list of column widths>

-style id

-width keyvalue 

-align (center|end|left|right|start)

-layout

-look

-tableBorders

-cellMarginTop
-cellMarginStart
-cellMarginLeftl
-cellMarginBottom
-cellMarginEnd
-cellMarginRight

**tablecell** *?-option value ...?* *creatingScript*

**tablerow** *?-option value ...?* *creatingScript*

**textbox** *?-option value ...?*

**url** *text* *url* *?-option value ...?*

**write** *?filename?*

: Writes the object as WordprocessingML docx file to *filename*. If
  the argument is ommited then document.docx will be used.

**writepart** *part* *filename*

**xpath** *xpath* ?*part*? *?selectNodes options?*

An alias for the method *selectNodes*. See there for the meaning of
the arguments.

# CHARACTER OPTIONS


**-bold** *onOffValue*

**-color** *auto|RRGGBB hex value*

**-cstyle** *id*

**-dstrike** *onOffValue*

**-emboss** *onOffValue*

**-font** *font familiy name*

**-fontsize** *measure*

If the value has no unit the integer means twentieths of a point.

**-highlight** (black|blue|cyan|green|magenta|red|yellow|white|darkBlue|darkCyan|darkGreen|darkMagenta|darkRed|darkYellow|darkGray|lightGray|none)

**-italic** *onOffValue*

**-noProof** *onOffValue*

**-rtl** *onOffValue*

**-strict** *onOffValue*

**-underline** *kind*

The allowd *kind* values are:

    single
    words
    double
    thick
    dotted
    dottedHeavy
    dash
    dashedHeavy
    dashLong
    dashLongHeavy
    dotDash
    dashDotHeavy
    dotDotDash
    dashDotDotHeavy
    wave
    wavyHeavy
    wavyDouble
    none

**-verticalAlign** *(baseline|superscript|subscript)


# PARAGRAPH OPTIONS


**-align** *kind*

The allowd *kind* values are:

    start
    center
    end
    both
    mediumKashida
    distribute
    numTab
    highKashida
    lowKashida
    thaiDistribute
    left
    right

**-indentation** *kind*

The allowd *kind* values are:

    end
    endChars
    firstLine
    firstLineChars
    hanging
    hangingChars
    start
    startChars

**-spacing** *{kind measue ?kind measue? ..}*

The value to the option must be a Tcl list of keyword measurement
pairs and defines the space before, after and inbetween the lines of
the paragraph. The space values are give as
[measurement](#measurement)
  

# OPTION VALUE TYPES

Several option values share the same value types. This types are:

## MEASUREMENT

A measurement is given either as integer value without unit or as
integer immediately followed by a unit. If the value is just a number
then what distance is specified by the number depends on the option
and is documented with the option. The allowed units are:

    mm
         Millimeter
    cm
         Centimeter
    in
         Inch (1in = 2.54cm)
    pt
         Point (1pt = 1/72in)
    pc
         Pica (1pc = 12pt)
    pi
         The same as pc: Pica (1pi = 12pt)

# DEPENDENCIES

    Tcl >= 8.6.7
    tclvfs::zip >= 1.4.2 (or Tcl 9)
    tdom >= 0.9.6
