% ooxml(n) | ooxml user documentation
# NAME

::ooxml::docx - Create ECMA-376 Office Open XML writer documents (the .docx format)

# SYNOPSIS

    package require ooxml
    ::ooxml::docx new
    
# DESCRIPTION

The command *::ooxml::docx* creates an ECMA-376 Office Open XML
WordprocessingML object command and returns that. The object has
methods to add and style content. The more vulgo named ".docx" files
created by the object can be read by word processing programs as
libreOffice or Word.

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
the last one wins. An unknown key is reported as error. In almost any
case the value given to the option or key will be type checked.

The created docx object commands currently support the following
methods:

**append** *text* *?-option value ...?*

: Appends *text* to the last paragraph of the content of the document.
  The option/value pairs control locally the appearance of the *text*
  within the paragraph. See [CHARACTER OPTIONS](#character) for the
  valid options and the type of their argument.

**comment** *?-option value ...?* *creatingScript*

The allowed options are

-author string

-date <xsd datetime>

-initals string

The content of the comment is determined by the evaluated script given
as last argument.

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

**field** *fieldname*

Appends the given field type to the last paragraph. Recognized
fieldnames are

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

Set the respectively header or footer. The expeced id has to be return
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


**sectionend**

Ends the last section page settings.

**sectionstart** *?-option value ...?*

Appends a new section to the document with new page settings. The
allowed options are the same as for the method pagesetup, see there.

**settings** *?-option value ...?*

**simplecomment** *?-option value ...?*

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

**tablecell** *?-option value ...?*

**tablerow** *?-option value ...?*

**textbox** *?-option value ...?*

**url** *text* *url* *?-option value ...?*

**write** *?filename?*

: Writes the object as WordprocessingML docx file to *filename*. If
  the argument is ommited then document.docx will be used.

**writepart** *part* *filename*

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
