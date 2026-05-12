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
    $docx destroy
    
The style and appearance of the content is determined in order by
local setting, by referencing to a defined style or finally by the
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

**br** *?n?*

: Inserts a line break into the current paragraph. The optional
  argument *n* (default 1) specifies the number of breaks to insert.

**comment** *?-option value ...?* *creatingScript*

This method creates a comment to the current point in the document
with comment content created by the *creatingScript*. Although the
standard allows much more rich content most word processor support
only basic text formatting for comments.

The allowed options are

-author string

-date <xsd datetime>

-initials string

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

**destroy**

: Destroys the docx object and releases all associated DOM documents
  and resources. Inherited from TclOO.

**configure** *?-option value ...?*

Set certain document properties. The recognized options are:

-category

-contentStatus

-created

-creator

-description

-identifier

-ignorable *prefixList*: Set the MC Ignorable prefixes (e.g. "w14 wp14").

-keywords

-language

-lastModifiedBy

-lastPrinted

-modified

-revision

-subject

-title

-version

**endnote** *?-option value ...?* *creatingScript*

: Creates an endnote with content produced by *creatingScript*. A
  reference mark is inserted at the current position in the document.
  The allowed options are:

  -refstyle *styleRef*: Character style applied to the reference mark. The value may be either the internal style ID (*w:styleId*) or the display name (*w:name*).

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
optional *-option value* pairs allows formatting of the (whole)
inserted field value. See [CHARACTER OPTIONS](#character) for the valid
  options and the type of their value.

**footnote** *?-option value ...?* *creatingScript*

: Creates a footnote with content produced by *creatingScript*. A
  reference mark is inserted at the current position in the document.
  The allowed options are:

  -refstyle *styleRef*: Character style applied to the reference mark. The value may be either the internal style ID (*w:styleId*) or the display name (*w:name*).

**footer** *creatingScript* ?returnvar?
**header** *creatingScript* ?returnvar?

Creates a new footer or header by evaluating the creating script.
The id of the created footer or header is returned and, if given,
stored in the returnvar.

**image** *file* *anchor|inline* *?-option value ...?*

An *inline* image is part of the text, like a character (and therefore
may change the line height). An *anchor* image is also anchored at an
exact place within the text (which determines the page at which the
image is shown) but can be freely placed at the page.

Both types of images share the following options:

-dimension *keyValueList*: Required. Sets the image size. Keys:
width (EMU), height (EMU).

-bwMode *(auto|black|blackGray|blackWhite|clr|gray|grayWhite|
hidden|invGray|ltGray|white)*: Black-and-white mode
(default "auto").

Anchor images additionally accept:

-anchorData *keyValueList*: Anchor attributes. Keys: behindDoc,
distT, distB, distL, distR, hidden, locked, layoutInCell,
allowOverlap, relativeHeight.

-positionH *relativeFrom*: Horizontal position reference
(column, character, insideMargin, leftMargin, margin, outsideMargin,
page, rightMargin). Default "column".

-alignH *(left|right|center|inside|outside)*: Horizontal alignment
(mutually exclusive with -posOffsetH).

-posOffsetH *EMU*: Horizontal offset (mutually exclusive with
-alignH).

-positionV *relativeFrom*: Vertical position reference (insideMargin,
line, margin, outsideMargin, page, paragraph, topMargin,
bottomMargin). Default "paragraph".

-alignV *(inline|top|center|bottom|inside|outside)*: Vertical
alignment (mutually exclusive with -posOffsetV).

-posOffsetV *EMU*: Vertical offset (mutually exclusive with -alignV).

-wrapMode *(none|square|topBottom)*: Text wrapping mode (default
"none").

-wrapData *keyValueList*: Wrap attributes (only with -wrapMode
square). Keys: wrapText, distT, distB, distL, distR.


**import** *part* *docx*

: Imports the given *part* from the file *docx* into the docx object,
  replacing what the object had for that part of the docx zip archive.

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

**markstart** *name*

: Starts a named bookmark span at the current position. The bookmark
  name must be unique within the document. Use **markend** to close
  the span.

**markend** *name*

: Ends the bookmark span previously opened by **markstart** with the
  same *name*.

**math** *?-option value ...?* *creatingScript*

: Creates a math zone. The *creatingScript* creates the math content
  using the math methods described below. By default the math is
  inline within the current paragraph. Options:

  -display *onOffValue*: If on, the math is rendered as a display
  equation on its own paragraph.

  -jc *(left|center|right|centerGroup)*: Justification of a display
  equation. Ignored when -display is off.

  Inside the *creatingScript*, the following math methods are
  available. All math methods must be called from within the
  *creatingScript* of a **math** call; calling them outside a math
  zone raises a Tcl error. Invalid option values also raise a Tcl
  error.

**mrun** *?-option value ...?* *text*

: Creates a math run (text within a math zone). Options:

  -lit *onOffValue*: Literal (non-italic) text.

  -sty *(p|b|i|bi)*: Math style (plain, bold, italic, bold-italic).

  -scr *(roman|script|fraktur|double-struck|sans-serif|monospace)*:
  Math font script.

  -nor *onOffValue*: Normal text (non-math font). Mutually exclusive
  with -sty/-scr.

  -aln *onOffValue*: Alignment point.

  -brk *onOffValue*: Line break. -brkAt *integer*: alignment position
  after break.

**mfrac** *?-option value ...?* *numeratorScript* *denominatorScript*

: Creates a fraction. The two scripts create the numerator and
  denominator content. Options:

  -type *(bar|lin|noBar|skw)*: Fraction type. Defaults to bar (stacked
  with horizontal rule). *lin* gives a linear (inline slash) form,
  *noBar* stacked without bar, *skw* skewed.

**msub** *baseScript* *subscriptScript*

: Creates a subscript construct.

**msup** *baseScript* *superscriptScript*

: Creates a superscript construct.

**msubsup** *baseScript* *subscriptScript* *superscriptScript*

: Creates a simultaneous sub-superscript construct.

**msqrt** *radicandScript*

: Creates a square root.

**mroot** *degreeScript* *radicandScript*

: Creates an nth-root. The *degreeScript* produces the root index
  (e.g. 3 for cube root) and *radicandScript* the expression under
  the radical.

**mnary** *?-option value ...?* *baseScript*

: Creates an n-ary operator (summation, product, integral, etc.).
  Options:

  -char *string*: The operator character (e.g. "∑", "∏", "∫").

  -limLoc *(undOvr|subSup)*: Limit placement — under/over or
  sub/superscript position.

  -grow *onOffValue*: Whether the operator grows with its content.

  -sub *script*: Subscript (lower limit) content script.

  -sup *script*: Superscript (upper limit) content script.

  -subHide *onOffValue*: Hide the subscript.

  -supHide *onOffValue*: Hide the superscript.

**mfunc** *nameScript* *argumentScript*

: Creates a function application construct (e.g. sin, cos, lim).

**mlimlow** *baseScript* *lowerLimitScript*

: Creates a lower-limit construct (limit below the base).

**mlimupp** *baseScript* *upperLimitScript*

: Creates an upper-limit construct (limit above the base).

**maccent** *?-option value ...?* *baseScript*

: Creates an accent construct — places a diacritical mark (hat, tilde,
  dot, vector arrow, etc.) above the base expression.

  -char *string*: The accent character. Default is the combining
  circumflex accent (U+0302). Common values:
  U+0303 (tilde), U+0307 (dot), U+0308 (double dot),
  U+20D7 (combining right arrow / vector).

**mbar** *?-option value ...?* *baseScript*

: Creates an overline or underline bar above or below the base.

  -pos *top|bot*: Position of the bar (default "top" = overline).

**mdelim** *?-option value ...?* *script*

: Creates a delimiter construct — wraps content in matching delimiters
  such as parentheses, brackets, braces, absolute-value bars, norms,
  floor/ceiling symbols, etc.

  The *script* body creates the content. For a single argument, math
  methods in the script create content directly within an m:e element.
  For multiple arguments separated by the separator character, use
  **mdelimarg** within the script.

  -begChr *string*: Opening delimiter character (default "(")
  -endChr *string*: Closing delimiter character (default ")")
  -sepChr *string*: Separator character between arguments (default "|")
  -grow *onOffValue*: Delimiters grow to match content height
  -shp *centered|match*: Delimiter shape

**mdelimarg** *script*

: Creates a delimiter argument (m:e child) inside an enclosing mdelim.
  Each call to mdelimarg wraps its script content in a separate m:e
  element, enabling multi-argument delimiters.

**mgroupchr** *?-option value ...?* *baseScript*

: Creates a grouping character construct — an underbrace, overbrace,
  or similar character that spans the base expression. Commonly used
  to label parts of an equation.

  -char *string*: The grouping character (default U+23DF = ⏟ bottom
  curly bracket). Use U+23DE for overbrace (⏞).
  -pos *top|bot*: Position of the character (default "bot")
  -vertJc *top|bot*: Vertical justification of the base relative to
  the grouping character

**mmatrix** *?-option value ...?* *rowScripts*

: Creates a matrix. The *rowScripts* argument is a list; each element
  is a script that creates cells via **mmatcell** calls. Wrap in
  **mdelim** for bracket/parenthesis notation.

  -baseJc *top|center|bottom*: Vertical alignment of the matrix
  relative to the math baseline
  -mcJc *left|center|right*: Column justification
  -count *integer*: Number of adjacent columns sharing the
  justification (1–255)

**mmatcell** *script*

: Creates a matrix cell (m:e child) inside a matrix row. Each call
  wraps its script content in a separate m:e element.

**numbering** *subcmd* *args*

: Manages numbering definitions (lists, outlines).

  The supported subcommands are:

  **abstractNum** *id* *levelDefinitions*: Creates an abstract
  numbering definition with the given integer *id*. The
  *levelDefinitions* argument is a list; each element defines one
  level using option/value pairs:

      -numberFormat *format*: Numbering format (decimal, upperRoman,
      lowerRoman, upperLetter, lowerLetter, bullet, etc.)
      -levelText *string*: Level text template (e.g. "%1." or "%1.%2")
      -align *alignment*: Number alignment (start, center, end, etc.)
      -start *integer*: Starting number (default 1)

  Paragraph and character options may also be given per level to
  style the number and the paragraph.

  A corresponding w:num entry is automatically created with the
  same numId as the abstractNumId.

  Example:

      $doc numbering abstractNum 0 {
          {-numberFormat decimal -levelText "%1." -align start}
          {-numberFormat lowerLetter -levelText "%2)" -align start}
      }

  **abstractNumIds**: Returns a sorted list of all abstractNum IDs.

  **delete** *type* *id*: Deletes a numbering definition.
  *type* is either **abstractNum** or **num**; *id* is the integer ID.

**pagebreak**

Appends a page break to the document.

**pagesetup** *?-option value ...?*

The allowed options are:

-margins keyValueList

-paperSource keyValueList

-sizeAndOrientation keyValueList

-defaultHeader id
-defaultFooter id
-firstHeader id
-firstFooter id
-evenHeader id
-evenFooter id

Set the respectively header or footer. The expected id has to be returned
from a header or footer method call.

-border

-pageNumbering

-sectionType *(continuous|evenPage|nextColumn|nextPage|oddPage)*

Specifies the type of the current section break. The default is
*nextPage*. Use *continuous* to start a new section on the same page
(e.g. for switching between single and multi-column layout).

-docGrid *keyValueList*

Configures the document grid for the section, controlling line
pitch and character spacing for CJK and other layout needs.
The accepted keys are:

    type        Grid type: "default", "lines", "linesAndChars",
                or "snapToChars".
    linePitch   Distance between lines (integer, in twips).
    charSpace   Additional character pitch (integer, in 4096ths
                of a point).

-vAlign *(top|center|bottom|both)*

Vertical alignment of text on the page. The value *both* distributes
text evenly between the top and bottom margins (justified).

**paragraph** *text* *?-pstyle style? ?-cstyle style? ?-option value ...?*

: Appends *text* as paragraph to the content of the document. If the
  *-pstyle* option is given, the referenced paragraph style will be
  used. If the *-cstyle* option is given the referenced character
  style will be used. Style values may be given either as the internal
  style ID (*w:styleId*) or as the display name (*w:name*). Resolution
  prefers an exact style-ID match and falls back to the display name.
  The other options
  may locally overwrite a style setting or add more properties. See
  [PARAGRAPH OPTIONS](#paragraph) and [CHARACTER OPTIONS](#character)
  for the valid options and the type of their argument.

**readpart** *part* *filename*

: Replaces the named XML *part* in the docx object with the contents
  of *filename*. The file is parsed as XML and replaces the in-memory
  DOM for that part. Example:

      $doc readpart word/settings.xml mysettings.xml

**replace** *from* *to* *?part_list?*

Replaces every *from* substring in any text with *to*, Only uniformly
styled text is replaced (if *from* spans over more than one text part
it will not be recognized). Without the optional *part list* argument
all parts of the document (the main documentbody, footnotes, comments
pp.) will be processed. If the argument is given then only the named
parts will be processed. The elements of the part list may be a Tcl
glob expression. Part names given by the *part list* argument which
are currently not exists in the object will be silently ignored.

**sectionend**

: Ends the current section, emitting the stored section properties as
  a w:sectPr inside the last paragraph. Must be preceded by a
  **sectionstart** call; raises an error otherwise.

**sectionstart** *?-option value ...?*

: Starts a new section in the document. If a previous section was
  already started (and not ended), it is implicitly ended first.
  The allowed options are the same as for **pagesetup**: -margins,
  -sizeAndOrientation, -sectionType, -docGrid, -vAlign,
  -border, -pageNumbering, header/footer references, etc.

**settings** *?-option value ...?*

: Modifies the document settings (word/settings.xml). Settings may be
  called multiple times; each call updates or adds the specified
  settings without affecting those set by earlier calls.

  Special option:

  -reset *onOffValue*: If on, resets all settings to defaults before
  applying the remaining options.

  Commonly used settings:

  -defaultTabStop *integer*: Default tab stop distance in twips.
  -autoHyphenation *onOffValue*: Enable automatic hyphenation.
  -evenAndOddHeaders *onOffValue*: Use different headers/footers for
  even and odd pages.
  -hideSpellingErrors *onOffValue*: Hide red squiggly underlines.
  -hideGrammaticalErrors *onOffValue*: Hide grammar error marks.
  -trackRevisions *onOffValue*: Enable revision tracking.
  -mirrorMargins *onOffValue*: Mirror left/right margins for
  facing pages.
  -embedTrueTypeFonts *onOffValue*: Embed TrueType fonts.
  -documentProtection *keyValueList*: Protect the document.
  Keys: edit (readOnly|comments|trackedChanges|forms|none),
  enforcement (on|off), formatting (on|off), plus optional
  cryptographic parameters.

  Many more settings are available. On/off settings use *onOffValue*.
  Refer to the property table in the source for the complete list.

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


**simpletable** *tabledata* *?-option value ...?*

: Creates a table from a nested list. Each element of *tabledata* is a
  row; each element within a row is a cell's text content.

  Example: `$doc simpletable {{A1 B1 C1} {A2 B2 C2}}`

  Options:

  -columnwidths *list*: Column widths (measurements).
  -firstStyle *styleName*: Character style for the first row.
  -lastStyle *styleName*: Character style for the last row.

  Table-level options (-style, -width, -align, -layout, -look,
  -tableBorders, cell margins) are also accepted — see **table**.

**style** *subcommand* *args*

: The supported subcommands are:

    : **paragraphdefault** *?-option value ...?*
    
    : **characterdefault** *?-option value ...?*
    
    : **paragraph** *name* *?-option value ...?*

    : **character** *name* *?-option value ...?*

    : **table** *name* *?-option value ...?*

    : The *name* argument is the display name of the style. The
      internal style ID (w:styleId) is derived from the name by
      stripping non-alphanumeric characters. Use the **-styleid**
      option to set the internal ID explicitly if the derived ID
      would be empty or ambiguous. Both display name and derived ID
      must be unique within the style type.

      Additional style creation options:

      -basedon *styleRef*: Parent style to inherit from. Accepts
      either an internal style ID or a display name (resolved to a
      style ID internally). The parent must be the same type as the
      style being created.

      -conditional *keyValueList*: (Table styles only.) Conditional
      formatting for table regions such as firstRow, lastRow,
      firstColumn, lastColumn, band1Vert, band1Horz, etc.

      All methods and options that reference a style (such as
      **-pstyle**, **-cstyle**, **-style**, **-basedon**,
      **-refstyle**) accept either the internal style ID (preferred,
      tried first) or the display name (fallback).

    : **ids** *styletype*
    
    : **names** *styletype*

    : **delete** *styletype* *name*

### LOW-LEVEL TAG CONSTRUCTORS

The package exports a broad family of `Tag_*` constructors as a low-level
OOXML-building API. Not every exported constructor is used by the higher-level
methods in this file, but they are intentionally kept available for callers
that need to emit schema elements directly.

**tab** *?n?*

: Inserts a tab character into the current paragraph. The optional
  argument *n* (default 1) specifies the number of tabs to insert.

**table** *?-option value ...?* *creatingScript*

Creates a table by defining every row and cell individually by the
*creatingScript*. The recognized options are:

-columnwidths *list*: List of column widths (measurements).

-style *styleRef*: Table style reference (style ID or display name).

-width *keyValueList*: Table width. Keys: type (auto|dxa|pct|nil),
value (measurement or percentage).

-align *(center|end|left|right|start)*: Table alignment.

-layout *keyValueList*: Table layout. Keys: type (autofit|fixed).

-look *keyValueList*: Table conditional formatting flags. Keys:
firstRow, lastRow, firstColumn, lastColumn, noHBand, noVBand
(all on/off).

-caption *string*: Table caption for accessibility.

-tableBorders: Table border options. For each border position
(top, left, bottom, right, insideH, insideV), use
`-<position>Border {type <borderType> color <hex> borderwidth <eighthPts> space <pts>}`.

-cellMarginTop, -cellMarginStart, -cellMarginLeft,
-cellMarginBottom, -cellMarginEnd, -cellMarginRight:
Default cell margins. Each takes a *keyValueList* with keys
type and value.

**tablecell** *?-option value ...?* *creatingScript*

: Creates a table cell within a **tablerow** script. The
  *creatingScript* creates the cell content (paragraphs, nested
  tables, etc.).

  Options:

  -cellWidth *keyValueList*: Cell width. Keys: type (auto|dxa|pct|nil),
  value (measurement or percentage).

  -span *integer*: Number of grid columns this cell spans.

  -hspan *(restart|continue)*: Horizontal merge state.

  -vspan *(restart|continue)*: Vertical merge state.

  -vAlign *(top|center|bottom|both)*: Vertical alignment of cell
  content.

  -textDirection *direction*: Text direction within the cell.

  -shading *keyValueList*: Cell shading pattern. Keys: type, color,
  fill (same as paragraph shading). The options -shading and
  -background are mutually exclusive.

  -background *RRGGBB*: Shorthand for a solid background fill color.
  The options -shading and -background are mutually exclusive.

  -tcFitText *onOffValue*: Shrink text to fit cell width.

  Cell border options (-topBorder, -leftBorder, -bottomBorder,
  -rightBorder, -tl2brBorder, -tr2blBorder) and cell margin
  options are also accepted.

**tablerow** *?-option value ...?* *creatingScript*

: Creates a table row within a **table** script. The *creatingScript*
  creates cells via **tablecell** calls.

  Options:

  -rowHeight *keyValueList*: Row height. Keys: value (measurement),
  hRule (auto|exact|atLeast).

  -headerrow *onOffValue*: Mark this row as a header row that repeats
  on each page.

  -cantSplit *onOffValue*: Prevent the row from splitting across
  pages.

  -align *(center|end|left|right|start)*: Row alignment (overrides
  table alignment).

  -cellSpacing *keyValueList*: Cell spacing within the row. Keys:
  type, value.

  -wBefore, -wAfter *keyValueList*: Width before/after the row
  (for indented rows). Keys: type, value.

**textbox** *?-option value ...?* *creatingScript*

: Creates an anchored text box at the current position. The
  *creatingScript* creates the text box content (paragraphs, tables,
  etc.) inside the box.

  Options:

  -name *string*: Display name for the text box shape. Auto-generated
  if omitted.

  -dimension *keyValueList*: Required. Sets the text box size. Keys:
  width (EMU), height (EMU).

  -bodyAtts *keyValueList*: Text body attributes. Keys: wrap
  (none|square), numCol (integer), and other wps:bodyPr attributes.

  Anchor positioning options (-anchorData, -positionH, -alignH,
  -posOffsetH, -positionV, -alignV, -posOffsetV, -wrapMode,
  -wrapData, -bwMode) are also accepted — see **image** anchor
  options.

**url** *text* *url* *?-option value ...?*

**write** *?filename?*

: Writes the object as WordprocessingML docx file to *filename*. If
  the argument is omitted then document.docx will be used.

**writepart** *part* *filename*

: Writes a single XML part from the docx object to *filename*.
  Raises an error if the part does not exist in the object.

**xpath** *xpath* *?part?* *?selectNodes options?*

An alias for the method *selectNodes*. See there for the meaning of
the arguments.

**xmlparts** *?pattern?*

Returns the paths of the xml parts of the docx object. If the optional
argument *pattern* is given only the xml parts of the docx docx object
which match the *pattern* using glob style matching are returned.


# CHARACTER OPTIONS


**-bold** *onOffValue*

**-color** *auto|RRGGBB hex value*

**-cstyle** *styleRef*

**-dstrike** *onOffValue*

**-emboss** *onOffValue*

**-font** *font familiy name*

**-fontsize** *measure*

If the value has no unit the integer means twentieths of a point.

**-highlight** (black|blue|cyan|green|magenta|red|yellow|white|darkBlue|darkCyan|darkGreen|darkMagenta|darkRed|darkYellow|darkGray|lightGray|none)

**-italic** *onOffValue*

**-noProof** *onOffValue*

**-rtl** *onOffValue*

**-strike** *onOffValue*

Toggles strikethrough on the text.

**-underline** *kind*

The allowed *kind* values are:

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

The allowed *kind* values are:

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

**-background** *RRGGBB hex value*

Sets a solid background fill color on the paragraph. This is a
shorthand for `-shading {type clear fill RRGGBB}`.

**-bidi** *onOffValue*

Right-to-left paragraph direction.

**-indentation** *keyValueList*

Sets the paragraph indentation. The accepted keys are:

    start           Left/start indentation (measurement).
    startChars      Left/start indentation in character units.
    end             Right/end indentation (measurement).
    endChars        Right/end indentation in character units.
    firstLine       First line additional indent (measurement).
    firstLineChars  First line indent in character units.
    hanging         Hanging indent (measurement).
    hangingChars    Hanging indent in character units.

**-keepLines** *onOffValue*

Keep all lines of the paragraph on the same page.

**-keepNext** *onOffValue*

Keep the paragraph on the same page as the next paragraph.

**-level** *integer*

Numbering level (used with **-numberingStyle**). Zero-based.

**-numberingStyle** *integer*

Numbering definition ID to apply to this paragraph.

**-pageBreakBefore** *onOffValue*

Start the paragraph on a new page.

**-shading** *keyValueList*

Sets the paragraph shading pattern. The accepted keys are:

    type    Shading pattern (required). Common values: clear, solid,
            horzStripe, vertStripe, diagStripe, diagCross, pct10,
            pct20, pct25, pct50, pct75, etc.
    color   Foreground pattern color (RRGGBB hex or "auto").
    fill    Background fill color (RRGGBB hex or "auto").

Example: `-shading {type clear fill FFFF00}` for a yellow background.

**-spacing** *keyValueList*

The value to the option must be a Tcl list of keyword measurement
pairs and defines the space before, after and between the lines of
the paragraph. The space values are given as
[measurement](#measurement)

The accepted keys are:

    after       Space after the paragraph.
    before      Space before the paragraph.
    line        Line spacing within the paragraph.
    lineRule    How "line" is interpreted: "auto" (default, value
                is in 240ths of a line), "exact" (value is in
                twips), or "atLeast" (minimum, value is in twips).

**-suppressLineNumbers** *onOffValue*

Suppress line numbers for this paragraph.

**-textframe** *keyValueList*

Configures the paragraph as a text frame. Keys: width, height,
wrap, vAnchor, hAnchor, xAlign, yAlign, dropCap, lines, vSpace,
hSpace, hrule, anchorLock.

**-widowControl** *onOffValue*

Enable widow/orphan control for this paragraph.
  

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
