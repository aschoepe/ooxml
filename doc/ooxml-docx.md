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
default setting.

The created docx object commands currently support the
following methods:

**append** *text* *?-option value ...?*

: Appends *text* to the last paragraph of the content of the document.
  The option/value pairs control locally the appearance of the *text*
  within the paragraph. See [CHARACTER OPTIONS](#character) for the
  valid options and the type of their argument.

**pagesetup** *?-option value ...?*

The allowed options are:

-margins keyValueList

-paperSource keyValueList

-sizeAndOrientaion keyValueList

**paragraph** *text* *?-style styleid -option value ...?*

: Appends *text* as paragraph to the content of the document. If the
  *-style* option is given, the referenced style will be used. The
  other options may locally overwrite a style stetting or add more
  properties. See [PARAGRAPH OPTIONS](#paragraph) and [CHARACTER
  OPTIONS](#character) for the valid options and the type of their
  argument.

**readpart** *part* *filename*

**write** *?filename?*

**writepart** *part* *filename*

**sectionstart** *?-option value ...?*

The allowed *-option value* pairs are the same as for the *pagesetup*
method, see there.

**sectionend**

**simpletable** *args*

**style** *subcommand* *args*

: The supported subcommands are:

    : **paragraphdefault** *?-option value ...?*
    
    : **characterdefault** *?-option value ...?*
    
    : **paragraph** *styleid* *?-option value ...?*

    : **character** *styleid* *?-option value ...?*
    
    : **ids** *styletype*
    
    : **delete** *styletype* *styleid*

**write** *?filename?*

Writes the document as ".docx" to the file *filename*. If *filename*
does not have the suffix ".docx" it will be appended to the name.

# CHARACTER OPTIONS

**-bold** *onOffValue*

**-color** *auto|RRGGBB hex value*

**-dstrike** *onOffValue*

**-font** *font familiy name*

**-fontsize** *measure*

**-italic** *onOffValue*

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
