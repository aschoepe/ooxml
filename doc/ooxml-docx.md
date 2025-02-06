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

**readpart** *part* *filename*

**writepart** *part* *filename*

**paragraph** *text* *?-option value ...?*

: Appends *text* as paragraph to the content of the document. 

**append** *text* *?-option value ...?*

: Appends *text* to the last paragraph of the content of the document.

**style** *subcommand* *args*

: The supported subcommands are:

    : **paragraphdefault** *?-option value ...?*
    
    : **characterdefault** *?-option value ...?*
    
    : **paragraph** *styleid* *?-option value ...?*

    : **character** *styleid* *?-option value ...?*
    
    : **ids** *styletype*
    
    : **delete** *styletype* *styleid*

**simpletable** *args*

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

**-spacing** *{kind measue ?kind measue? ..}*

: The value to the option must be a Tcl list of keyword measurement
  pairs and defines the space before, after and inbetween the lines of
  the paragraph. The space values are give as [measurement](#measurement)
  
  
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

# OPTION VALUE TYPES

Several option values share the same value types. This types are:

## MEASUREMENT

A measurement is given either as interger value without unit or as
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
