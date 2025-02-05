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

    : **paragraphdefault**
    
    : **characterdefault**
    
    : **paragraph**

    : **character**
    
    : **ids** *styletype*
    
    : **delete** *styletype* *styleid*

**simpletable** *args*

**write** *?filename?*

:   Writes the document as ".docx" to the file *filename*. If
    *filename* does not have the suffix ".docx" it will be appended to
    the name. 

# DEPENDENCIES

    Tcl >= 8.6.7
    tclvfs::zip >= 1.4.2 (or Tcl 9)
    tdom >= 0.9.6
