# NAME

::ooxml::xl_sheets, ::ooxml::read, ::ooxml::write - Read and create
ECMA-376 Office Open XML Spreadsheets (the .xlsx format)

# SYNOPSIS

    package require ooxml
    ::ooxml::xl_sheets file
    ::ooxml::xl_read file args
    ::ooxml::xl_write args
    
# DESCRIPTION

The commands of this package read or create ECMA-376 Office Open XML
Spreadsheets (libreoffic and Excel ".xlsx" files). The most important
three are:

    ::ooxml::xl\_read - imports a .xlsx spreadsheet files into a Tcl array
    ::ooxml::xl\_write - creates spreadsheet object command
    ::ooxml::tablelist_to_xl - export a Tcl tablelist to a .xlsx spreadsheet

To create a spreadsheet from scratch use the ::ooxml::xl\_write to
create a spreadsheet object command:

    set spreadsheet [::ooxml::xl_write new ?options?]
    
The following options are currently supported:

**-creator** *NAME*

:   *Creator* specifies the creator property of the spreadsheet.

**-created** *UTC-TIMESTAMP*

:   *Created* specifies the creation time property of the spreadsheet.

**-nodifiedby** *NAME*

:   *Modifiedby* specifies the modified by property of the spreadsheet.

**-modified** *UTC-TIMESTAMP*

:   *Modified* specifies the modified property of the spreadsheet.

**-application** *NAME*

:   *Application* specifies the application property of the spreadsheet.

The created spreadsheet object commands currently supports the
following methods:

**numberformat** *args*

:

**defaultdatestyle** *STYLEID*

:

**font** *args*

:

**fill** *args*

:

**border** *args*

:

**style** *args*

:

**worksheet** *name*

:

**column** *sheet* *args*

:

**row** *sheet* *args*

:

**cell** *sheet* *args*

:

**autofilter** *sheet* *indexFrom* *indexTo*

:

**freeze** *sheet* *index*

:

**presetstyles**

:

**presetsheets**

:

**view** *args*

: 

**write** *filename*
    
:    Writes the spreadsheet to the file *filename*.



# DEPENDENCIES

    Tcl >= 8.6.7
    tclvfs::zip >= 1.4.2 (or Tcl 9)
    tdom >= 0.9.0
