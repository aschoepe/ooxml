% ooxml(n) | ooxml user documentation
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

During this documentation the term workbook means a whole spreadsheet
with all its tables. The term worksheet means a specific table out of
to worksheet it consits to.

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

:   Defines an number format for the current workbook and
    returns an ID to refer that style.
    
    The following options are currently supported:
    
    :   -format FORMAT
    
        FORMAT can be any Excel format-string 
        
    :   -general

         Excel general-format
         
    :   -date
    
        Date format
        
    :   -time
    
        Time format 
        
    :   -datetime
    
        Date/Time format 
        
    :   -iso8601
    
        Date/Time in ISO8601 notation 
        
    :   -number
    
        Integer
    
    :   -decimal
    
        Decimal number with 2 decimal places 
    :   -red
    
        Color red on negative values (can be combined with number and
        decimal) 
        
    :   -separator
    
        Thousand separators (can be combined with number and decimal)
        
    :   -fraction
    
        Fractions 
        
    :   -scientific
    
        Scientific numbers 
        
    :   -percent
    
        Percentage 
        
    :   -text | -string
    
        Text
        
    :   -tag TAGNAME
    
        This option gives the format a name. This name may be used
        instead of the returned ID. 

**font** *args*

:   Defines a font for the current workbook and returns an ID to refer
    that style.
    
    The following options are currently supported:
    

    : -list
      Returns the list of currently defined fonts, in stead of an ID.
      
    : -name NAME
    
      (default = "Calibri")
      
    : -family FAMILY

      (defauft = 2) 
      
    : -size SIZE
    
      (default = 12)
      
    : -color COLOR
    
      (default = "theme 1")
      
    : -scheme SCHEME
    
      (default = "minor") 
    : -bold
    : -italic
    : -underline
    : -tag TAGNAME
    
      This option gives the font a name. This name may be used instead
      of the returned ID.

**fill** *args*

:

**border** *args*

:

**style** *args*

:

**defaultdatestyle** *STYLEID*

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


# COLOR

Serveral method options expect a *COLOR* argument. There are several
kinds of valid values.

The value may be *auto* or *none*. 

If not, the value may be an integer between 0 and 65 (including). See
the list below. 

If not, the value may be the name of a predefined color. Case is
ignored. See the list below.

If not, the value may be a 6 digit hexadecimal number, which is then
used as RGB value.


# DEPENDENCIES

    Tcl >= 8.6.7
    tclvfs::zip >= 1.4.2 (or Tcl 9)
    tdom >= 0.9.0
