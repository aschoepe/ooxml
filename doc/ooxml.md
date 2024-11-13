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
    
      See [COLOR](#color) for the valid values.
      
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
kind of valid values.

The value may be *auto* or *none*. 

If not, the value may be the index number of a predefined color (an
integer between 0 and 65 (including)). See the list below.

If not, the value may be the name of a predefined color. Case is
ignored. See the list below.

If not, the value may be a 6 digit hexadecimal number, which is then
used as RGB value.

The list of predefined colors is:

Color ID Name             (A)RGB
-------- ---------------  --------
0        Black            00000000
1        White            00FFFFFF
2        Red              00FF0000
3        Lime             0000FF00
4        Blue             000000FF
5        Yellow           00FFFF00
6        Fuchsia          00FF00FF
7        Aqua             0000FFFF
8        Black            00000000
9        White            00FFFFFF
10       Red              00FF0000
11       Lime             0000FF00
12       Blue             000000FF
13       Yellow           00FFFF00
14       Fuchsia          00FF00FF
15       Aqua             0000FFFF
16       Maroon           00800000
17       Green            00008000
18       Navy             00000080
19       Olive            00808000
20       Purple           00800080
21       Teal             00008080
22       Silver           00C0C0C0
23       Gray             00808080
24       Portage          009999FF
25       Lipstick         00993366
26       Cream            00FFFFCC
27       LightCyan        00CCFFFF
28       Purple           00660066
29       LightCoral       00FF8080
30       NavyBlue         000066CC
31       LavenderBlue     00CCCCFF
32       Navy             00000080
33       Fuchsia          00FF00FF
34       Yellow           00FFFF00
35       Aqua             0000FFFF
36       Purple           00800080
37       Maroon           00800000
38       Teal             00008080
39       Blue             000000FF
40       DeepSkyBlue      0000CCFF
41       LightCyan        00CCFFFF
42       BlueRomance      00CCFFCC
43       Canary           00FFFF99
44       LightSkyBlue     0099CCFF
45       CarnationPink    00FF99CC
46       Mauve            00CC99FF
47       PeachOrange      00FFCC99
48       RoyalBlue        003366FF
49       MediumTurquoise  0033CCCC
50       Citrus           0099CC00
51       TangerineYellow  00FFCC00
52       OrangePeel       00FF9900
53       SafetyOrange     00FF6600
54       Scampi           00666699
55       Nobel            00969696
56       PrussianBlue     00003366
57       Eucalyptus       00339966
58       Myrtle           00003300
59       Karaka           00333300
60       SaddleBrown      00993300
61       Lipstick         00993366
62       DarkSlateBlue    00333399
63       NightRider       00333333
64       SystemForeground n/a
65       SystemBackground n/a

# DEPENDENCIES

    Tcl >= 8.6.7
    tclvfs::zip >= 1.4.2 (or Tcl 9)
    tdom >= 0.9.0
