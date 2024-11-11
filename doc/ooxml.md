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

**pageSetup** *sheet* *args*

:   This method sets the page setup of the worksheet. 
    
    **Caveat**: If for a sheet no page setup propertiers at all are
    set at least some applications use defaults derivated from other
    sources. But if at least one page setup propertier is explicitly
    set with this method some other propertiers not set will get the
    default value from the specification, which may be another value
    than without any page property setting. For example if you just
    set only the **-orientation** of a sheet then this may also result
    in another page size setting (because the application page size
    default may be differ from the default recommended by the
    specification (which would in this case Letter pape).
    
    The currently supported arguments are:
     
    **-blackAndWhite** *xsd boolean*
     
    :   Print black and white. Default: false
     
    **-cellComments** *(none|asDisplayed|atEnd)*
    
    :   How to print cell comments. Default: none
    
    **-copies** *integer*
    
    :   Number of copies to print. Default: 1
    
    **-draft** *xsd boolean*
    
    :   Print without graphics. Default: false
    
    **-errors** *(displayed|blank|dash|NA)*
    
    :   Specifies how to print cell values for cells with errors.
        Default: displayed
    
    **-firstPageNumber** *integer*
    
    :   Page number for first printed page. If no value is specified,
        then 'automatic' is assumed. Default: 1
  
    **-fitToHeight** *integer*
    
    :   Number of vertical pages to fit on. Default: 1
    
    **-fitToWidth** *integer*
    
    :   Number of horizontal pages to fit on. Default: 1
    
    **-horziontalDpi** *integer*
    
    :   Horizontal print resolution of the device. Default: 600
    
    **-orientation** *(default|portrait|landscape)*
    
    :   Orientation of the page. Default: default
    
    **-pageOrder** *(downThenOver|overThenDown)*
    
    :   Order of printed pages. Default: downThenOver
    
    **-paperHeight** *height*
    
    :   Height of custom paper as a number followed by a unit
        identifier (as 297mm or 11in).
        
        The specification say that when the *-paperHeight* and
        *-paperWidth* are specified, the application used to render
        shall ignore *-paperSize*.
        
        No default. 
        
     **-paperSize** *paperSizeID*
     
     :  Specifies the paper size to use by *paperSizeID*. 
     
        When *paperSizeID* values not present in the below list are
        used, the behavior is implementation- defined.
        
        The specification say that when the *-paperHeight* and
        *-paperWidth* are specified, the application used to render
        shall ignore *-paperSize*.
        
        Default: 1 (which is Letter Size)

        The possible *paperSizeID*s and their meaning are:
        
        Paper Size ID  Paper Size 
        -------------  ------------------
        1              Letter paper (8.5 in. by 11 in.)
        2              Letter small paper (8.5 in. by 11 in.)
        3              Tabloid paper (11 in. by 17 in.)
        4              Ledger paper (17 in. by 11 in.)
        5              Legal paper (8.5 in. by 14 in.)
        6              Statement paper (5.5 in. by 8.5 in.)
        7              Executive paper (7.25 in. by 10.5 in.)
        8              A3 paper (297 mm by 420 mm)
        9              A4 paper (210 mm by 297 mm)
        10             A4 small paper (210 mm by 297 mm)
        11             A5 paper (148 mm by 210 mm)
        12             B4 paper (250 mm by 353 mm)
        13             B5 paper (176 mm by 250 mm)
        14             Folio paper (8.5 in. by 13 in.)
        15             Quarto paper (215 mm by 275 mm)
        16             Standard paper (10 in. by 14 in.)
        17             Standard paper (11 in. by 17 in.)
        18             Note paper (8.5 in. by 11 in.)
        19             #9 envelope (3.875 in. by 8.875 in.)
        20             #10 envelope (4.125 in. by 9.5 in.)
        21             #11 envelope (4.5 in. by 10.375 in.)
        22             #12 envelope (4.75 in. by 11 in.)
        23             #14 envelope (5 in. by 11.5 in.)
        24             C paper (17 in. by 22 in.)
        25             D paper (22 in. by 34 in.)
        26             E paper (34 in. by 44 in.)
        27             DL envelope (110 mm by 220 mm)
        28             C5 envelope (162 mm by 229 mm)
        29             C3 envelope (324 mm by 458 mm)
        30             C4 envelope (229 mm by 324 mm)
        31             C6 envelope (114 mm by 162 mm)
        32             C65 envelope (114 mm by 229 mm)
        33             B4 envelope (250 mm by 353 mm)
        34             B5 envelope (176 mm by 250 mm)
        35             B6 envelope (176 mm by 125 mm)
        36             Italy envelope (110 mm by 230 mm)
        37             Monarch envelope (3.875 in. by 7.5 in.).
        38             6 3/4 envelope (3.625 in. by 6.5 in.)
        39             US standard fanfold (14.875 in. by 11 in.)
        40             German standard fanfold (8.5 in. by 12 in.)
        41             German legal fanfold (8.5 in. by 13 in.)
        42             ISO B4 (250 mm by 353 mm)
        43             Japanese double postcard (200 mm by 148 mm)
        44             Standard paper (9 in. by 11 in.)
        45             Standard paper (10 in. by 11 in.)
        46             Standard paper (15 in. by 11 in.)
        47             Invite envelope (220 mm by 220 mm)
        50             Letter extra paper (9.275 in. by 12 in.)
        51             Legal extra paper (9.275 in. by 15 in.)
        52             Tabloid extra paper (11.69 in. by 18 in.)
        53             A4 extra paper (236 mm by 322 mm)
        54             Letter transverse paper (8.275 in. by 11 in.)
        55             A4 transverse paper (210 mm by 297 mm)
        56             Letter extra transverse paper (9.275 in. by 12 in.)
        57             SuperA/SuperA/A4 paper (227 mm by 356 mm)
        58             SuperB/SuperB/A3 paper (305 mm by 487 mm)
        59             Letter plus paper (8.5 in. by 12.69 in.)
        60             A4 plus paper (210 mm by 330 mm)
        61             A5 transverse paper (148 mm by 210 mm)
        62             JIS B5 transverse paper (182 mm by 257 mm)
        63             A3 extra paper (322 mm by 445 mm)
        64             A5 extra paper (174 mm by 235 mm)
        65             ISO B5 extra paper (201 mm by 276 mm)
        66             A2 paper (420 mm by 594 mm)
        67             A3 transverse paper (297 mm by 420 mm)
        68             A3 extra transverse paper (322 mm by 445 mm)
        69             Japanese Double Postcard (200 mm x 148 mm)
        70             A6 (105 mm x 148 mm)
        71             Japanese Envelope Kaku #2
        72             Japanese Envelope Kaku #3
        73             Japanese Envelope Chou #3
        74             Japanese Envelope Chou #4
        75             Letter Rotated (11in x 8 1/2 11 in)
        76             A3 Rotated (420 mm x 297 mm)
        77             A4 Rotated (297 mm x 210 mm)
        78             A5 Rotated (210 mm x 148 mm)
        79             B4 (JIS) Rotated (364 mm x 257 mm)
        80             B5 (JIS) Rotated (257 mm x 182 mm)
        81             Japanese Postcard Rotated (148 mm x 100 mm)
        82             Double Japanese Postcard Rotated (148 mm x 200 mm)
        83             A6 Rotated (148 mm x 105 mm)
        84             Japanese Envelope Kaku #2 Rotated
        85             Japanese Envelope Kaku #3 Rotated
        86             Japanese Envelope Chou #3 Rotated
        87             Japanese Envelope Chou #4 Rotated
        88             B6 (JIS) (128 mm x 182 mm)
        89             B6 (JIS) Rotated (182 mm x 128 mm)
        90             (12 in x 11 in)
        91             Japanese Envelope You #4
        92             Japanese Envelope You #4 Rotated
        93             PRC 16K (146 mm x 215 mm)
        94             PRC 32K (97 mm x 151 mm)
        95             PRC 32K(Big) (97 mm x 151 mm)
        96             PRC Envelope #1 (102 mm x 165 mm)
        97             PRC Envelope #2 (102 mm x 176 mm)
        98             PRC Envelope #3 (125 mm x 176 mm)
        99             PRC Envelope #4 (110 mm x 208 mm)
        100            PRC Envelope #5 (110 mm x 220 mm)
        101            PRC Envelope #6 (120 mm x 230 mm)
        102            PRC Envelope #7 (160 mm x 230 mm)
        103            PRC Envelope #8 (120 mm x 309 mm)
        104            PRC Envelope #9 (229 mm x 324 mm)
        105            PRC Envelope #10 (324 mm x 458 mm)
        106            PRC 16K Rotated
        107            PRC 32K Rotated
        108            PRC 32K(Big) Rotated
        109            PRC Envelope #1 Rotated (165 mm x 102 mm)
        110            PRC Envelope #2 Rotated (176 mm x 102 mm)
        111            PRC Envelope #3 Rotated (176 mm x 125 mm)
        112            PRC Envelope #4 Rotated (208 mm x 110 mm)
        113            PRC Envelope #5 Rotated (220 mm x 110 mm)
        114            PRC Envelope #6 Rotated (230 mm x 120 mm)
        115            PRC Envelope #7 Rotated (230 mm x 160 mm)
        116            PRC Envelope #8 Rotated (309 mm x 120 mm)
        117            PRC Envelope #9 Rotated (324 mm x 229 mm)
        118            PRC Envelope #10 Rotated (458 mm x 324 mm)
        
     **-paperWide** *width*
     
     :  Wide of custom paper as a number followed by a unit
        identifier (as 21cm or 8.5in).
         
        The specification say that when the *-paperHeight* and
        *-paperWidth* are specified, the application used to render
        shall ignore *-paperSize*.
        
        No default.
         
     **-scale** *integer*
     
     :  Print scaling in percent. The rendering applicationi may
        respect only values ranging from 10 to 400.
        
        This setting is overridden when *-fitToWidth* and/or
        *-fitToHeight* are in use.
        
        Default: 100
        
     **-useFirstPageNumber** *xsd boolean*
     
     :  Use *-firstPageNumber* value for first page number, and do not
        auto number the pages. Default: false
        
     **-verticalDpi** *integer*
     
     :  Vertical print resolution of the device. Default: 600


**write** *filename*
    
:    Writes the spreadsheet to the file *filename*.



# DEPENDENCIES

    Tcl >= 8.6.7
    tclvfs::zip >= 1.4.2 (or Tcl 9)
    tdom >= 0.9.0
