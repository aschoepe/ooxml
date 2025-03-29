
# RESOURCES

There is indeed an official and free available standard for
WordprocessingML:
<https://standards.iso.org/ittf/PubliclyAvailableStandards/index.html>
and look for the ISO/IEC 29500 documents.

You for sure will need the pdf document with a titel as "INTERNATIONAL
STANDARD ISO/IEC 29500-1 Fourth edition 2016-11-01 Part 1:
Fundamentals and Markup Language Reference". The file name is
something crude as C071691e.pdf.

Very helpful (indispensable perhaps) is also wml.xsd

For introductional material to WordprocessingML see
<http://www.officeopenxml.com/anatomyofOOXML.php>

Microsoft has
<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offfflp>

Other libraries do similar things.

For python see <https://python-docx.readthedocs.io/en/latest/>.

For C# and .net see <https://github.com/dotnet/Open-XML-SDK>

Good Tcl level is required. And tDOM knowledge is helpful, or course.

# BASICS

Working with Tcl is mostly a pleasure. Working with tDOM hopefully is
and working with this Tcl docx framework hopefully will be too.
Working with WordprocessingML is not always a pleasure. In some
hindsight it looks more like the translation of a binary format into
XML than as a well-designed document markup.

At least it is a voluminous beast - which is kind of inevitable (to
some degree) for a powerful layout markup language.

Almost all WordprocessingML elements have no text content. 


# GENERATING ELEMENTS AND ATTRIBUTES ON USER INPUT



The name of the option, as -bold or -margins, points to a Tcl List, a
list of at least two elements.

The first element of that list is the list of elements to create. 

The second argument is a list with information about the attributes to
set and how the values for the attributes are extracted from the value
given to the option and the types of that value.

If the second argument has only one element and that element itself is
not a list with more than 1 elements then the name of the
attribute to create is *w:val* and this element is the type callback
for the value.

Otherwise every element of the second argument list itself is threated
as a list. If the second argument list has one element 
then the name of created attribute will be the first element of that
element and the value, after calling the as second element given type
callback. will be the value of the option.

Otherwise - which means that this option may set more than one
attribute at once - the value to the option is seen as a keyword value
list. 

If the list element has two members then the name of the keyword and
the name of the attribute to create is the first element of the list
and the second argument the type callback.

If the list element has three members then the the first element of
that list is the keyword, the second element is the name of the
attribute to create and the third element ist the type callback.

If the list the option name points to has three elements then the
second list element will be ignored and the processing is done by the
specialized procedure given with the third element.

If the second argument has two elements then 
