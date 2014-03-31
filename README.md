BibleRefParser
==============

This simple project creates a .NET DLL that can parse an incoming reference string, e.g. Jn 3:16 and return either a short reference, e.g. Jn 3:16 or a long reference John 3:16.

It supports a flexible input of ranges and lists, e.g. Jn 3:16-20, 4:1;5:6. and can return a list of verses:
John 3:16
John 3:17
John 3:18
John 3:19
John 3:20
John 4:1
John 5:6

The main object is the library. The library has the parser where you send the text to parse:

Library.BibleReference.ParseReference("Jn 3:16-20, 4:1;5:6", false)

The library also has a path property. If you provide the library with a path, it will scan the directory for Bible texts in a simple text format. One line per verse, prefixed with book chapter verse. Genesis 1:1 from KJV in the file would look like:

01001001 In the beginning God created the heaven and the earth.

These files are not included in this project.

The code was written in VB.NET with Visual Studio 2013.

Enjoy.
