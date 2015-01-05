wordTCL
=======

Introduction
------------

Some scripts to create docx Word 2010


These TCL scripts can be used to create simple DOCX file (MS Office 2010 format).

You need tcl version 8.6 (not actually tested with tcl 8.5 or others).
Some pachake are needed also:
  + tdom: for reading/creating XML file
  + vfs::zip: for opening DOCX format (used for reading files)
  + TclOO: tcl object implementation

There is no release version for this moment, only this draft version.

Features
--------

* Create new document based on a template file (empty docx file)
* Add paragraph (normal)
* Add paragraph of type 'heading'
* Add images (only external images)

How to use
----------

See "testLibWord.tcl" and "testLibWordList.tcl"


Future
------

For now I have to make the first release. 
Update zipper.tcl to use only vfs for crc calculation (for now there is a patch for linux platform). 
Add documentation on source code.
Adding more Word Element support  


