# ExcelXML
A C++ dll to import/export named ranges to/from Excel from/to XML

Use the dll from VBA to transfer data back & forth between Excel and an XML document.

For example, an Excel table like this...

![alt text](https://github.com/sdb317/ExcelXML/blob/master/ExcelTable.png?raw=true "An example of an Excel table")

...becomes an XML document, like this...

![alt text](https://github.com/sdb317/ExcelXML/blob/master/XMLdoc.png?raw=true "An example of an exported XML document")

...and vice versa!

## Try it out with **Test.xlsm**

The VBA code is straightforward. It registers the function names from the dll and then makes sure that the dll is on the path (and can be found).

Then from Excel to XML:

```
Public Sub TestRangeToDocument()
    On Error GoTo Catch
    If Not RangeToDocument("SixNations", WorkbookDir & "\Test.xml") Then
        MsgBox GetLastErrorMessage()
    End If
Exit Sub
Catch:
    MsgBox Err.Description
End Sub
```

...and from XML to Excel:

```
Public Sub TestDocumentToRange()
    On Error GoTo Catch
    If Not DocumentToRange(WorkbookDir & "\Test.xml", "$A$10", "SixNations", "Sheet1", "*", "20170322") Then
        MsgBox GetLastErrorMessage()
    End If
Exit Sub
Catch:
    MsgBox Err.Description
End Sub
```

A log file is created in %localappdata%/Temp with the name ExcelXML_<PID>.log. Also errors can be retrieved in VBA by using `GetLastErrorMessage()`.

The full api consists of:

```
WorkbookToDocument // Takes all worksheets' used ranges in an Excel workbook and writes them to an XML document
WorksheetToDocument // Takes a worksheet's used range in an Excel spreadsheet and writes it to an XML document
RangeToDocument // Takes a range in an Excel workbook and writes it to a node in an XML document
DocumentToRange // Reads a node in an XML document and writes it to a range in an Excel workbook
```

There are also some self-explanatory options that can be set to true or false (default):

```
IncludeErrorValues=false; // Otherwise ommitted
IncludeEmptyValues=false; // Otherwise ommitted
ConstrainedByTarget=false; // Make sure only the target range in Excel is overwritten
```
---

Caveats

- The pre-compiled header `stdafx.h` may need to be changed depending on the Excel version (currently targets 2016)

- The dll export mechanism may need to change depending on whether Excel is 32-bit or 64-bit ([more details...](https://msdn.microsoft.com/en-us/library/office/bb687861.aspx))

- `ROT.cpp` does some pretty 'black belt' hacking to make sure the automation interface is for the right Excel instance. If there are multiple Excel sessions open with the same main window title, this may have difficulties.

