// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>


#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

#include <atlbase.h>
#include <atlstr.h>

// Change these to point to your version of Microsoft Office
#import "C:\Program Files\Common Files\Microsoft Shared\Office16\MSO.DLL" rename_namespace("MSO")
#import "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB" rename_namespace("VB")
#import "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE" inject_statement("#include \"MSO.tlh\"") inject_statement("#include \"vbe6ext.tlh\"") exclude("IRange", "IDummy", "IFont", "IPicture") rename("RGB","ExcelRGB") rename("DialogBox", "ExcelDialogBox") rename("CopyFile", "ExcelCopyFile") rename_namespace("ExcelAutomation")

#import "C:\WINDOWS\system32\msxml6.dll" rename_namespace("MSXML")

#include "xlcall.h"
#include "framewrk.h"

void LogMessage(_bstr_t Message, bool Popup=false);

ExcelAutomation::_ApplicationPtr GetExcelApplicationPtr(HWND);

_bstr_t GetToday();
_bstr_t GetNow();

