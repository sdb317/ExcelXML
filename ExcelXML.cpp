// ExcelXML.cpp : Defines the exported functions for the DLL application.
/*

Private Declare Function GetLastErrorMessage Lib "ExcelXML.dll" () As String
Private Declare Function WorkbookToDocument Lib "ExcelXML.dll" (ByVal WorkbookName As String, ByVal DocumentName As String) As Boolean
Private Declare Function WorksheetToDocument Lib "ExcelXML.dll" (ByVal WorksheetName As String, ByVal DocumentName As String) As Boolean
Private Declare Function RangeToDocument Lib "ExcelXML.dll" (ByVal RangeName As String, ByVal DocumentName As String) As Boolean
Private Declare Function DocumentToRange Lib "ExcelXML.dll" (ByVal DocumentName As String, ByVal RangeName As String, ByVal pSourceRangeName As String, ByVal pSourceWorksheetName As String, ByVal pSourceWorkbookName As String, ByVal pSourceDate As String) As Boolean
Private Declare Sub SetIncludeErrorValues Lib "ExcelXML.dll" (ByVal Include As Boolean)
Private Declare Sub SetIncludeEmptyValues Lib "ExcelXML.dll" (ByVal Include As Boolean)
Private Declare Sub SetConstrainedByTarget Lib "ExcelXML.dll" (ByVal Constrained As Boolean)

*/
//

#include "stdafx.h"
#include "ExcelXML.h"

ExcelAutomation::_ApplicationPtr pApplication=NULL;
bool IncludeErrorValues=false;
bool IncludeEmptyValues=false;
bool ConstrainedByTarget=false;

#include "CExcelArchiveDocument.h"

extern _bstr_t LastError;

// For 32-bit use: #pragma comment(linker, "/EXPORT:GetLastErrorMessage=_GetLastErrorMessage@0,@1")
EXCELXML_API BSTR WINAPI GetLastErrorMessage() {return SysAllocStringByteLen((LPCSTR)LastError,LastError.length());}

// For 32-bit use: #pragma comment(linker, "/EXPORT:WorkbookToDocument=_WorkbookToDocument@8,@2")
EXCELXML_API VARIANT_BOOL WINAPI 
WorkbookToDocument // Takes all worksheets' used ranges in an Excel workbook and writes them to an XML document
    (
    LPCSTR pWorkbookName, // The full path of the workbook, or missing fo 'ThisWorkbook'
    LPCSTR pDocumentName // The full path of the XML document
    )
{
    LastError=L"";
    try
        {
        if (pApplication!=NULL)
            {
            ExcelAutomation::_WorkbookPtr pWorkbook;
            _bstr_t WorkbookName(pWorkbookName);
            if (WorkbookName.length()>0)
                {
                // pWorkbook=GetObject((LPCTSTR)WorkbookName);
                }
            else
                {
                pWorkbook=((ExcelAutomation::_WorkbookPtr)pApplication->ThisWorkbook);
                }
            if (pWorkbook!=NULL)
                {
                }
            else
                {
                LogMessage(_bstr_t(L"Error connecting to \'Workbook\'"));
                }
            }
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        LogMessage(_bstr_t(L"Error in \'WorkbookToDocument\': ")+e.ErrorMessage());
        return VARIANT_FALSE;
        }
    return VARIANT_TRUE;
}

// For 32-bit use: #pragma comment(linker, "/EXPORT:WorksheetToDocument=_WorksheetToDocument@8,@3")
EXCELXML_API VARIANT_BOOL WINAPI 
WorksheetToDocument // Takes a worksheet's used range in an Excel spreadsheet and writes it to an XML document
    (
    LPCSTR pWorksheetName, // The name of the worksheet, or missing for the 'ActiveSheet'
    LPCSTR pDocumentName // The full path of the XML document
    )
{
    LastError=L"";
    try
        {
        if (pApplication!=NULL)
            {
            ExcelAutomation::_WorksheetPtr pWorksheet;
            _bstr_t WorksheetName(pWorksheetName);
            if (WorksheetName.length()>0)
                {
                pWorksheet=((ExcelAutomation::_WorksheetPtr)pApplication->ThisWorkbook->Worksheets->Item[WorksheetName]);
                }
            else
                {
                pWorksheet=((ExcelAutomation::_WorksheetPtr)pApplication->ActiveSheet);
                }
            if (pWorksheet!=NULL)
                {
                }
            else
                {
                LogMessage(_bstr_t(L"Error connecting to \'Worksheet\'"));
                }
            }
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        LogMessage(_bstr_t(L"Error in \'WorksheetToDocument\': ")+e.ErrorMessage());
        return VARIANT_FALSE;
        }
    return VARIANT_TRUE;
}

// For 32-bit use: #pragma comment(linker, "/EXPORT:RangeToDocument=_RangeToDocument@8,@4")
EXCELXML_API VARIANT_BOOL WINAPI 
RangeToDocument // Takes a range in an Excel workbook and writes it to a node in an XML document
    (
    LPCSTR pRangeName, // The workbook range or worksheet range for the 'ActiveSheet'
    LPCSTR pDocumentName // The full path of the XML document
    )
{
    LastError=L"";
    try
        {
        if (pApplication!=NULL)
            {
            _bstr_t RangeName(pRangeName);
            _bstr_t DocumentName(pDocumentName);
            CExcelArchiveDocument ExcelArchiveDocument;
            if (ExcelArchiveDocument.SetDocumentName(DocumentName))
                {
                ExcelArchiveDocument.SetApplication(pApplication);
                TCHAR Delimiter[]=_T(",");
                if (_tcsstr((LPCTSTR)RangeName,Delimiter)) // Multiple ranges in one call
                    {
                    TCHAR* pNextRangeName=NULL;
                    TCHAR* pContext=(LPTSTR)RangeName;
                    bool Status=true; // Assume all will work
                    while ((pNextRangeName=_tcstok_s(pContext,Delimiter,&pContext))!= NULL)
                        {
                        _bstr_t NextRangeName(pNextRangeName);
                        ExcelArchiveDocument.SetRange(NextRangeName);
                        if (!ExcelArchiveDocument.Export())
                            Status=false;
                        }
                    if (Status) // If all succeeded
                        {
                        ExcelArchiveDocument->save(DocumentName);
                        return VARIANT_TRUE;
                        }
                    else
                        {
                        LogMessage(L"Export failed");
                        return VARIANT_FALSE;
                        }
                    }
                else
                    {
                    ExcelArchiveDocument.SetRange(RangeName); // A single range
                    if (ExcelArchiveDocument.Export())
                        {
                        ExcelArchiveDocument->save(DocumentName);
                        return VARIANT_TRUE;
                        }
                    else
                        {
                        LogMessage(L"Export failed");
                        return VARIANT_FALSE;
                        }
                    }
                }
            else
                {
                LogMessage(L"Unable to load an existing document or create a new document");
                return VARIANT_FALSE;
                }
            }
        else
            {
            LogMessage(L"Unable to connect to Excel");
            return VARIANT_FALSE;
            }
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        LogMessage(_bstr_t(L"Error in \'RangeToDocument\': ")+e.ErrorMessage());
        return VARIANT_FALSE;
        }
    return VARIANT_FALSE;
}

// For 32-bit use: #pragma comment(linker, "/EXPORT:DocumentToRange=_DocumentToRange@24,@5")
EXCELXML_API VARIANT_BOOL WINAPI 
DocumentToRange // Reads a node in an XML document and writes it to a range in an Excel workbook
    (
    LPCSTR pDocumentName, // The full path of the XML document
    LPCSTR pRangeName, // The identifier for the workbook range or worksheet range in Excel to write to
    LPCSTR pSourceRangeName, // If specified, the source range node in the document (if different to the target)
    LPCSTR pSourceWorksheetName, // If specified, the source worksheet node in the document (if different to the target)
    LPCSTR pSourceWorkbookName, // If specified, the source workbook node in the document (if different to the target)
    LPCSTR pSourceDate // If specified, the source date node in the document (if different to the target)
    )
{
    LastError=L"";
    try
        {
        if (pApplication!=NULL)
            {
            _bstr_t DocumentName(pDocumentName);
            _bstr_t RangeName(pRangeName);
            _bstr_t SourceRangeName(pSourceRangeName);
            _bstr_t SourceWorksheetName(pSourceWorksheetName);
            _bstr_t SourceWorkbookName(pSourceWorkbookName);
            _bstr_t SourceDate(pSourceDate);
            CExcelArchiveDocument ExcelArchiveDocument;
            if (ExcelArchiveDocument.SetDocumentName(DocumentName))
                {
                ExcelArchiveDocument.SetApplication(pApplication);
                ExcelArchiveDocument.SetRange(RangeName);
                ExcelArchiveDocument.SetSourceRange(SourceRangeName);
                ExcelArchiveDocument.SetSourceWorksheet(SourceWorksheetName);
                ExcelArchiveDocument.SetSourceWorkbook(SourceWorkbookName);
                ExcelArchiveDocument.SetSourceDate(SourceDate);
                if (ExcelArchiveDocument.Import())
                    {
                    return VARIANT_TRUE;
                    }
                else
                    {
                    LogMessage(L"Import failed");
                    return VARIANT_FALSE;
                    }
                }
            else
                {
                LogMessage(L"Unable to load the document");
                return VARIANT_FALSE;
                }
            }
        else
            {
            LogMessage(L"Unable to connect to Excel");
            return VARIANT_FALSE;
            }
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        LogMessage(_bstr_t(L"Error in \'DocumentToRange\': ")+e.ErrorMessage());
        return VARIANT_FALSE;
        }
    return VARIANT_FALSE;
}

// For 32-bit use: #pragma comment(linker, "/EXPORT:SetIncludeErrorValues=_SetIncludeErrorValues@4,@6")
EXCELXML_API void WINAPI SetIncludeErrorValues(bool Include) {IncludeErrorValues=Include;}

// For 32-bit use: #pragma comment(linker, "/EXPORT:SetIncludeEmptyValues=_SetIncludeEmptyValues@4,@7")
EXCELXML_API void WINAPI SetIncludeEmptyValues(bool Include) {IncludeEmptyValues=Include;}

// For 32-bit use: #pragma comment(linker, "/EXPORT:SetConstrainedByTarget=_SetConstrainedByTarget@4,@8")
EXCELXML_API void WINAPI SetConstrainedByTarget(bool Constrained) {ConstrainedByTarget=Constrained;}

