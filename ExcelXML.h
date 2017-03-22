// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the ExcelXML_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// EXCELXML_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef EXCELXML_EXPORTS
#define EXCELXML_API extern "C" __declspec(dllexport)
#else
#define EXCELXML_API __declspec(dllimport)
#endif

EXCELXML_API BSTR WINAPI GetLastErrorMessage();
EXCELXML_API VARIANT_BOOL WINAPI WorkbookToDocument(LPCSTR pWorkbookName,LPCSTR pDocumentName);
EXCELXML_API VARIANT_BOOL WINAPI WorksheetToDocument(LPCSTR pWorksheetName,LPCSTR pDocumentName);
EXCELXML_API VARIANT_BOOL WINAPI RangeToDocument(LPCSTR pRangeName,LPCSTR pDocumentName);
EXCELXML_API VARIANT_BOOL WINAPI DocumentToRange(LPCSTR pDocumentName,LPCSTR pRangeName,LPCSTR pSourceRangeName=NULL,LPCSTR pSourceWorksheetName=NULL,LPCSTR pSourceWorkbookName=NULL,LPCSTR pSourceDate=NULL);
EXCELXML_API void WINAPI SetIncludeErrorValues(bool Include);
EXCELXML_API void WINAPI SetIncludeEmptyValues(bool Include);
EXCELXML_API void WINAPI SetConstrainedByTarget(bool Constrained);

