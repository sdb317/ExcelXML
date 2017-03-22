#ifndef _CEXCELARCHIVEDOCUMENT_H_7889DDC3_91A8_4CC6_9B40A3C45750
#define _CEXCELARCHIVEDOCUMENT_H_7889DDC3_91A8_4CC6_9B40A3C45750

///////////////////////////////////////////////////////////
// File :		CExcelArchiveDocument.h
// Created :	10/22/13
//

#include "CExcelDocument.h"

class CExcelArchiveDocument : public CExcelDocument
{
public:
// Construction/destruction

// Operations
	virtual bool Export();
	virtual bool Import();

protected:
	bool ExportRange(CNode VersionNode,ExcelAutomation::_WorkbookPtr pWorkbook,ExcelAutomation::_WorksheetPtr pWorksheet,ExcelAutomation::RangePtr pRange);
    bool ImportWorkbook(CNode VersionNode,ExcelAutomation::_WorkbookPtr pWorkbook,ExcelAutomation::_WorksheetPtr pWorksheet);
    bool ImportWorksheet(CNode WorkbookNode,ExcelAutomation::_WorksheetPtr pWorksheet);
    bool ImportRange(CNode WorksheetNode);
    bool ImportCell(CNode CellNode,ExcelAutomation::RangePtr pTargetRange);
	virtual _bstr_t GetRootElement() {return L"<Versions/>";}
	virtual _bstr_t GetWorkbookPath() {return L"";}
	virtual _bstr_t GetWorksheetPath() {return L"";}
	virtual _bstr_t GetRangePath() {return L"";}

// Attributes

};
#endif //_CEXCELARCHIVEDOCUMENT_H_7889DDC3_91A8_4CC6_9B40A3C45750
