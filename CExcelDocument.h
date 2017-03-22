#ifndef _CEXCELDOCUMENT_H_3F1893B0_7116_4545_B868841EBEC
#define _CEXCELDOCUMENT_H_3F1893B0_7116_4545_B868841EBEC

///////////////////////////////////////////////////////////
// File :		CExcelDocument.h
// Created :	10/22/13
//

#include "CDocument.h"

class CExcelDocument : public CDocument
{
public:
// Construction/destruction

// Operations
	_bstr_t GetDocumentName() const;
	bool SetDocumentName(_bstr_t value);
	virtual bool Export();
	virtual bool Import();
	void SetApplication(ExcelAutomation::_ApplicationPtr value) {pApplication=value;}
	void SetWorkbook(ExcelAutomation::_WorkbookPtr value) {pWorkbook=value;}
	void SetWorksheet(ExcelAutomation::_WorksheetPtr value) {pWorksheet=value;}
	void SetRange(ExcelAutomation::RangePtr value) {pRange=value;}
	void SetRange(_bstr_t value) {RangeName=value;try{pRange=pApplication->GetRange(RangeName);} catch(_com_error& e){LogMessage(_bstr_t(L"Error in \'CDocument::SetRange\': ")+e.ErrorMessage());}}
	ExcelAutomation::RangePtr GetRange() {return pRange;}
	void SetSourceRange(_bstr_t value) {SourceRangeName=value;}
	_bstr_t GetSourceRange() {return SourceRangeName;}
	void SetSourceWorksheet(_bstr_t value) {SourceWorksheetName=value;}
	_bstr_t GetSourceWorksheet() {return SourceWorksheetName;}
	void SetSourceWorkbook(_bstr_t value) {SourceWorkbookName=value;}
	_bstr_t GetSourceWorkbook() {return SourceWorkbookName;}
	void SetSourceDate(_bstr_t value) {SourceDate=value;}
	_bstr_t GetSourceDate() {return SourceDate;}

protected:
	bool ExportWorkbook(MSXML::IXMLDOMNodePtr pNode);
	bool ExportWorksheet(MSXML::IXMLDOMNodePtr pNode);
	bool ExportRange(MSXML::IXMLDOMNodePtr pNode);
	virtual _bstr_t GetRootElement()=0;
	virtual _bstr_t GetWorkbookPath()=0;
	virtual _bstr_t GetWorksheetPath()=0;
	virtual _bstr_t GetRangePath()=0;

// Attributes

private:
	_bstr_t DocumentName;

protected:
	bool Replace;
	bool SaveFormulae;
	bool SaveFormatting;
	ExcelAutomation::_ApplicationPtr pApplication;
	ExcelAutomation::_WorkbookPtr pWorkbook;
	ExcelAutomation::_WorksheetPtr pWorksheet;
	ExcelAutomation::RangePtr pRange;
	_bstr_t RangeName;
	_bstr_t SourceRangeName;
    _bstr_t SourceWorksheetName;
    _bstr_t SourceWorkbookName;
    _bstr_t SourceDate;

};
#endif //_CEXCELDOCUMENT_H_3F1893B0_7116_4545_B868841EBEC
