/* ==========================================================================
	File :			CExcelArchiveDocument.cpp
	
	Class :			CExcelArchiveDocument

	Date :			10/22/13

	Purpose :		

	Description :	

	Usage :			

   ========================================================================*/

#include "stdafx.h"
#include "propvarutil.h"
#include "CExcelArchiveDocument.h"

extern bool IncludeErrorValues;
extern bool IncludeEmptyValues;
extern bool ConstrainedByTarget;

////////////////////////////////////////////////////////////////////
// Public functions
//

bool CExcelArchiveDocument::Export()
/* ============================================================
	Function :		CExcelArchiveDocument::Export
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	none

	Usage :			

   ============================================================*/
{
    try
        {
        CNode VersionsNode(GetInterfacePtr()->documentElement);
        CNode VersionNode=FindNode(_bstr_t(_T("Version[@Name=\""))+(GetSourceDate().length()?GetSourceDate():GetToday())+_bstr_t(_T("\"]")));
        if (VersionNode.IsEmpty())
            {
            VersionNode=VersionsNode.AppendChild(L"Version");
            VersionNode.SetAttribute(L"Name",GetToday());
            }
        if (GetRange()!=NULL)
            {
            ExcelAutomation::_WorksheetPtr pWorksheet=GetRange()->Worksheet;
            ExcelAutomation::_WorkbookPtr pWorkbook=pWorksheet->Parent;
            return ExportRange(VersionNode,pWorkbook,pWorksheet,GetRange());
            }
        else
            {
            LogMessage(L"Range is empty");
            return false;
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::Export\': ")+e.ErrorMessage());
        return false;
        }
    return true;
}

bool CExcelArchiveDocument::Import()
/* ============================================================
	Function :		CExcelArchiveDocument::Import
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	none

	Usage :			

   ============================================================*/
{
    try
        {
        CNode VersionsNode(GetInterfacePtr()->documentElement);
        CNode VersionNode=FindNode(_bstr_t(_T("Version[@Name=\""))+(GetSourceDate().length()?GetSourceDate():GetToday())+_bstr_t(_T("\"]")));
        if (!VersionNode.IsEmpty())
            {
            if (GetRange()!=NULL)
                {
                ExcelAutomation::_WorksheetPtr pWorksheet=GetRange()->Worksheet;
                ExcelAutomation::_WorkbookPtr pWorkbook=pWorksheet->Parent;
                return ImportWorkbook(VersionNode,pWorkbook,pWorksheet);
                }
            else
                {
                LogMessage(L"Range is empty");
                return false;
                }
            }
        else
            {
            LogMessage(L"Version is empty");
            return false;
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::Import\': ")+e.ErrorMessage());
        return false;
        }
    return true;
}

////////////////////////////////////////////////////////////////////
// Protected functions
//

////////////////////////////////////////////////////////////////////
// Private functions
//

bool CExcelArchiveDocument::ExportRange(CNode VersionNode,ExcelAutomation::_WorkbookPtr pWorkbook,ExcelAutomation::_WorksheetPtr pWorksheet,ExcelAutomation::RangePtr pRange)
/* ============================================================
	Function :		CExcelArchiveDocument::ExportRange
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	

	Usage :			

   ============================================================*/
{
    try
        {   
        CNode WorkbooksNode=FindNode(_bstr_t(_T("Workbooks")),VersionNode);
        if (WorkbooksNode.IsEmpty())
            {
            WorkbooksNode=VersionNode.AppendChild(L"Workbooks");
            }
        if (!WorkbooksNode.IsEmpty())
            {
            CNode WorkbookNode=FindNode(_bstr_t(_T("Workbook[@Name=\""))+(GetSourceWorkbook().length()?GetSourceWorkbook():pWorkbook->Name)+_bstr_t(_T("\"]")),WorkbooksNode); // pWorkbook->FullName
            if (WorkbookNode.IsEmpty())
                {
                WorkbookNode=WorkbooksNode.AppendChild(L"Workbook");
                WorkbookNode.SetAttribute(L"Name",pWorkbook->Name); // pWorkbook->FullName
                }
            if (!WorkbookNode.IsEmpty())
                {
                CNode WorksheetsNode=FindNode(_bstr_t(_T("Worksheets")),WorkbookNode);
                if (WorksheetsNode.IsEmpty())
                    {
                    WorksheetsNode=WorkbookNode.AppendChild(L"Worksheets");
                    }
                if (!WorksheetsNode.IsEmpty())
                    {
                    CNode WorksheetNode=FindNode(_bstr_t(_T("Worksheet[@Name=\""))+(GetSourceWorksheet().length()?GetSourceWorksheet():pWorksheet->Name)+_bstr_t(_T("\"]")),WorksheetsNode);
                    if (WorksheetNode.IsEmpty())
                        {
                        WorksheetNode=WorksheetsNode.AppendChild(L"Worksheet");
                        WorksheetNode.SetAttribute(L"Name",_bstr_t(pWorksheet->Name));
                        }
                    if (!WorksheetNode.IsEmpty())
                        {
                        CNode RangesNode=FindNode(_bstr_t(_T("Ranges")),WorksheetNode);
                        if (RangesNode.IsEmpty())
                            {
                            RangesNode=WorksheetNode.AppendChild(L"Ranges");
                            }
                        if (!RangesNode.IsEmpty())
                            {
                            CNode RangeNode=FindNode(_bstr_t(_T("Range[@Name=\""))+RangeName+_bstr_t(_T("\"]")),RangesNode);
                            if (RangeNode.IsEmpty())
                                {
                                RangeNode=RangesNode.AppendChild(L"Range");
                                }
                            else
                                {
                                CNode NewRangeNode(*this,_bstr_t(_T("Range")));
                                RangesNode.ReplaceChild(RangeNode,NewRangeNode);
                                RangeNode=NewRangeNode;
                                }
                            if (!RangeNode.IsEmpty())
                                {
                                RangeNode.SetAttribute(L"Name",RangeName);
                                RangeNode.SetAttribute(L"Address",_bstr_t(pRange->Name));
                                long RowCount=pRange->Rows->Count;
                                RangeNode.SetAttribute(L"Rows",_bstr_t((long)RowCount));
                                long ColumnCount=pRange->Columns->Count;
                                RangeNode.SetAttribute(L"Columns",_bstr_t((long)ColumnCount));
                                long RowOffset=pRange->Row;
                                long ColumnOffset=pRange->Column;
                                for (long i=0;i<RowCount;i++)
                                    {
                                    for (long j=0;j<ColumnCount;j++)
                                        {
                                        try
                                            {
			                                ExcelAutomation::RangePtr pCell=pWorksheet->Cells->Item[i+RowOffset,j+ColumnOffset];
			                                _variant_t Value=pCell->Value2;
                                            if ((Value!=_variant_t())||IncludeEmptyValues) // Not empty or include them anyway
                                                {
                                                if ((Value.vt!=VT_ERROR)||IncludeErrorValues) // Not an Excel '#' error or include them anyway
			                                        {
													CNode CellNode=RangeNode.AppendChild(L"Cell");
													CellNode.SetAttribute(L"Row",_bstr_t((long)i));
													CellNode.SetAttribute(L"Column",_bstr_t((long)j));
													CellNode.SetAttribute(L"Address",pCell->Address[0L][0L][ExcelAutomation::xlA1][0L][0L]);
                                                    if (Value!=_variant_t()) // Not empty
                                                        {
                                                        if (Value.vt!=VT_ERROR) // Not an Excel '#' error
			                                                {
                                                            CellNode.SetAttribute(L"Value",(_bstr_t)Value); // .replace(/'/g,'&quot;').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
                                                            if (IsVariantString(Value))
                                                                {
                                                                CellNode.SetAttribute(L"Type",L"Text");
                                                                }
                                                            else
                                                                {
                                                                CellNode.SetAttribute(L"Type",L"Number");
                                                                }
			                                                }
                                                        else
			                                                {
                                                            CellNode.SetAttribute(L"Type",L"Error");
                                                            switch (LOWORD(((*(tagVARIANT*)(&Value))).scode))
                                                                {
                                                                case ExcelAutomation::xlErrDiv0:
                                                                    CellNode.SetAttribute(L"Value",L"#DIV/0");
                                                                    break;
                                                                case ExcelAutomation::xlErrNA:
                                                                    CellNode.SetAttribute(L"Value",L"#N/A");
                                                                    break;
                                                                case ExcelAutomation::xlErrName:
                                                                    CellNode.SetAttribute(L"Value",L"#NAME");
                                                                    break;
                                                                case ExcelAutomation::xlErrNull:
                                                                    CellNode.SetAttribute(L"Value",L"#NULL");
                                                                    break;
                                                                case ExcelAutomation::xlErrNum:
                                                                    CellNode.SetAttribute(L"Value",L"#NUM");
                                                                    break;
                                                                case ExcelAutomation::xlErrRef:
                                                                    CellNode.SetAttribute(L"Value",L"#REF");
                                                                    break;
                                                                case ExcelAutomation::xlErrValue:
                                                                    CellNode.SetAttribute(L"Value",L"#VALUE");
                                                                    break;
                                                                default:;
                                                                }
                                                            }
                                                        }
			                                        }
                                                }
			                                }
                                        catch (_com_error& e)
                                            {
                                            LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::ExportRange\' cell: ")+e.ErrorMessage());
                                            }
                                        }
                                    }
                                }
                            else
                                {
                                LogMessage(L"Range is empty");
                                return false;
                                }
                            }
                        else
                            {
                            LogMessage(L"Range is empty");
                            return false;
                            }
                        }
                    else
                        {
                        LogMessage(L"Worksheet is empty");
                        return false;
                        }
                    }
                else
                    {
                    LogMessage(L"Worksheets is empty");
                    return false;
                    }
                }
            else
                {
                LogMessage(L"Workbook is empty");
                return false;
                }
            }
        else
            {
            LogMessage(L"Workbooks is empty");
            return false;
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::ExportRange\': ")+e.ErrorMessage());
        return false;
        }
    return true;
}

bool CExcelArchiveDocument::ImportWorkbook(CNode VersionNode,ExcelAutomation::_WorkbookPtr pWorkbook,ExcelAutomation::_WorksheetPtr pWorksheet)
/* ============================================================
	Function :		CExcelArchiveDocument::ImportWorkbook
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	Workbook name to query for, or '*' for any

	Usage :			

   ============================================================*/
{
    try
        {
        CNode WorkbookNode=FindNode(_bstr_t(_T("Workbooks/Workbook"))+(GetSourceWorkbook()==_bstr_t(_T("*"))?_bstr_t(_T("")):_bstr_t(_T("[@Name=\""))+(GetSourceWorkbook().length()?GetSourceWorkbook():pWorkbook->Name)+_bstr_t(_T("\"]"))),VersionNode); // pWorkbook->FullName
        if (!WorkbookNode.IsEmpty())
            {
            return ImportWorksheet(WorkbookNode,pWorksheet);
            }
        else
            {
            LogMessage(L"Workbook is empty");
            return false;
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::ImportWorkbook\': ")+e.ErrorMessage());
        return false;
        }
    return false;
}

bool CExcelArchiveDocument::ImportWorksheet(CNode WorkbookNode,ExcelAutomation::_WorksheetPtr pWorksheet)
/* ============================================================
	Function :		CExcelArchiveDocument::ImportWorksheet
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	Worksheet name to query for, or '*' for any

	Usage :			

   ============================================================*/
{
    try
        {
        CNode WorksheetNode=FindNode(_bstr_t(_T("Worksheets/Worksheet"))+(GetSourceWorksheet()==_bstr_t(_T("*"))?_bstr_t(_T("")):_bstr_t(_T("[@Name=\""))+(GetSourceWorksheet().length()?GetSourceWorksheet():pWorksheet->Name)+_bstr_t(_T("\"]"))),WorkbookNode);
        if (!WorksheetNode.IsEmpty())
            {
            return ImportRange(WorksheetNode);
            }
        else
            {
            LogMessage(L"Worksheet is empty");
            return false;
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::ImportWorksheet\': ")+e.ErrorMessage());
        return false;
        }
    return false;
}

bool CExcelArchiveDocument::ImportRange(CNode WorksheetNode)
/* ============================================================
	Function :		CExcelArchiveDocument::ImportRange
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	Range name to query for, or '*' for any

	Usage :			

   ============================================================*/
{
    try
        {
        CNode RangeNode=FindNode(_bstr_t(_T("Ranges/Range"))+(GetSourceRange()==_bstr_t(_T("*"))?_bstr_t(_T("")):_bstr_t(_T("[@Name=\""))+(GetSourceRange().length()?GetSourceRange():RangeName)+_bstr_t(_T("\"]"))),WorksheetNode);
        if (!RangeNode.IsEmpty())
            {
            long RowCount=GetRange()->Rows->Count;
            long ColumnCount=GetRange()->Columns->Count;
            CNodeList CellNodes=FindNodes(_bstr_t(_T("Cell")),RangeNode);
            CNode CellNode;
            while (!((CellNode=CellNodes.GetNext()).IsEmpty()))
                {
                long i=0,j=0;
                _variant_t Row=CellNode.GetAttribute(L"Row");
                _variant_t Column=CellNode.GetAttribute(L"Column");
                if ((Row==_variant_t())||(Column==_variant_t()))
                    {
                    _bstr_t Offset=CellNode.GetAttribute(L"Offset"); // Obsolete
                    _stscanf((LPCTSTR)Offset,_T("%d,%d"),&i,&j);
                    }
                else
                    {
                    i=(long)Row;
                    j=(long)Column;
                    }
                if (((i<RowCount)&&(j<ColumnCount))||!ConstrainedByTarget)
                    {
                    ImportCell(CellNode,GetRange()->Offset[i][j]->Resize[1L][1L]);
                    }
                }
            return true;
            }
        else
            {
            LogMessage(L"Range is empty");
            return false;
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::ImportRange\': ")+e.ErrorMessage());
        return false;
        }
    return false;
}

bool CExcelArchiveDocument::ImportCell(CNode CellNode,ExcelAutomation::RangePtr pCell)
/* ============================================================
	Function :		CExcelArchiveDocument::ImportCell
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	

	Usage :			

   ============================================================*/
{
    try
        {   
        if (CellNode.GetAttribute(L"Value")==_variant_t()) // If blank
            {
            pCell->ClearContents();
            }
        else
            {
            _bstr_t Type=CellNode.GetAttribute(L"Type");
            _variant_t Value;
            if (Type==_bstr_t(_T("Error")))
                {
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#DIV/0")))
                    Value=_variant_t((long)ExcelAutomation::xlErrDiv0,VT_ERROR);
                else
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#N/A")))
                    Value=_variant_t((long)ExcelAutomation::xlErrNA,VT_ERROR);
                else
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#NAME")))
                    Value=_variant_t((long)ExcelAutomation::xlErrName,VT_ERROR);
                else
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#NULL")))
                    Value=_variant_t((long)ExcelAutomation::xlErrNull,VT_ERROR);
                else
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#NUM")))
                    Value=_variant_t((long)ExcelAutomation::xlErrNum,VT_ERROR);
                else
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#REF")))
                    Value=_variant_t((long)ExcelAutomation::xlErrRef,VT_ERROR);
                else
                if ((_bstr_t)CellNode.GetAttribute(L"Value")==_bstr_t(_T("#VALUE")))
                    Value=_variant_t((long)ExcelAutomation::xlErrValue,VT_ERROR);
                }
            else
                if (Type==_bstr_t(_T("Number")))
                    Value=_variant_t(atof((_bstr_t)CellNode.GetAttribute(L"Value")));
                else
                    Value=_variant_t(CellNode.GetAttribute(L"Value"));
            pCell->Value2=Value;
            }
        return true;
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelArchiveDocument::ImportCell\' for \'")+pCell->Address[0L][0L][ExcelAutomation::xlA1][0L][0L]+_bstr_t(L"\': ")+e.ErrorMessage());
        return false;
        }
    return false;
}

