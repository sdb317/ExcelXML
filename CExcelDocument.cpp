/* ==========================================================================
	File :			CExcelDocument.cpp
	
	Class :			CExcelDocument

	Date :			10/22/13

	Purpose :		

	Description :	

	Usage :			

   ========================================================================*/

#include "stdafx.h"
#include "CExcelDocument.h"

////////////////////////////////////////////////////////////////////
// Public functions
//

_bstr_t CExcelDocument::GetDocumentName() const
/* ============================================================
	Function :		CExcelDocument::GetDocumentName
	Description :	Accessor. Getter for "DocumentName".
	Access :		Public
					
	Return :		_bstr_t
	Parameters :	none

	Usage :			Call to get the value of "DocumentName".

   ============================================================*/
{
	return DocumentName;
}

bool CExcelDocument::SetDocumentName(_bstr_t value)
/* ============================================================
	Function :		CExcelDocument::SetDocumentName
	Description :	Accessor. Setter for "DocumentName".
	Access :		Public
					
	Return :		bool
	Parameters :	_bstr_t value
	Usage :			Call to set the value of "DocumentName".

   ============================================================*/
{
	DocumentName=value;
    try
        {
        if (!(GetInterfacePtr()->load(DocumentName))) // Load an existing document
            {
            if (!(GetInterfacePtr()->loadXML(GetRootElement()))) // Create a new document
                {
                return false; // No point continuing
                }
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CExcelDocument::SetDocumentName\': ")+e.ErrorMessage());
        return false;
        }
    return true;
}

bool CExcelDocument::Export()
/* ============================================================
	Function :		CExcelDocument::Export
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	none

	Usage :			

   ============================================================*/
{
	return true;
}

bool CExcelDocument::Import()
/* ============================================================
	Function :		CExcelDocument::Import
	Description :	
	Access :		Public
					
	Return :		bool
	Parameters :	none

	Usage :			

   ============================================================*/
{
	return true;
}

////////////////////////////////////////////////////////////////////
// Protected functions
//


////////////////////////////////////////////////////////////////////
// Private functions
//


