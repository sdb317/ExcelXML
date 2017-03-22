/* ==========================================================================
	File :			CDocument.cpp
	
	Class :			CDocument

	Date :			10/17/13

	Purpose :		

	Description :	

	Usage :			

   ========================================================================*/

#include "stdafx.h"
#include "CDocument.h"

////////////////////////////////////////////////////////////////////
// Public functions
//

CDocument::CDocument()
/* ============================================================
	Function :		CDocument::CDocument
	Description :	Constructor.
	Access :		Public
					
	Return :		void
	Parameters :	MSXML::IXMLDOMDocumentPtr DocumentPtr
	Usage :			

   ============================================================*/
{
    try
        {
        CreateInstance(__uuidof(MSXML::DOMDocument60));
        GetInterfacePtr()->async=VARIANT_FALSE;
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CDocument::CDocument\': ")+e.ErrorMessage());
        }
}

CDocument::~CDocument()
/* ============================================================
	Function :		CDocument::~CDocument
	Description :	Destructor.
	Access :		Public
					
	Return :		void
	Parameters :	none

	Usage :			

   ============================================================*/
{
	// TODO: Implement
}

CNodeList CDocument::FindNodes(_bstr_t Query,CNode StartingFrom)
/* ============================================================
	Function :		CDocument::FindNodes
	Description :	
	Access :		Public
					
	Return :		CNodeList
	Parameters :	_bstr_t Query,CNode StartingFrom
	Usage :			

   ============================================================*/
{
    try
        {
#ifdef _DEBUG
        LogMessage(_bstr_t(L"Query to \'CDocument::FindNodes\': ")+Query);
#endif
	    if (StartingFrom.GetInterfacePtr()==NULL)
            {
            return CNodeList(GetInterfacePtr()->documentElement->selectNodes(Query));
            }
	    else
            {
            return CNodeList(StartingFrom.GetInterfacePtr()->selectNodes(Query));
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CDocument::FindNodes\': ")+e.ErrorMessage());
        }
    return CNodeList(NULL);
}

CNode CDocument::FindNode(_bstr_t Query,CNode StartingFrom)
/* ============================================================
	Function :		CDocument::FindNode
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	_bstr_t Query,CNode StartingFrom
	Usage :			

   ============================================================*/
{
    try
        {
#ifdef _DEBUG
        LogMessage(_bstr_t(L"Query to \'CDocument::FindNode\': ")+Query);
#endif
	    if (StartingFrom.GetInterfacePtr()==NULL)
            {
            return CNode(GetInterfacePtr()->documentElement->selectSingleNode(Query));
            }
	    else
            {
            return CNode(StartingFrom.GetInterfacePtr()->selectSingleNode(Query));
            }
        }
    catch (_com_error& e)
        {
        LogMessage(_bstr_t(L"Error in \'CDocument::FindNode\': ")+e.ErrorMessage());
        }
    return CNode(NULL);
}

////////////////////////////////////////////////////////////////////
// Protected functions
//

////////////////////////////////////////////////////////////////////
// Private functions
//

