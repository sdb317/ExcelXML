/* ==========================================================================
	File :			CNodeList.cpp
	
	Class :			CNodeList

	Date :			10/17/13

	Purpose :		

	Description :	

	Usage :			

   ========================================================================*/

#include "stdafx.h"
#include "CNodeList.h"

////////////////////////////////////////////////////////////////////
// Public functions
//

CNodeList::CNodeList(MSXML::IXMLDOMNodeListPtr NodeListPtr)
/* ============================================================
	Function :		CNodeList::CNodeList
	Description :	Constructor.
	Access :		Public
					
	Return :		void
	Parameters :	MSXML::IXMLDOMNodeListPtr NodeListPtr
	Usage :			

   ============================================================*/
{
    try
        {
        if (NodeListPtr!=NULL)
            Attach(NodeListPtr,true);
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
}

CNodeList::~CNodeList()
/* ============================================================
	Function :		CNodeList::~CNodeList
	Description :	Destructor.
	Access :		Public
					
	Return :		void
	Parameters :	none

	Usage :			

   ============================================================*/
{
	// TODO: Implement
}

CNode CNodeList::GetNext()
/* ============================================================
	Function :		CNodeList::GetNext
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	none

	Usage :			

   ============================================================*/
{
    try
        {
	    return CNode(GetInterfacePtr()->nextNode());
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return CNode(NULL);
}

////////////////////////////////////////////////////////////////////
// Protected functions
//

////////////////////////////////////////////////////////////////////
// Private functions
//

