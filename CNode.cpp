/* ==========================================================================
	File :			CNode.cpp
	
	Class :			CNode

	Date :			10/17/13

	Purpose :		

	Description :	

	Usage :			

   ========================================================================*/

#include "stdafx.h"
#include "CNode.h"
#include "CDocument.h"

////////////////////////////////////////////////////////////////////
// Public functions
//

CNode::CNode(MSXML::IXMLDOMNodePtr NodePtr)
/* ============================================================
	Function :		CNode::CNode
	Description :	Constructor.
	Access :		Public
					
	Return :		void
	Parameters :	MSXML::IXMLDOMNodePtr NodePtr
	Usage :			

   ============================================================*/
{
    try
        {
        if (NodePtr!=NULL)
            Attach(NodePtr,true);
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
}

CNode::CNode(CDocument Document,_bstr_t Name)
/* ============================================================
	Function :		CNode::CNode
	Description :	Constructor.
	Access :		Public
					
	Return :		void
	Parameters :	CDocument Document,_bstr_t Name
	Usage :			

   ============================================================*/
{
    try
        {
        Attach(Document.GetInterfacePtr()->createElement(Name),true);
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
}

CNode::~CNode()
/* ============================================================
	Function :		CNode::~CNode
	Description :	Destructor.
	Access :		Public
					
	Return :		void
	Parameters :	none

	Usage :			

   ============================================================*/
{
	// TODO: Implement
}

CNode CNode::AppendChild(CNode Child)
/* ============================================================
	Function :		CNode::AppendChild
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	CNode Child
	Usage :			

   ============================================================*/
{
    try
        {
        return CNode(GetInterfacePtr()->appendChild(Child));
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return CNode(NULL);
}

CNode CNode::AppendChild(_bstr_t ChildName)
/* ============================================================
	Function :		CNode::AppendChild
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	_bstr_t ChildName
	Usage :			

   ============================================================*/
{
    try
        {
        CNode Element(GetInterfacePtr()->ownerDocument->createElement(ChildName));
        return AppendChild(Element);
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return CNode(NULL);
}

CNode CNode::InsertChild(CNode Child,CNode Before)
/* ============================================================
	Function :		CNode::InsertChild
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	CNode Child,CNode Before
	Usage :			

   ============================================================*/
{
    try
        {
        return CNode(GetInterfacePtr()->insertBefore(Child,_variant_t(Before)));
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return CNode(NULL);
}

CNode CNode::RemoveChild(CNode Child)
/* ============================================================
	Function :		CNode::RemoveChild
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	CNode Child
	Usage :			

   ============================================================*/
{
    try
        {
        return CNode(GetInterfacePtr()->removeChild(Child));
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return CNode(NULL);
}

CNode CNode::ReplaceChild(CNode OldChild,CNode NewChild)
/* ============================================================
	Function :		CNode::ReplaceChild
	Description :	
	Access :		Public
					
	Return :		CNode
	Parameters :	CNode OldChild,CNode NewChild
	Usage :			

   ============================================================*/
{
    try
        {
        return CNode(GetInterfacePtr()->replaceChild(NewChild,OldChild));
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return CNode(NULL);
}

_variant_t CNode::GetAttribute(_bstr_t Name) const
/* ============================================================
	Function :		CNode::GetAttribute
	Description :	Accessor. Getter for "Attribute".
	Access :		Public
					
	Return :		_variant_t
	Parameters :	_bstr_t Name
	Usage :			Call to get the value of "Attribute".

   ============================================================*/
{
    try
        {
        // return _variant_t(((MSXML::IXMLDOMAttributePtr)GetInterfacePtr()->selectSingleNode(_bstr_t(_T("@"))+Name))->value);
        return _variant_t(((MSXML::IXMLDOMAttributePtr)GetInterfacePtr()->attributes->getNamedItem(Name))->value);
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
    return _variant_t();
}

void CNode::SetAttribute(_bstr_t Name,_bstr_t Value)
/* ============================================================
	Function :		CNode::SetAttribute
	Description :	
	Access :		Public
					
	Return :		void
	Parameters :	_bstr_t Name,_bstr_t Value
	Usage :			

   ============================================================*/
{
    try
        {
        CNode Attribute(GetInterfacePtr()->ownerDocument->createAttribute(Name));
        ((MSXML::IXMLDOMAttributePtr)Attribute)->value=Value;
        GetInterfacePtr()->attributes->setNamedItem(Attribute);
        }
    catch (_com_error& e)
        {
        ATLTRACE(L"%s\n",(LPCTSTR)e.ErrorMessage());
        }
}

////////////////////////////////////////////////////////////////////
// Protected functions
//

////////////////////////////////////////////////////////////////////
// Private functions
//

