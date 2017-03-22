#ifndef _CDOCUMENT_H_4187291C_C154_4A5B_86E1D848BD40
#define _CDOCUMENT_H_4187291C_C154_4A5B_86E1D848BD40

///////////////////////////////////////////////////////////
// File :		CDocument.h
// Created :	10/17/13
//

#include "CNode.h"
#include "CNodeList.h"

class CDocument : public MSXML::IXMLDOMDocumentPtr
{
public:
// Construction/destruction
	CDocument();
	virtual ~CDocument();

// Operations
	CNodeList FindNodes(_bstr_t Query,CNode StartingFrom=CNode((MSXML::IXMLDOMNodePtr)NULL));
	CNode FindNode(_bstr_t Query,CNode StartingFrom=CNode((MSXML::IXMLDOMNodePtr)NULL));

// Attributes

};
#endif //_CDOCUMENT_H_4187291C_C154_4A5B_86E1D848BD40
