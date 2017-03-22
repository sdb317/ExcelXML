#ifndef _CNODELIST_H_BAE4CBBF_D428_4E4E_A3EA798858
#define _CNODELIST_H_BAE4CBBF_D428_4E4E_A3EA798858

///////////////////////////////////////////////////////////
// File :		CNodeList.h
// Created :	10/17/13
//

#include "CNode.h"

class CNodeList : public MSXML::IXMLDOMNodeListPtr
{
public:
// Construction/destruction
	CNodeList(MSXML::IXMLDOMNodeListPtr NodeListPtr = NULL);
	virtual ~CNodeList();

// Operations
	CNode GetNext();

// Attributes

};
#endif //_CNODELIST_H_BAE4CBBF_D428_4E4E_A3EA798858
