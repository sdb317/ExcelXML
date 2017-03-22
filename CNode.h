#ifndef _CNODE_H_B3D065BA_7D7C_49CC_AB90A2C9762F
#define _CNODE_H_B3D065BA_7D7C_49CC_AB90A2C9762F

///////////////////////////////////////////////////////////
// File :		CNode.h
// Created :	10/17/13
//

class CDocument;

class CNode : public MSXML::IXMLDOMNodePtr
{
public:
// Construction/destruction
	CNode(MSXML::IXMLDOMNodePtr NodePtr = NULL);
	CNode(CDocument Document,_bstr_t Name);
	virtual ~CNode();
    bool IsEmpty() {return (GetInterfacePtr()==NULL)?true:false;}

// Operations
	CNode AppendChild(CNode Child);
	CNode AppendChild(_bstr_t ChildName);
	CNode InsertChild(CNode Child,CNode Before);
	CNode RemoveChild(CNode Child);
	CNode ReplaceChild(CNode OldChild,CNode NewChild);
	_variant_t GetAttribute(_bstr_t Name) const;
	void SetAttribute(_bstr_t Name,_bstr_t Value);

// Attributes

};
#endif //_CNODE_H_B3D065BA_7D7C_49CC_AB90A2C9762F
