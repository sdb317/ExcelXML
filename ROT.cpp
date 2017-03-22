#include "stdafx.h"

ExcelAutomation::_ApplicationPtr GetExcelApplicationPtr(HWND hExcelWnd)
{
    LogMessage(_bstr_t(_T("ROT::GetExcelApplicationPtr")));
    LogMessage(_bstr_t(_T("hExcelWnd: "))+_bstr_t((long)hExcelWnd));
	ExcelAutomation::_ApplicationPtr ApplicationPtr=NULL;
    IBindCtx *pBindCtx;
    HRESULT hr;
    hr=CreateBindCtx(0,&pBindCtx); // Get a BindCtx.
    if (FAILED(hr)) 
        {
        LogMessage(_bstr_t(_T("CreateBindCtx failed")));
        return ApplicationPtr;
        }
    IRunningObjectTable *pRunningObjectTable;
    hr=pBindCtx->GetRunningObjectTable(&pRunningObjectTable); // Get running-object table.
    if (FAILED(hr)) 
        {
        LogMessage(_bstr_t(_T("GetRunningObjectTable failed")));
        pBindCtx->Release();
        return ApplicationPtr;
        }
    IEnumMoniker *pEnumMoniker;
    hr=pRunningObjectTable->EnumRunning(&pEnumMoniker); // Get enumeration interface.
    if (FAILED(hr)) 
        {
        LogMessage(_bstr_t(_T("EnumRunning failed")));
        pRunningObjectTable->Release();
        pBindCtx->Release();
        return ApplicationPtr;
        }
    pEnumMoniker->Reset(); // Start at the beginning.
    ULONG fetched;
    IMoniker *pMoniker;
    int n=0;
    while (pEnumMoniker->Next(1,&pMoniker,&fetched)==S_OK) // Churn through enumeration.
        {
        try
            {
            DWORD Type=NULL; // Pointer to an integer from the MKSYS enumeration 
            LPOLESTR pName=NULL;
            CLSID Clsid;  //Pointer to an object's CLSID
            ZeroMemory(&Clsid,sizeof(Clsid));
            hr=pMoniker->IsSystemMoniker(&Type); // Get the moniker class. We're only interested in file monikers for the Workbook object, although the Excel app is an item moniker.
            if (SUCCEEDED(hr))
                {
                hr=pMoniker->GetDisplayName(pBindCtx,NULL,&pName); // Get display name.
/*
                if (SUCCEEDED(hr))
                    {
                    hr=pMoniker->GetClassID(&Clsid);
                    }
*/
                }
            _bstr_t Name(pName);
            LogMessage(_bstr_t(_T(" Name: "))+Name);
            if 
                ( // Don't try to bind to everything just possible workbooks
                    (SUCCEEDED(hr))
                    &&
//                    ((Type==MKSYS_FILEMONIKER)||(Type==MKSYS_URLMONIKER))
//                    &&
                    (_tcsstr(_tcslwr((wchar_t *)(LPCTSTR)Name),_T(".xl")))
                )
                {
                IDispatch* pDispatch=NULL;
                hr=pMoniker->BindToObject(pBindCtx,NULL,IID_IDispatch,(void**)&pDispatch);
                if (SUCCEEDED(hr))
                    {
                    ExcelAutomation::_WorkbookPtr pWorkbook=NULL;
                    pDispatch->QueryInterface(__uuidof(ExcelAutomation::_Workbook),(void**)&pWorkbook);
                    pDispatch->Release();
                    pDispatch=NULL;
                    if (pWorkbook!=NULL)
                        {
                        ExcelAutomation::_ApplicationPtr pApplication=pWorkbook->GetApplication();
                        pWorkbook=NULL;
                        if (pApplication!=NULL)
                            {
                            HWND hWnd=(HWND)pApplication->GetHwnd();
                            LogMessage(_bstr_t(_T("  hWnd: "))+_bstr_t((long)hWnd));
                            if (hWnd==hExcelWnd) // It must be us!!!
                                {
                                pMoniker->Release(); // Release interfaces.
                                ApplicationPtr=pApplication;
                                pApplication=NULL;
                                break;
                                }
/*
                            HINSTANCE hInstance=(HINSTANCE)pApplication->Hinstance;
                            if (hInstance==hExcelInstance)
                                {
                                pMoniker->Release(); // Release interfaces.
                                ApplicationPtr=pApplication;
                                pApplication=NULL;
                                break;
                                }
*/
                            pApplication=NULL;
                            }
                        else
                            {
                            LogMessage(_bstr_t(_T("Failed to retrieve ExcelAutomation::_Application")));
                            }
                        }
                    else
                        {
                        LogMessage(_bstr_t(_T("Failed to retrieve ExcelAutomation::_Workbook")));
                        }
                    }
                else
                    {
                    LogMessage(_bstr_t(_T("BindToObject failed")));
                    }
                }
            }
        catch (_com_error& e)
            {
            LogMessage(_bstr_t(_T("ROT::GetExcelApplicationPtr error: "))+e.ErrorMessage());
            }
/*
        LPOLESTR pName;
        pMoniker->GetDisplayName(pBindCtx,NULL,&pName); // Get DisplayName.
        char szName[512];
        WideCharToMultiByte(CP_ACP,0,pName,-1,szName,512,NULL,NULL); // Convert it to ASCII.
        if (!strcmp(szName,m_szDocName)) 
            {
            IDispatch *pDisp;
            hr=pMoniker->BindToObject(pBindCtx,NULL,IID_IDispatch,(void**)&pDisp); // Bind to this ROT entry.
            if (!FAILED(hr)) 
                {
                pDispatch=pDisp;
                }
            }
        pMoniker->Release(); // Release interfaces.
        if (pDispatch!=NULL) break; // Break out if we obtained the IDispatch successfully.
*/
        pMoniker->Release(); // Release interfaces.
        }
    pEnumMoniker->Release(); // Release interfaces.
    pRunningObjectTable->Release();
    pBindCtx->Release();
    return ApplicationPtr;
}

