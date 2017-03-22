// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"

#include <ctime>
#include <fstream>

static std::ofstream LogFile;
_bstr_t LastError=L"";

void LogMessage(_bstr_t Message, bool Popup)
{
    ATLTRACE("%s\n",(LPCTSTR)Message);
    LastError=Message;
    if (LogFile.is_open())
        {
        time_t StandardTime;
        time(&StandardTime);
        struct tm* pLocalTime;
        pLocalTime=localtime(&StandardTime);
        TCHAR LogTime[256];
        ZeroMemory(LogTime,sizeof(LogTime));
        _tcsftime(LogTime,sizeof(LogTime),_T("%Y-%m-%d %H:%M:%S - "),pLocalTime);
        LogFile << (LPCSTR)(_bstr_t(LogTime)+Message) << std::endl;
        }
    if (Popup)
        MessageBox(NULL,(LPCTSTR)Message,L"ExcelXML",MB_OK|MB_ICONEXCLAMATION);
/*
    else
        if (pApplication!=NULL)
            pApplication->StatusBar[0]=Message;
*/
}

struct SCoInitialize
    {
    SCoInitialize() {CoInitialize(NULL);}
    ~SCoInitialize() {CoUninitialize();}
    };

SCoInitialize InitializeCOM;

extern ExcelAutomation::_ApplicationPtr pApplication;

BOOL APIENTRY DllMain
    ( 
    HMODULE hModule,
    DWORD  ul_reason_for_call,
    LPVOID lpReserved
    )
{
	switch (ul_reason_for_call)
	    {
	    case DLL_PROCESS_ATTACH:
	        {
            TCHAR LogFileFolder[256];
            ZeroMemory(LogFileFolder,sizeof(LogFileFolder));
            if (!ExpandEnvironmentStrings(_T("%LOCALAPPDATA%"),LogFileFolder,sizeof(LogFileFolder)-1)) // W7
                {
                if (!ExpandEnvironmentStrings(_T("%TEMP%"),LogFileFolder,sizeof(LogFileFolder)-1)) // XP
                    {
                    break; // What OS is it then?
                    }
                }
            else
                {
                _tcscat(LogFileFolder,_T("\\Temp"));
                }
            DWORD ProcessId=GetCurrentProcessId();
            TCHAR PID[64];
            ZeroMemory(PID,sizeof(PID));
            _itot(ProcessId,PID,10);
            _bstr_t LogFileName=_bstr_t(LogFileFolder)+_bstr_t(_T("\\ExcelXML_"))+_bstr_t(PID)+_bstr_t(_T(".log"));
            LogFile.open((LPCTSTR)LogFileName,std::ios_base::trunc);
            LogMessage(L"ExcelXML initialised");
            HWND CurrentWnd=::GetForegroundWindow(); // Need this to ensure Excel registers itself in the ROT
            if (::SetForegroundWindow(GetDesktopWindow()))
                ::SetForegroundWindow(CurrentWnd);
            HWND hExcelWnd=0L;
            XLOPER12 x;
	        if (Excel12f(xlGetHwnd,&x,0)==xlretSuccess)
                {
                hExcelWnd=(HWND)x.val.w;
                }
            try
                {
	            pApplication=GetExcelApplicationPtr(hExcelWnd);
                if (pApplication!=NULL)
                    LogMessage(L"Retrieved Excel application instance");
                else
                    LogMessage(L"Failed to retrieve Excel application instance");
                }
            catch (_com_error& e)
                {
                LogMessage(e.ErrorMessage());
                }
            break;
            }
	    case DLL_THREAD_ATTACH:
	    case DLL_THREAD_DETACH:
	    case DLL_PROCESS_DETACH:
		    break;
	    }
	return TRUE;
}

