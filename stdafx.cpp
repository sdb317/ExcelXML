#include "stdafx.h"
#include "time.h"

_bstr_t GetToday()
{
    TCHAR Today[10];
    ZeroMemory(Today,sizeof(Today));
    time_t CurrentTime=time(NULL);
    _tcsftime(Today,sizeof(Today),L"%Y%m%d",localtime(&CurrentTime));
    return _bstr_t(Today);
}

_bstr_t GetNow()
{
    TCHAR Now[8];
    ZeroMemory(Now,sizeof(Now));
    time_t CurrentTime=time(NULL);
    _tcsftime(Now,sizeof(Now),L"%H%M%S",localtime(&CurrentTime));
    return _bstr_t(Now);
}

