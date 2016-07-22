[!if INCLUDE_STDAFX]
#include "stdafx.h"
[!endif]

#include "[!output APPLICATION_HEADER]"

#include <objbase.h>

[!output SAFE_APPCLASS_NAME] * pApp;

// TODO: move into macros
int __stdcall DLLStart(void)
{
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	pApp = new [!output SAFE_APPCLASS_NAME](hDll, "[!output APPLICATION_NAME]");

	if (pApp != NULL)
		pApp->Connect();

	return 0;
}

//
// The DLLStop function must be present.
//
void __stdcall DLLStop(void)
{
	if (pApp != NULL)
		delete pApp;
}