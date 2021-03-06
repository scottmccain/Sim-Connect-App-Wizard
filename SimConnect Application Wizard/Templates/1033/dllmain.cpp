// dllmain.cpp : Defines the entry point for the DLL application.
[!if INCLUDE_STDAFX]
#include "stdafx.h"
[!endif]

HINSTANCE hDll;

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
	)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		hDll = (HINSTANCE)hModule;
		break;

	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

