
[!if INCLUDE_STDAFX]
#include "stdafx.h"
[!endif]

#include "[!output APPLICATION_HEADER]"


BEGIN_DISPATCH_MAP([!output SAFE_APPCLASS_NAME])
END_DISPATCH_MAP

[!output SAFE_APPCLASS_NAME]::[!output SAFE_APPCLASS_NAME](HINSTANCE hInstance, LPCSTR szAppName) : CSimConnectApp(hInstance, szAppName)
{
}

[!output SAFE_APPCLASS_NAME]::~[!output SAFE_APPCLASS_NAME]()
{

}

HRESULT [!output SAFE_APPCLASS_NAME]::OnConnect()
{
	HRESULT hr = S_OK;

	// TODO: Add Initialization code here

	return hr;
}