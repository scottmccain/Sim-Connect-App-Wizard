#pragma once
#include "stdafx.h"

#include "SimConnect.h"

#define DECLARE_DISPATCH_MAP virtual void DispatchProc(SIMCONNECT_RECV* pData, DWORD cbData);

#define BEGIN_DISPATCH_MAP(type) void type::DispatchProc(SIMCONNECT_RECV* pData, DWORD cbData) { \
	switch (pData->dwID) {

#define END_DISPATCH_MAP default: break; } }

#define ADD_DISPATCH_HANDLER(key, paramtype, handler) case key: \
	handler(reinterpret_cast<paramtype *>(pData), cbData); \
	break;

#define DECLARE_DISPATCH_HANDLER(function, type) void function(type *pData, DWORD cbData);

#define DECLARE_SIMCONNECT_APP(classname) classname *pApp;

#define DECLARE_SIMCONNECT_APPSTART(classname, appname) int __stdcall DLLStart(void) \
	{ \
	CoInitializeEx(NULL, COINIT_MULTITHREADED); \
	pApp = new classname(hDll, appname); \
if (pApp) \
	pApp->Connect(); \
	return 0; \
	}

#define DECLARE_SIMCONNECT_APPSTOP() void __stdcall DLLStop(void) \
	{ \
if (pApp != NULL) \
	delete pApp; \
	}

#define EVENT_MAP_BEGIN static enum EVENT_ID {
#define DECLARE_EVENT(id) id,
#define EVENT_MAP_END };

#define INPUT_MAP_BEGIN static enum INPUT_ID {
#define DECLARE_INPUT(id) id,
#define INPUT_MAP_END };

#define GROUP_MAP_BEGIN static enum GROUP_ID  {
#define DECLARE_GROUP(id) id,
#define GROUP_MAP_END };

#define REQUEST_MAP_BEGIN static enum REQUEST_ID {
#define DECLARE_REQUEST(id) id,
#define REQUEST_MAP_END };

// TODO: add more
#define DATA_TYPE_DEGREES "degrees"
#define DATA_TYPE_KNOTS "knots"
#define DATA_TYPE_BOOL "bool"
#define DATA_TYPE_FEET_PER_MINUTE "feet per minute"
#define DATA_TYPE_MASK "mask"
#define DATA_TYPE_ENUM "enum"

class CSimConnectApp
{
	static void CALLBACK DispatchProcStatic(SIMCONNECT_RECV* pData, DWORD cbData, void *pContext)
	{
		CSimConnectApp* pThis = reinterpret_cast<CSimConnectApp*>(pContext);
		pThis->DispatchProc(pData, cbData);
	}

protected:
	HANDLE m_hSimConnect;
	HINSTANCE m_hInstance;
	LPCSTR m_szAppName;

	virtual HRESULT OnConnect() = 0;
	virtual void DispatchProc(SIMCONNECT_RECV* pData, DWORD cbData) = 0;

public:
	CSimConnectApp(HINSTANCE, LPCSTR szAppName);
	virtual ~CSimConnectApp(void);

	HRESULT Connect();
};

