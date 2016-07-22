#include "stdafx.h"
#include "SimConnectApp.h"

CSimConnectApp::CSimConnectApp(HINSTANCE hInstance, LPCSTR szAppName) : m_hInstance(hInstance), m_szAppName(szAppName), m_hSimConnect(NULL)
{
}

CSimConnectApp::~CSimConnectApp(void)
{
	if (m_hSimConnect != NULL)
		SimConnect_Close(m_hSimConnect);
}

HRESULT CSimConnectApp::Connect()
{
	// if we are already connected then just return
	if (m_hSimConnect != NULL) return S_OK;

	HRESULT hr;
	if (SUCCEEDED(hr = SimConnect_Open(&m_hSimConnect, m_szAppName, NULL, 0, NULL, 0)))
	{
		OnConnect();

		// hookup our pump, all ininitializtion should be done at this point
		return SimConnect_CallDispatch(m_hSimConnect, CSimConnectApp::DispatchProcStatic, this);
	}

	return hr;
}
