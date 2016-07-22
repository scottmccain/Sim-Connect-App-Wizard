// BaseDialog.cpp: implementation of the CBaseDialog class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DialogBase.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

std::map<HWND, CDialogBase *> CDialogBase::DialogMap;

CDialogBase::CDialogBase(HINSTANCE hInstance, int nResId, HWND hParent) : m_hModuleInstance(hInstance), m_szDialogResource(MAKEINTRESOURCE(nResId)), m_hwndParent(hParent)
{
}

CDialogBase::~CDialogBase()
{
}

int CDialogBase::DoModal()
{
	return(DialogBoxParam(m_hModuleInstance, m_szDialogResource, m_hwndParent, DialogProcStatic, reinterpret_cast<LPARAM>(this)));
}

BOOL CALLBACK CDialogBase::DialogProcStatic(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_INITDIALOG)
	{
		CDialogBase *pThis = reinterpret_cast<CDialogBase *>(lParam);
		DialogMap[hDlg] = pThis;

		if (pThis)
			pThis->InitializeDialog(hDlg);
	}

	if (DialogMap.find(hDlg) != DialogMap.end())
	{
		return(DialogMap[hDlg]->DialogProc(hDlg, message, wParam, lParam));
	}

	return FALSE;
}

BOOL CDialogBase::DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	return FALSE;
}

void CDialogBase::OnOK(void)
{
	EndDialog(m_hwndDialog, IDOK);
}

void CDialogBase::OnCancel(void)
{
	EndDialog(m_hwndDialog, IDCANCEL);
}

