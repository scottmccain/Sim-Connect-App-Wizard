#pragma once

#include <map>

/////////////////////////////// MACROS

#define DECLARE_MESSAGE_MAP virtual BOOL DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
#define DECLARE_MESSAGE_HANDLER(handler) BOOL handler(HWND hDlg, WPARAM wParam, LPARAM lParam);

#define BEGIN_MESSAGE_MAP(classname) BOOL classname::DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {\
	switch (message) {

#define ADD_MESSAGE_HANDLER(message, func) case message: return func(hDlg, wParam, lParam); break;

#define END_MESSAGE_MAP default: break; } return FALSE; }

class CDialogBase
{
	static BOOL CALLBACK DialogProcStatic(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
	static std::map<HWND, CDialogBase*> DialogMap;

public:

	CDialogBase(HINSTANCE hInstance, int nResId, HWND hParent = NULL);
	virtual ~CDialogBase();

	int DoModal();

	void OnOK(void);
	void OnCancel(void);

protected:

	HINSTANCE m_hModuleInstance;
	HWND m_hwndParent;
	HWND m_hwndDialog;
	LPCTSTR m_szDialogResource;

	virtual BOOL DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) = 0;

	void InitializeDialog(HWND hDlg)
	{
		this->m_hwndDialog = hDlg;
	}
};



