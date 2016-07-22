// TODO: add application stuff
// [!output SAFE_APPCLASS_NAME]

#include "SimConnectApp.h"

class [!output SAFE_APPCLASS_NAME] : public CSimConnectApp
{

protected:
	virtual HRESULT OnConnect();

	DECLARE_DISPATCH_MAP

public:
	[!output SAFE_APPCLASS_NAME](HINSTANCE, LPCSTR);
	virtual ~[!output SAFE_APPCLASS_NAME](void);
};