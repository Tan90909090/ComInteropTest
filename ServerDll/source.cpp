#include "../ComComponentRegisterer/ComServer.hpp"

#include <shlwapi.h>
#pragma comment(lib, "Shlwapi.lib") // for QISearch

#define COLOR_DEFAULT "\x1b[0m"
#define COLOR_RED "\x1b[31m"
#define COLOR_GREEN "\x1b[32m"
#define COLOR_BLUE "\x1b[34m"
#define COLOR_YELLOW "\x1b[33m"

struct CClassFactory : IClassFactory
{
	STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
	{
		static const QITAB rgqit[] = {
			QITABENT(CClassFactory, IClassFactory),
			{},
		};
		return QISearch(this, rgqit, riid, ppv);
	}
	STDMETHOD_(ULONG, AddRef)(void) override
	{
		return 1; // This object is a singleton.
	}
	STDMETHOD_(ULONG, Release)(void) override
	{
		return 1; // This object is a singleton.
	}
	STDMETHOD(CreateInstance)(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override
	{HRESULT hr;
		hr = CLASS_E_NOAGGREGATION;
		printf(COLOR_BLUE "[*] %s end, hr = %08x\n", __func__, hr);
		return hr;
	}
	STDMETHOD(LockServer)(BOOL fLock) override
	{
		printf(COLOR_RED "[*] %s, fLock=%d\n", __func__, fLock);
		return S_OK; // for simplicity
	}
};
CClassFactory g_classFactory;

// Export functions
// Anottations are copied from combaseapi.h
__control_entrypoint(DllExport) STDAPI DllCanUnloadNow()
{
	return S_FALSE; // for simplicity
}
_Check_return_ STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
	*ppv = NULL;

	return CLASS_E_CLASSNOTAVAILABLE;
}
