// ComComponentRegisterer.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "ComServer.hpp"

#include <shlwapi.h>
#pragma comment(lib, "Shlwapi.lib") // for QISearch


#define COLOR_DEFAULT "\x1b[0m"
#define COLOR_RED "\x1b[31m"
#define COLOR_GREEN "\x1b[32m"
#define COLOR_BLUE "\x1b[34m"
#define COLOR_YELLOW "\x1b[33m"

struct CDummyNamespaceWalkCB : INamespaceWalkCB2
{
	CDummyNamespaceWalkCB()
	{
		puts(COLOR_RED "[!] CDummyNamespaceWalkCB constructor!");
	}
	CDummyNamespaceWalkCB(CDummyNamespaceWalkCB&) = delete;
	CDummyNamespaceWalkCB& operator = (CDummyNamespaceWalkCB&) = delete;

	STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
	{
		RPC_CSTR pszGuid{};
		std::ignore = UuidToStringA(&riid, &pszGuid);
		printf(COLOR_GREEN "[*] CDummyNamespaceWalkCB %s start, GUID=%s\n", __func__, pszGuid);
		RpcStringFreeA(&pszGuid);

		static const QITAB rgqit[] = {
			QITABENT(CDummyNamespaceWalkCB, INamespaceWalkCB),
			QITABENT(CDummyNamespaceWalkCB, INamespaceWalkCB2),
			{},
		};
		puts(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB QISearch start");
		int hr = QISearch(this, rgqit, riid, ppv);
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB QISearch end, hr = %08x\n", hr);
		printf(COLOR_GREEN "[*] CDummyNamespaceWalkCB %s end, hr=%08x\n", __func__, hr);
		return hr;
	}
	STDMETHOD_(ULONG, AddRef)(void) override
	{
		auto result = InterlockedIncrement(&_refCount);
		printf(COLOR_YELLOW "[+] CDummyNamespaceWalkCB %s, RefCount=%d\n", __func__, result);
		return result;
	}
	STDMETHOD_(ULONG, Release)(void) override
	{
		ULONG result = InterlockedDecrement(&_refCount);
		printf(COLOR_YELLOW "[+] CDummyNamespaceWalkCB %s, RefCount=%d\n", __func__, result);
		if (result == 0)
		{
			printf(COLOR_RED "[!] CDummyNamespaceWalkCB deleted!\n");
			delete this;
		}
		return result;
	}
	STDMETHOD(FoundItem)(IShellFolder* psf, PCUITEMID_CHILD pidl) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(EnterFolder)(IShellFolder* psf, PCUITEMID_CHILD pidl) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(LeaveFolder)(IShellFolder* psf, PCUITEMID_CHILD pidl) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(InitializeProgressDialog)(LPWSTR* ppszTitle, LPWSTR* ppszCancel) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(WalkComplete)(HRESULT hr) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
private:
	ULONG _refCount = 1;
};

struct CDummyNamespaceWalk : INamespaceWalk
{
	CDummyNamespaceWalk()
	{
		puts(COLOR_RED "[!] CDummyNamespaceWalk constructor!");
	}
	CDummyNamespaceWalk(CDummyNamespaceWalk&) = delete;
	CDummyNamespaceWalk& operator = (CDummyNamespaceWalk&) = delete;

	STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
	{
		RPC_CSTR pszGuid{};
		std::ignore = UuidToStringA(&riid, &pszGuid);
		printf(COLOR_GREEN "[*] CDummyNamespaceWalk %s start, GUID=%s\n", __func__, pszGuid);
		RpcStringFreeA(&pszGuid);

		static const QITAB rgqit[] = {
			QITABENT(CDummyNamespaceWalk, INamespaceWalk),
			{},
		};
		puts(COLOR_DEFAULT "[*] CDummyNamespaceWalk QISearch start");
		int hr = QISearch(this, rgqit, riid, ppv);
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalk QISearch end, hr = %08x\n", hr);
		printf(COLOR_GREEN "[*] CDummyNamespaceWalk %s end, hr=%08x\n", __func__, hr);
		return hr;
	}
	STDMETHOD_(ULONG, AddRef)(void) override
	{
		auto result = InterlockedIncrement(&_refCount);
		printf(COLOR_YELLOW "[+] CDummyNamespaceWalk %s, RefCount=%d\n", __func__, result);
		return result;
	}
	STDMETHOD_(ULONG, Release)(void) override
	{
		ULONG result = InterlockedDecrement(&_refCount);
		printf(COLOR_YELLOW "[+] CDummyNamespaceWalk %s, RefCount=%d\n", __func__, result);
		if (result == 0)
		{
			printf(COLOR_RED "[!] CDummyNamespaceWalk deleted!\n");
			delete this;
		}
		return result;
	}
	HRESULT __stdcall Walk(IUnknown* punkToWalk, DWORD dwFlags, int cDepth, INamespaceWalkCB* pnswcb) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalk %s start, call pnswcb\n", __func__);
		LPWSTR pszTitle = nullptr, pszCancel = nullptr;
		pnswcb->InitializeProgressDialog(&pszTitle, &pszCancel);
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalk %s end\n", __func__);
		return S_OK;
	}
	HRESULT __stdcall GetIDArrayResult(UINT* pcItems, PIDLIST_ABSOLUTE** prgpidl) override
	{
		printf(COLOR_DEFAULT "[*] CDummyNamespaceWalk %s\n", __func__);
		return S_OK;
	}
private:
	ULONG _refCount = 1;
};

template<class TInstance>
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
	{
		RPC_CSTR pszGuid{};
		std::ignore = UuidToStringA(&riid, &pszGuid);
		printf(COLOR_BLUE "[*] %s start, GUID=%s\n", __func__, pszGuid);
		RpcStringFreeA(&pszGuid);

		HRESULT hr;
		if (pUnkOuter != nullptr)
		{
			hr = CLASS_E_NOAGGREGATION;
		}
		else
		{
			auto pInstance = new TInstance();
			hr = pInstance->QueryInterface(riid, ppvObject);
			pInstance->Release();
		}

		printf(COLOR_BLUE "[*] %s end, hr = %08x\n", __func__, hr);
		return hr;
	}
	STDMETHOD(LockServer)(BOOL fLock) override
	{
		// 大丈夫、呼び出されていない。
		printf(COLOR_RED "[*] %s, fLock=%d\n", __func__, fLock);
		return S_OK; // I dont know that behavior is expected...
	}
};

int main()
{
	CClassFactory<CDummyNamespaceWalk> classFactoryDummyNamespaceWalk;
	CClassFactory<CDummyNamespaceWalkCB> classFactoryDummyNamespaceWalkCB;

	// WindowsTerminalの都合で新規ターミナルを作りたくなったんだけれども
	// AllocConsoleで作ったコンソールをstdioに関連付けるのがどうあがいても無理だった
	// ↓どちらもAccessViolationで死ぬ。msvcrt.dllのは古いやつだから？
	// じゃあWindowsTerminalの設定変えればよかったや


	int hr = CoInitialize(nullptr);
	if (FAILED(hr)) { throw hr; }

	// https://ichigopack.net/win32com/com_exeserver_2.html#:~:text=%E3%81%AA%E3%82%8B%E3%81%A7%E3%81%97%E3%82%87%E3%81%86%E3%80%82-,%E3%83%A1%E3%82%BD%E3%83%83%E3%83%89%E4%BB%B2%E4%BB%8B%E3%81%AE%E6%96%B9%E6%B3%95,-%E5%85%88%E3%81%AEproxy
	// LOCAL_SERVERなEXEの場合、自作interfaceをプロセス間通信させるためにproxy/stubが必要だぜ！面倒くさいぜ！企画倒れだぜ！INPROC_SERVERなDLLにしますね！
	// proxy/stub実装はINPROCなDLLが必要らしい！終わり終わり！ https://ichigopack.net/win32com/com_proxystub_1.html
	DWORD dwRegisterFoo{};
	DWORD dwRegisterDummyNamespaceWalk{};
	DWORD dwRegisterDummyNamespaceWalkCB{};
	auto dwClsContext = CLSCTX_LOCAL_SERVER;
	auto flags = REGCLS_MULTIPLEUSE | REGCLS_SUSPENDED;

	hr = CoRegisterClassObject(CLSID_DummyNamespaceWalk, &classFactoryDummyNamespaceWalk, dwClsContext, flags, &dwRegisterDummyNamespaceWalk);
	if (FAILED(hr)) { throw hr; }

	hr = CoRegisterClassObject(CLSID_DummyNamespaceWalkCB, &classFactoryDummyNamespaceWalkCB, dwClsContext, flags, &dwRegisterDummyNamespaceWalkCB);
	if (FAILED(hr)) { throw hr; }

	hr = CoResumeClassObjects();
	if (FAILED(hr)) { throw hr; }

	puts(COLOR_GREEN "Running...");
	// getchar待ちだと止まったのでメッセージループが必要らしい。
	MessageBoxW(nullptr, L"Com Server is running. Close this to unregister.", L"Title", MB_OK);

	hr = CoRevokeClassObject(dwRegisterDummyNamespaceWalkCB);
	if (FAILED(hr)) { throw hr; }
	hr = CoRevokeClassObject(dwRegisterDummyNamespaceWalk);
	if (FAILED(hr)) { throw hr; }
	hr = CoRevokeClassObject(dwRegisterFoo);
	if (FAILED(hr)) { throw hr; }
	CoUninitialize();
}