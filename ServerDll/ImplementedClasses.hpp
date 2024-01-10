#pragma once

#include "ClsIds.hpp"

#include <tuple> // for std::ignore
#include <shlwapi.h>
#pragma comment(lib, "Shlwapi.lib") // for QISearch

#define USE_ESCAPE_CHARACTER_TO_COLOR_OUTPUT 1
#if USE_ESCAPE_CHARACTER_TO_COLOR_OUTPUT
	#define CONSOLE_COLOR_DEFAULT "\x1b[0m"
	#define CONSOLE_COLOR_RED "\x1b[31m"
	#define CONSOLE_COLOR_GREEN "\x1b[32m"
	#define CONSOLE_COLOR_BLUE "\x1b[34m"
	#define CONSOLE_COLOR_YELLOW "\x1b[33m"
#else
	#define CONSOLE_COLOR_DEFAULT
	#define CONSOLE_COLOR_RED
	#define CONSOLE_COLOR_GREEN
	#define CONSOLE_COLOR_BLUE
	#define CONSOLE_COLOR_YELLOW
#endif

struct CDummyNamespaceWalkCB : INamespaceWalkCB2
{
	CDummyNamespaceWalkCB()
	{
		puts(CONSOLE_COLOR_RED "[!] CDummyNamespaceWalkCB constructor!" CONSOLE_COLOR_DEFAULT);
	}
	CDummyNamespaceWalkCB(CDummyNamespaceWalkCB&) = delete;
	CDummyNamespaceWalkCB& operator = (CDummyNamespaceWalkCB&) = delete;

	STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
	{
		RPC_CSTR pszGuid{};
		std::ignore = UuidToStringA(&riid, &pszGuid);
		printf(CONSOLE_COLOR_GREEN "[*] CDummyNamespaceWalkCB %s start, GUID=%s\n" CONSOLE_COLOR_DEFAULT, __func__, pszGuid);
		RpcStringFreeA(&pszGuid);

		static const QITAB rgqit[] = {
			QITABENT(CDummyNamespaceWalkCB, INamespaceWalkCB),
			QITABENT(CDummyNamespaceWalkCB, INamespaceWalkCB2),
			{},
		};
		puts("[*] CDummyNamespaceWalkCB QISearch start");
		int hr = QISearch(this, rgqit, riid, ppv);
		printf("[*] CDummyNamespaceWalkCB QISearch end, hr = %08x\n", hr);
		printf(CONSOLE_COLOR_GREEN "[*] CDummyNamespaceWalkCB %s end, hr=%08x\n" CONSOLE_COLOR_DEFAULT, __func__, hr);
		return hr;
	}
	STDMETHOD_(ULONG, AddRef)() override
	{
		auto result = InterlockedIncrement(&_refCount);
		printf(CONSOLE_COLOR_YELLOW "[+] CDummyNamespaceWalkCB %s, RefCount=%d\n", __func__, result);
		return result;
	}
	STDMETHOD_(ULONG, Release)() override
	{
		ULONG result = InterlockedDecrement(&_refCount);
		printf(CONSOLE_COLOR_YELLOW "[+] CDummyNamespaceWalkCB %s, RefCount=%d\n" CONSOLE_COLOR_DEFAULT, __func__, result);
		if (result == 0)
		{
			printf(CONSOLE_COLOR_RED "[!] CDummyNamespaceWalkCB deleted!\n" CONSOLE_COLOR_DEFAULT);
			delete this;
		}
		return result;
	}
	STDMETHOD(FoundItem)(IShellFolder* psf, PCUITEMID_CHILD pidl) override
	{
		printf("[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(EnterFolder)(IShellFolder* psf, PCUITEMID_CHILD pidl) override
	{
		printf("[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(LeaveFolder)(IShellFolder* psf, PCUITEMID_CHILD pidl) override
	{
		printf("[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(InitializeProgressDialog)(LPWSTR* ppszTitle, LPWSTR* ppszCancel) override
	{
		printf("[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
	STDMETHOD(WalkComplete)(HRESULT hr) override
	{
		printf("[*] CDummyNamespaceWalkCB %s\n", __func__);
		return S_OK;
	}
private:
	ULONG _refCount = 1;
};

struct CDummyNamespaceWalk : INamespaceWalk
{
	CDummyNamespaceWalk()
	{
		puts(CONSOLE_COLOR_RED "[!] CDummyNamespaceWalk constructor!" CONSOLE_COLOR_DEFAULT);
	}
	CDummyNamespaceWalk(CDummyNamespaceWalk&) = delete;
	CDummyNamespaceWalk& operator = (CDummyNamespaceWalk&) = delete;

	STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
	{
		RPC_CSTR pszGuid{};
		std::ignore = UuidToStringA(&riid, &pszGuid);
		printf(CONSOLE_COLOR_GREEN "[*] CDummyNamespaceWalk %s start, GUID=%s\n" CONSOLE_COLOR_DEFAULT, __func__, pszGuid);
		RpcStringFreeA(&pszGuid);

		static const QITAB rgqit[] = {
			QITABENT(CDummyNamespaceWalk, INamespaceWalk),
			{},
		};
		puts("[*] CDummyNamespaceWalk QISearch start");
		int hr = QISearch(this, rgqit, riid, ppv);
		printf("[*] CDummyNamespaceWalk QISearch end, hr = %08x\n", hr);
		printf(CONSOLE_COLOR_GREEN "[*] CDummyNamespaceWalk %s end, hr=%08x\n" CONSOLE_COLOR_DEFAULT, __func__, hr);
		return hr;
	}
	STDMETHOD_(ULONG, AddRef)() override
	{
		auto result = InterlockedIncrement(&_refCount);
		printf(CONSOLE_COLOR_YELLOW "[+] CDummyNamespaceWalk %s, RefCount=%d\n" CONSOLE_COLOR_DEFAULT, __func__, result);
		return result;
	}
	STDMETHOD_(ULONG, Release)() override
	{
		ULONG result = InterlockedDecrement(&_refCount);
		printf(CONSOLE_COLOR_YELLOW "[+] CDummyNamespaceWalk %s, RefCount=%d\n" CONSOLE_COLOR_DEFAULT, __func__, result);
		if (result == 0)
		{
			printf(CONSOLE_COLOR_RED "[!] CDummyNamespaceWalk deleted!\n" CONSOLE_COLOR_DEFAULT);
			delete this;
		}
		return result;
	}
	HRESULT __stdcall Walk(IUnknown* punkToWalk, DWORD dwFlags, int cDepth, INamespaceWalkCB* pnswcb) override
	{
		printf("[*] CDummyNamespaceWalk %s start, call pnswcb\n", __func__);
		LPWSTR pszTitle = nullptr, pszCancel = nullptr;
		pnswcb->InitializeProgressDialog(&pszTitle, &pszCancel);
		printf("[*] CDummyNamespaceWalk %s end\n", __func__);
		return S_OK;
	}
	HRESULT __stdcall GetIDArrayResult(UINT* pcItems, PIDLIST_ABSOLUTE** prgpidl) override
	{
		printf("[*] CDummyNamespaceWalk %s\n", __func__);
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
	STDMETHOD_(ULONG, AddRef)() override
	{
		return 1; // This object is a singleton.
	}
	STDMETHOD_(ULONG, Release)() override
	{
		return 1; // This object is a singleton.
	}
	STDMETHOD(CreateInstance)(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override
	{
		RPC_CSTR pszGuid{};
		std::ignore = UuidToStringA(&riid, &pszGuid);
		printf(CONSOLE_COLOR_BLUE "[*] %s start, GUID=%s\n" CONSOLE_COLOR_DEFAULT, __func__, pszGuid);
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

		printf(CONSOLE_COLOR_BLUE "[*] %s end, hr = %08x\n" CONSOLE_COLOR_DEFAULT, __func__, hr);
		return hr;
	}
	STDMETHOD(LockServer)(BOOL fLock) override
	{
		// 大丈夫、呼び出されていない。
		printf(CONSOLE_COLOR_RED "[*] %s, fLock=%d\n" CONSOLE_COLOR_DEFAULT, __func__, fLock);
		return S_OK; // I dont know that behavior is expected...
	}
};
