#pragma once


#include <cstdio>
#include <windows.h>
#pragma comment(lib, "ole32.lib") // for CoInitialize etc.
#pragma comment(lib, "Rpcrt4.lib")

inline GUID StringToGuid(const wchar_t* lpszGuid)
{
	GUID guid{};
	HRESULT hr = IIDFromString(lpszGuid, &guid);
	if (FAILED(hr)) { throw hr; }
	return guid;
}

#include <Shobjidl.h>
#pragma comment(lib, "Shell32.lib") // for IID_NamespaceWalk, etc.

inline const GUID CLSID_DummyNamespaceWalk = StringToGuid(L"{3A9F4A2A-C7C6-4D9D-8F9E-D71EE470B57F}");
inline const GUID CLSID_DummyNamespaceWalkCB = StringToGuid(L"{1B94A81D-A105-428D-8902-849F8D6483A2}");
