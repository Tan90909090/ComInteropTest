#include "ClsIds.hpp"
#include "ImplementedClasses.hpp"

#include <shlwapi.h>
#pragma comment(lib, "Shlwapi.lib") // for QISearch and SHDeleteKey
#pragma comment(lib, "Advapi32") // for registry functions

#include <array>
#include <string>
#include <type_traits>
using namespace std::string_literals;

namespace
{
	std::wstring g_dllPath;
	const CLSID g_arrayClsid[] = { CLSID_DummyNamespaceWalk, CLSID_DummyNamespaceWalkCB };
	CClassFactory<CDummyNamespaceWalk> g_classFactoryDummyNamespaceWalk;
	CClassFactory<CDummyNamespaceWalkCB> g_classFactoryDummyNamespaceWalkCB;

	std::wstring GuidToString(const GUID& guid)
	{
		std::array<wchar_t, 40> str{};
		int result = StringFromGUID2(guid, str.data(), static_cast<int>(str.size()));
		if (result == 0) { throw result; }
		return str.data();
	}

	bool ExistSubKey(HKEY hKeyRoot, const std::wstring& subKey)
	{
		HKEY hKey = nullptr;
		LSTATUS status = RegOpenKeyExW(hKeyRoot, subKey.c_str(), 0, KEY_READ, &hKey);
		if (hKey != nullptr)
		{
			RegCloseKey(hKey);
		}
		return status == ERROR_SUCCESS;
	}
	HRESULT CreateAndSetClsidKeyAndValue(const std::wstring& keyName, const std::wstring& valueName, const std::wstring& data)
	{
		HKEY hKey{};
		auto status = RegCreateKeyExW(
			HKEY_CLASSES_ROOT,
			keyName.c_str(),
			0,
			nullptr,
			REG_OPTION_NON_VOLATILE,
			KEY_WRITE,
			nullptr,
			&hKey,
			nullptr);
		if (status != ERROR_SUCCESS)
		{
			return HRESULT_FROM_WIN32(status);
		}

		status = RegSetValueExW(
			hKey,
			valueName.c_str(),
			0,
			REG_SZ,
			static_cast<const BYTE*>(static_cast<const void*>(data.c_str())),
			static_cast<DWORD>(data.size()) * sizeof(std::decay_t<decltype(data)>::value_type));
		if (status != ERROR_SUCCESS)
		{
			return HRESULT_FROM_WIN32(status);
		}

		status = RegCloseKey(hKey);
		if (status != ERROR_SUCCESS)
		{
			return HRESULT_FROM_WIN32(status);
		}

		return S_OK;
	};
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved) noexcept
{
	if (fdwReason == DLL_PROCESS_ATTACH)
	{
		g_dllPath.resize(MAX_PATH);
		GetModuleFileNameW(hinstDLL, g_dllPath.data(), static_cast<DWORD>(g_dllPath.size()));
		g_dllPath.shrink_to_fit();
	}

	return TRUE;
}

// Export functions
// Anottations are copied from combaseapi.h
__control_entrypoint(DllExport) STDAPI DllCanUnloadNow()
{
	return S_FALSE; // return a constant value for simplicity
}

_Check_return_ STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
	if (ppv == nullptr) { return E_POINTER; }

	*ppv = NULL;
	if (rclsid == CLSID_DummyNamespaceWalk)
	{
		return g_classFactoryDummyNamespaceWalk.QueryInterface(riid, ppv);
	}
	else if (rclsid == CLSID_DummyNamespaceWalkCB)
	{
		return g_classFactoryDummyNamespaceWalkCB.QueryInterface(riid, ppv);
	}

	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllRegisterServer()
{
	// I omitted ProgID because I won't use it.
	for (const auto& clsid : g_arrayClsid)
	{
		auto keyClsid = L"CLSID\\"s + GuidToString(clsid);
		auto keyInprocServer32 = keyClsid + L"\\InprocServer32"s;
		HRESULT hr;

		hr = CreateAndSetClsidKeyAndValue(keyClsid, L""s, L"Dummy Class for Test"s);
		if (FAILED(hr)) { return hr; }
		hr = CreateAndSetClsidKeyAndValue(keyInprocServer32, L""s, g_dllPath);
		if (FAILED(hr)) { return hr; }
		hr = CreateAndSetClsidKeyAndValue(keyInprocServer32, L"ThreadingModel"s, L"Apartment"s);
		if (FAILED(hr)) { return hr; }
	}
	return S_OK;
}

STDAPI DllUnregisterServer()
{
	for (const auto& clsid : g_arrayClsid)
	{
		auto keyClsid = L"CLSID\\"s + GuidToString(clsid);

		// check where the target key exists. If the key is not exist, SHDeleteKeyW will fail.
		if (!ExistSubKey(HKEY_CLASSES_ROOT, keyClsid))
		{
			continue;
		}

		// I can not implement this feature using RegDeleteKeyExW nor RegDeleteTreeW... Deleting a key is hard...
		auto status = SHDeleteKeyW(HKEY_CLASSES_ROOT, keyClsid.c_str());
		if (status != ERROR_SUCCESS)
		{
			return HRESULT_FROM_WIN32(status);
		}
	}
	return S_OK;
}

STDAPI DllInstall(BOOL bInstall, _In_opt_ PCWSTR pszCmdLine)
{
	// I ignore a command line parameter for simplify.
	// As result, user need to administrator priviledge to `regsvr32.exe SampleDll.dll`.
	return bInstall ? DllRegisterServer() : DllUnregisterServer();
}
