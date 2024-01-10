#include "../ServerDll/ClsIds.hpp"

int main()
{
	HRESULT hr = CoInitialize(nullptr);
	if (FAILED(hr)) { throw hr; }

	// COMオブジェクト生成
	INamespaceWalk* pWalk = nullptr;
	INamespaceWalkCB* pWalkCB = nullptr;

	hr = CoCreateInstance(
		CLSID_DummyNamespaceWalk,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&pWalk));
	if (FAILED(hr)) { throw hr; }
	hr = CoCreateInstance(
		CLSID_DummyNamespaceWalkCB,
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&pWalkCB));
	if (FAILED(hr)) { throw hr; }

	// 適当に使う
	INamespaceWalkCB2* pWalkCB2 = nullptr;
	hr = pWalkCB->QueryInterface(&pWalkCB2);
	if (FAILED(hr)) { throw hr; }

	LPWSTR pszTitle = nullptr, pszCaption = nullptr;
	hr = pWalkCB->InitializeProgressDialog(&pszTitle, &pszCaption);
	if (FAILED(hr)) { throw hr; }

	hr = pWalkCB2->WalkComplete(S_OK);
	if (FAILED(hr)) { throw hr; }

	hr = pWalk->Walk(nullptr, 0, 0, pWalkCB);
	if (FAILED(hr)) { throw hr; }

	// 使い終わったので解放
	hr = pWalkCB2->Release();
	hr = pWalkCB->Release();
	hr = pWalk->Release();

	// 既に完全にRelease済み。更に使おうとするとUser-After-Freeなので未定義動作
	// hr = pWalk->Release(); // 試したら、DebugビルドのDLLだと「Exception thrown: read access violation. pWalk->was 0xFFFFFFFFFFFFFFEF.」エラー、ReleaseビルドのDLLだとここでは何も起こらず

	CoUninitialize();
}
