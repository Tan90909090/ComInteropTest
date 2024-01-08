// ClientTestCpp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "../ServerDll/ClsIds.hpp"

int main()
{
	HRESULT hr = CoInitialize(nullptr);
	if (FAILED(hr)) { throw hr; }

	INamespaceWalk* pWalk = nullptr;
	INamespaceWalkCB* pWalkCB = nullptr;
	// ここでフリーズされるからこね～～～～～
	// もちろん未実行だと"0x80040154"
	// →MessageLoop回してやれば0x80004002になった。ちゃんとCreateInstanceが呼ばれていそう！
	// あれ、CoCreateInstanceでIUnknown取得すればちゃんと通ったぞ。→IFooへのQueryInterfaceは失敗しますがー？？？
	// うーん、CLSCTX_LOCAL_SERVER諦めたほうがいいのかも……分からん……
	// 試しに|CLSCTX_INPROC_HANDLERもOR指定してみたけど変化なし
	// https://ichigopack.net/win32com/com_exeserver_2.html#:~:text=%E3%81%AA%E3%82%8B%E3%81%A7%E3%81%97%E3%82%87%E3%81%86%E3%80%82-,%E3%83%A1%E3%82%BD%E3%83%83%E3%83%89%E4%BB%B2%E4%BB%8B%E3%81%AE%E6%96%B9%E6%B3%95,-%E5%85%88%E3%81%AEproxy のことでは？それはそうでは？
	auto dwClsContext = CLSCTX_INPROC_SERVER;
	hr = CoCreateInstance(
		CLSID_DummyNamespaceWalk,
		nullptr,
		dwClsContext,
		IID_PPV_ARGS(&pWalk));
	if (FAILED(hr)) { throw hr; }
	hr = CoCreateInstance(
		CLSID_DummyNamespaceWalkCB,
		nullptr,
		dwClsContext,
		IID_PPV_ARGS(&pWalkCB));
	if (FAILED(hr)) { throw hr; }

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

	hr = pWalkCB2->Release();
	hr = pWalkCB->Release();
	hr = pWalk->Release();

	// 過剰ReleaseしようとするとUser-After-Freeなので未定義動作
	// hr = pWalk->Release(); // 試したら「Exception thrown: read access violation. pWalk->was 0xFFFFFFFFFFFFFFEF.」エラー

	CoUninitialize();
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
