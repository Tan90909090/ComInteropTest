// ClientTestCpp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "../ServerDll/ClsIds.hpp"

int main()
{
	HRESULT hr = CoInitialize(nullptr);
	if (FAILED(hr)) { throw hr; }

	INamespaceWalk* pWalk = nullptr;
	INamespaceWalkCB* pWalkCB = nullptr;
	// �����Ńt���[�Y����邩�炱�ˁ`�`�`�`�`
	// ������񖢎��s����"0x80040154"
	// ��MessageLoop�񂵂Ă���0x80004002�ɂȂ����B������CreateInstance���Ă΂�Ă������I
	// ����ACoCreateInstance��IUnknown�擾����΂����ƒʂ������B��IFoo�ւ�QueryInterface�͎��s���܂����[�H�H�H
	// ���[��ACLSCTX_LOCAL_SERVER���߂��ق��������̂����c�c�������c�c
	// ������|CLSCTX_INPROC_HANDLER��OR�w�肵�Ă݂����Ǖω��Ȃ�
	// https://ichigopack.net/win32com/com_exeserver_2.html#:~:text=%E3%81%AA%E3%82%8B%E3%81%A7%E3%81%97%E3%82%87%E3%81%86%E3%80%82-,%E3%83%A1%E3%82%BD%E3%83%83%E3%83%89%E4%BB%B2%E4%BB%8B%E3%81%AE%E6%96%B9%E6%B3%95,-%E5%85%88%E3%81%AEproxy �̂��Ƃł́H����͂����ł́H
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

	// �ߏ�Release���悤�Ƃ����User-After-Free�Ȃ̂Ŗ���`����
	// hr = pWalk->Release(); // ��������uException thrown: read access violation. pWalk->was 0xFFFFFFFFFFFFFFEF.�v�G���[

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
