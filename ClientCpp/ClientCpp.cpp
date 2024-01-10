#include "../ServerDll/ClsIds.hpp"

int main()
{
	HRESULT hr = CoInitialize(nullptr);
	if (FAILED(hr)) { throw hr; }

	// COM�I�u�W�F�N�g����
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

	// �K���Ɏg��
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

	// �g���I������̂ŉ��
	hr = pWalkCB2->Release();
	hr = pWalkCB->Release();
	hr = pWalk->Release();

	// ���Ɋ��S��Release�ς݁B�X�Ɏg�����Ƃ����User-After-Free�Ȃ̂Ŗ���`����
	// hr = pWalk->Release(); // ��������ADebug�r���h��DLL���ƁuException thrown: read access violation. pWalk->was 0xFFFFFFFFFFFFFFEF.�v�G���[�ARelease�r���h��DLL���Ƃ����ł͉����N���炸

	CoUninitialize();
}
