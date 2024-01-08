// ComComponentRegisterer.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "../ServerDll/ClsIds.hpp"
#include "../ServerDll/ImplementedClasses.hpp"

int main()
{
	CClassFactory<CDummyNamespaceWalk> classFactoryDummyNamespaceWalk;
	CClassFactory<CDummyNamespaceWalkCB> classFactoryDummyNamespaceWalkCB;

	int hr = CoInitialize(nullptr);
	if (FAILED(hr)) { throw hr; }

	// https://ichigopack.net/win32com/com_exeserver_2.html#:~:text=%E3%81%AA%E3%82%8B%E3%81%A7%E3%81%97%E3%82%87%E3%81%86%E3%80%82-,%E3%83%A1%E3%82%BD%E3%83%83%E3%83%89%E4%BB%B2%E4%BB%8B%E3%81%AE%E6%96%B9%E6%B3%95,-%E5%85%88%E3%81%AEproxy
	// LOCAL_SERVER��EXE�̏ꍇ�A����interface���v���Z�X�ԒʐM�����邽�߂�proxy/stub���K�v�����I�ʓ|���������I���|�ꂾ���IINPROC_SERVER��DLL�ɂ��܂��ˁI
	// proxy/stub������INPROC��DLL���K�v�炵���I�I���I���I https://ichigopack.net/win32com/com_proxystub_1.html
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

	puts(CONSOLE_COLOR_GREEN "Running..." CONSOLE_COLOR_DEFAULT);
	// getchar�҂����Ǝ~�܂����̂Ń��b�Z�[�W���[�v���K�v�炵���B
	MessageBoxW(nullptr, L"Com Server is running. Close this to unregister.", L"Title", MB_OK);

	hr = CoSuspendClassObjects();
	if (FAILED(hr)) { throw hr; }

	hr = CoRevokeClassObject(dwRegisterDummyNamespaceWalkCB);
	if (FAILED(hr)) { throw hr; }
	hr = CoRevokeClassObject(dwRegisterDummyNamespaceWalk);
	if (FAILED(hr)) { throw hr; }
	hr = CoRevokeClassObject(dwRegisterFoo);
	if (FAILED(hr)) { throw hr; }
	CoUninitialize();
}
