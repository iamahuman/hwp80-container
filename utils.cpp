#include "common.h"

ATLAXWININIT g_lpfnAtlAxWinInit = NULL;
ATLAXATTACHCONTROLPROC g_lpfnAtlAxAttachControl = NULL;
ATLAXCREATECONTROL g_lpfnAtlAxCreateControl = NULL;
ATLAXGETCONTROL g_lpfnAtlAxGetControl = NULL;

void print_err_inner(LPCTSTR lpFile, int nLine, LPCTSTR lpRepr, DWORD code, HRESULT hresult) {
	DWORD res = 0, res2 = 0;
	LPTSTR lpBuf = NULL, lpBuf2 = NULL;
	DWORD_PTR args[] = {(DWORD_PTR) lpFile, (DWORD_PTR) nLine, (DWORD_PTR) lpRepr, 0, (DWORD_PTR) code, 0};

	if (FAILED(hresult)) {
		//_com_error errobj(hresult);
		//args[3] = (DWORD_PTR) errobj.ErrorMessage();
		args[3] = (DWORD_PTR) TEXT("COM failure");
		args[4] = (DWORD_PTR) hresult;
	} else if (code != 0) {
		do {
			res = FormatMessage(
				FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL, code, 0, (LPTSTR) &lpBuf, 0, NULL);
			if (res > 0) break;
		} while (0);
		args[3] = (DWORD_PTR) (res != 0 && lpBuf != NULL ? lpBuf : TEXT("Unknown error"));
	} else {
		args[3] = (DWORD_PTR) TEXT("Attempted to report null error to user");
	}

	res2 = FormatMessage(
		FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ARGUMENT_ARRAY,
		//lpRepr != NULL ? TEXT("Error on %1!s!:%2!d! (%3!s!): %4!s! (%5!x!)"): TEXT("Error on %1!s!:%2!d!: %4!s! (%5!x!)"),
		lpRepr != NULL ? TEXT("[E]%1!s!:%2!d!:<<%3!s!>>:%5!x!:%4!s!\r\n%0") : TEXT("[E]%1!s!:%2!d!::%5!x!:%4!s!\r\n%0"),
		0, 0, (LPTSTR) &lpBuf2, 0, (va_list *) args);

	if (res2 == 0 || lpBuf2 == NULL) {
		//MessageBox(g_hWnd, TEXT("wtf"), NULL, MB_ICONHAND | MB_OK);
		OutputDebugString(TEXT("[E]?\r\n"));
	} else {
		//MessageBox(g_hWnd, lpBuf2, NULL, MB_ICONHAND | MB_OK);
		OutputDebugString(lpBuf2);
	}

	if (lpBuf2 != NULL)
		LocalFree(lpBuf2);

	if (lpBuf != NULL)
		LocalFree(lpBuf);
}

int atl_init() {
	//WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), "[atl_init] pending..\n", 21, &tmp, NULL);
	HRESULT hr;

	if ((g_hATL = LoadLibrary(TEXT("atl.dll"))) == NULL) {
		print_lasterr("Loading Library `atl.dll`");
		goto fail_load;
	}

	hr = CoInitialize(0);
	if (FAILED(hr)) {
		print_hresult("CoInitialize", hr);
		goto fail_co_init;
	}

	g_lpfnAtlAxAttachControl = (ATLAXATTACHCONTROLPROC) GetProcAddress(g_hATL, "AtlAxAttachControl");
	if (g_lpfnAtlAxAttachControl == NULL) {
		print_lasterr("Getting address of function `AtlAxAttachControl` of `atl.dll`");
		goto fail_init_atl;
	}

	g_lpfnAtlAxWinInit = (ATLAXWININIT) GetProcAddress(g_hATL, "AtlAxWinInit");
	if (g_lpfnAtlAxWinInit == NULL) {
		print_lasterr("Getting address of function `AtlAxWinInit` of `atl.dll`");
		goto fail_init_atl;
	}

	g_lpfnAtlAxCreateControl = (ATLAXCREATECONTROL) GetProcAddress(g_hATL, "AtlAxCreateControl");
	if (g_lpfnAtlAxCreateControl == NULL) {
		print_lasterr("Getting address of function `AtlAxWinCreateControl` of `atl.dll`");
		goto fail_init_atl;
	}

	g_lpfnAtlAxGetControl = (ATLAXGETCONTROL) GetProcAddress(g_hATL, "AtlAxGetControl");
	if (g_lpfnAtlAxGetControl == NULL) {
		print_lasterr("Getting address of function `AtlAxWinGetControl` of `atl.dll`");
		goto fail_init_atl;
	}

	if (!(*g_lpfnAtlAxWinInit)()) {
		print_lasterr("AtlAxWinInit");
		goto fail_init_atl;
	}

	//WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), "[atl_init] OK\n", 14, &tmp, NULL);
	return 0;

fail_init_atl:
	g_lpfnAtlAxWinInit = NULL;
	g_lpfnAtlAxAttachControl = NULL;

	CoUninitialize();

fail_co_init:
	if (g_hATL != NULL) {
		FreeLibrary(g_hATL);
		g_hATL = NULL;
	}

fail_load:
	//WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), "[atl_init] FAIL\n", 16, &tmp, NULL);
	return -1;
}

int atl_deinit() {
	//WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), "[atl_deinit] pending..\n", 23, &tmp, NULL);

	UnregisterClass(TEXT("AtlAxWin"), g_hInst);

	CoUninitialize();

	if (g_hATL != NULL) {
		FreeLibrary(g_hATL);
		g_hATL = NULL;
	}

	//WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), "[atl_deinit] OK\n", 16, &tmp, NULL);
	return 0;
}
