#include "common.h"
#include <commctrl.h>
#include <shellapi.h>
#include <strsafe.h>
#include "hwpctrl.h"

/*** Common definitions and global variables ***/

HINSTANCE g_hInst = NULL;
HMODULE g_hATL = NULL;
HWND g_hWnd = NULL;

HWND g_hCtrl = NULL;
IUnknown *g_container = NULL;
IDispatch *g_hwpCtrl = NULL;

#define REQ_INVALID 0
#define REQ_OPEN 1
#define REQ_SAVEAS 2

struct request {
	int type;
	LPOLESTR lpPath;
	LPOLESTR lpFormat;
	LPOLESTR lpArg;
} g_requests[2048] = {{0}};

size_t g_req_cnt = 0;

DIDX_HWPCTRL g_didx_hwpCtrl;

int ole_init();
int hwp_init();
void ole_deinit();
LRESULT CALLBACK window_proc(HWND, UINT, WPARAM, LPARAM);
ATOM register_class();
int parse_opts();
int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int);

int g_showCmd = 0, g_bOnce = 0;

LPWSTR lk_strdup_w(LPCWSTR s) {
	LPCWSTR ptr = s;
	LPWSTR ret;
	size_t i, chars;

	while (*(ptr++))
		;

	chars = (size_t)(ptr - s);

	ret = (LPWSTR) LocalAlloc(LMEM_FIXED | LMEM_ZEROINIT, chars * sizeof(WCHAR));

	if (ret != NULL) {
		for (i = 0; i < chars; i++)
			ret[i] = s[i];
		if (ret[i-1] != L'\0') // coherency check failure (probably race condition)
			TerminateProcess(GetCurrentProcess(), -1337);
	} else TerminateProcess(GetCurrentProcess(), -2);

	return ret;
}


int ole_init() {
	int res = -28;
	HRESULT hr;
	OLECHAR str_RegisterModule[] = L"RegisterModule",
		str_Open[] = L"Open",
		str_Save[] = L"Save",
		str_SaveAs[] = L"SaveAs",
		str_Clear[] = L"Clear",
		str_ShowToolBar[] = L"ShowToolBar";
	size_t i;

	hr = (*g_lpfnAtlAxCreateControl)(L"{BD9C32DE-3155-4691-8972-097D53B10052}", g_hCtrl, NULL, &g_container);
	if (FAILED(hr)) {
		print_hresult("AtlAxCreateControl", hr);
		goto fail_co_create;
	}

	{
		IUnknown *ptrCtrl = NULL;
		hr = (*g_lpfnAtlAxGetControl)(g_hCtrl, &ptrCtrl);
		if (FAILED(hr)) {
			print_hresult("AtlAxGetControl", hr);
			goto fail_get_ctrl;
		}

		hr = ptrCtrl->QueryInterface(IID_IDispatch, (void **)&g_hwpCtrl);
		ptrCtrl->Release();
		if (FAILED(hr)) {
			print_hresult("QueryInterface IID_IDispatch", hr);
			goto fail_get_ctrl;
		}
	}

	for (i = 0; i < DIDX_HWPCTRL_NR; i++) {
		LPOLESTR rgszNames[1];
		DISPID dispId;
		switch (i) {
		case DIDX_HWPCTRL_REGISTER_MODULE:	rgszNames[0] = str_RegisterModule;	break;
		case DIDX_HWPCTRL_OPEN:			rgszNames[0] = str_Open;		break;
		case DIDX_HWPCTRL_SAVE:			rgszNames[0] = str_Save;		break;
		case DIDX_HWPCTRL_SAVE_AS:		rgszNames[0] = str_SaveAs;		break;
		case DIDX_HWPCTRL_CLEAR:		rgszNames[0] = str_Clear;		break;
		case DIDX_HWPCTRL_SHOW_TOOL_BAR:	rgszNames[0] = str_ShowToolBar;		break;
		default:
			goto fail_get_ids_of_names;
		}
		hr = g_hwpCtrl->GetIDsOfNames(IID_NULL, rgszNames, 1, LOCALE_USER_DEFAULT, &dispId);
		if (FAILED(hr)) {
			print_hresult("GetIDsOfNames", hr);
			goto fail_get_ids_of_names;
		}
		CopyMemory(&g_didx_hwpCtrl[i], &dispId, sizeof(DISPID));
		{
			LPWSTR lpBuf = NULL;
			DWORD_PTR args[] = {(DWORD_PTR) rgszNames[0], (DWORD_PTR) dispId, 0};

			FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_ARGUMENT_ARRAY |
					FORMAT_MESSAGE_FROM_STRING,
					TEXT("[V]%1!s!: %2!08x!\r\n%0"),
					0, 0, (LPTSTR) &lpBuf, 0, (va_list *) args);
			if (lpBuf != NULL) {
				OutputDebugStringW(lpBuf);
				LocalFree(lpBuf);
				lpBuf = NULL;
			}
		}
	}

	if ((res = hwp_init()) != 0) {
		goto fail_hwp_init;
	}

	return 0;

fail_hwp_init:
fail_get_ids_of_names:
fail_get_ctrl:
	if (g_hwpCtrl != NULL) {
		g_hwpCtrl->Release();
		g_hwpCtrl = NULL;
	}

fail_co_create:
	return res;
}

int hwp_init() {
	HRESULT hr;
	VARIANT vres;
	size_t i;

	hr = HwpCtrl_RegisterModule(g_hwpCtrl, g_didx_hwpCtrl, &vres, L"FilePathCheckDLL", L"FilePathChecker");
	if (FAILED(hr)) {
		print_hresult("HwpCtrl::RegisterModule", hr);
		return -2;
	}

	if (vres.vt != VT_BOOL || vres.boolVal != VARIANT_TRUE) {
		print_hresult("HwpCtrl::RegisterModule", 0xdeadbeef);
		return -3;
	}

	OutputDebugString(TEXT("[I]HwpInitOK\r\n"));

	for (i = 0; i < g_req_cnt; i++) {
		const struct request *const req = &g_requests[i];

		if (req->type == REQ_OPEN) {
			hr = i == 0 ? S_OK : HwpCtrl_Clear(g_hwpCtrl, g_didx_hwpCtrl, &vres, hwpDiscard);
			if (FAILED(hr)) {
				print_hresult("HwpCtrl::Clear", hr);
				return -4;
			}

			hr = HwpCtrl_Open(g_hwpCtrl, g_didx_hwpCtrl, &vres, req->lpPath, req->lpFormat, req->lpArg ? req->lpArg : L"lock:false;setcurdir:false;suspendpassword:true;forceopen:true;code:cp949;versionwarning:false;");
			if (FAILED(hr)) {
				print_hresult("HwpCtrl::Open", hr);
				return -5;
			}

			if (vres.vt != VT_BOOL || vres.boolVal != VARIANT_TRUE) {
				print_hresult("HwpCtrl::Open", 0xdeadbeef);
				return 1;
			}
		} else if (req->type == REQ_SAVEAS) {
			hr = HwpCtrl_SaveAs(g_hwpCtrl, g_didx_hwpCtrl, &vres, req->lpPath, req->lpFormat, req->lpArg ? req->lpArg : L"lock:false;backup:false;compress:false;autosave:false;");
			if (FAILED(hr)) {
				print_hresult("HwpCtrl::SaveAs", hr);
				return -6;
			}

			if (vres.vt != VT_BOOL || vres.boolVal != VARIANT_TRUE) {
				print_hresult("HwpCtrl::SaveAs", 0xdeadbeef);
				return 1;
			}
		}
	}

	return 0;
}

void ole_deinit() {
	if (g_container != NULL) {
		g_container->Release();
		g_container = NULL;
	}

	if (g_hwpCtrl != NULL) {
		g_hwpCtrl->Release();
		g_hwpCtrl = NULL;
	}
}

#define ONESHOT_QLIMIT_TIMER 0x384e1337

int g_exitCode = -88;

LRESULT CALLBACK window_proc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam) {
	int res;

	switch (nMsg) {
	case WM_CREATE:
		g_hCtrl = CreateWindow(TEXT("AtlAxWin"), TEXT(""), WS_CHILD | WS_VISIBLE,
			0, 0, 600, 440,
			hWnd, NULL, g_hInst, NULL);

		if (g_hCtrl == NULL) {
			print_lasterr("CreateWindow - AtlAxWin");
			g_exitCode = -7;
			return -1;
		}

		if ((res = ole_init()) != 0) {
			g_exitCode = res;
			return -1;
		}

		if (g_bOnce == 1) { // Terminate immediately
			g_exitCode = 0;
			return -1;
		} else if (g_bOnce == 2) { // Delayed stop
			if (!SetTimer(hWnd, ONESHOT_QLIMIT_TIMER, USER_TIMER_MINIMUM, NULL)) {
				g_exitCode = -99;
				return -1;
			}
			g_exitCode = 0;
		} else { // Keep it open
			g_exitCode = 0;
			return 0;
		}
		break;
	case WM_SIZE:
		if (g_hCtrl != NULL) {
			RECT cliRect = {0, 0, 600, 440};
			GetClientRect(g_hWnd, &cliRect);
			MoveWindow(g_hCtrl, 0, 0, cliRect.right, cliRect.bottom, TRUE);
		}
		break;
	case WM_TIMER:
		if (wParam == ONESHOT_QLIMIT_TIMER) {
			DestroyWindow(hWnd);
			break;
		}
		return DefWindowProc(hWnd, nMsg, wParam, lParam);
	case WM_DESTROY:
		ole_deinit();

		if (g_hCtrl != NULL) {
			DestroyWindow(g_hCtrl);
			g_hCtrl = NULL;
		}

		if (g_hWnd == hWnd)
			g_hWnd = NULL;

		PostQuitMessage(g_exitCode);
		break;
	default:
		return DefWindowProc(hWnd, nMsg, wParam, lParam);
	}
	return 0;
}

ATOM register_class() {
	WNDCLASS wndClass = {0};

	wndClass.style = CS_OWNDC; // so that its graphical output can be captured reliably
	wndClass.lpfnWndProc = window_proc;
	wndClass.cbClsExtra = 0;
	wndClass.cbWndExtra = 0;
	wndClass.hInstance = g_hInst;
	wndClass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wndClass.hCursor = LoadCursor(NULL, IDC_ARROW);
	wndClass.hbrBackground = NULL;
	wndClass.lpszMenuName = NULL;
	wndClass.lpszClassName = TEXT("Luke1337_HwpAxOLEAutCtr");

	return RegisterClass(&wndClass);
}

int parse_opts() {
	int argc = 0, i, is_open ,is_save;
	LPWSTR *argv, lpCmdLine, lpArg;
	const LCID lcid = LOCALE_USER_DEFAULT;
	const DWORD flags = NORM_IGNORECASE;

	lpCmdLine = GetCommandLineW();
	if (lpCmdLine == NULL)
		return -1;
	
	argv = CommandLineToArgvW(lpCmdLine, &argc);
	if (argv == NULL)
		return -1;
	
	for (i = 1; i < argc; i++) {
		lpArg = argv[i];
		//OutputDebugStringW(lpArg);
		is_open = CompareStringW(lcid, flags, lpArg, -1, L"-open", -1) == CSTR_EQUAL;
		is_save = CompareStringW(lcid, flags, lpArg, -1, L"-save", -1) == CSTR_EQUAL;
		if (is_open || is_save) {
			if (!(i + 1 < argc && g_req_cnt + 1 <= sizeof(g_requests)/sizeof(*g_requests)))
				continue;
			g_requests[g_req_cnt].type = is_save ? REQ_SAVEAS : REQ_OPEN;
			g_requests[g_req_cnt].lpPath = lk_strdup_w(argv[++i]);
			if (i + 1 < argc && (
					(L'A' <= argv[i+1][0] && argv[i+1][0] <= L'Z') ||
					(L'a' <= argv[i+1][0] && argv[i+1][0] <= L'z'))) {
				g_requests[g_req_cnt].lpFormat = lk_strdup_w(argv[++i]);
				if (i + 1 < argc && (
						(L'A' <= argv[i+1][0] && argv[i+1][0] <= L'Z') ||
						(L'a' <= argv[i+1][0] && argv[i+1][0] <= L'z'))) {
					g_requests[g_req_cnt].lpArg = lk_strdup_w(argv[++i]);
				} else {
					g_requests[g_req_cnt].lpArg = NULL;
				}
			} else {
				g_requests[g_req_cnt].lpFormat = NULL;
				g_requests[g_req_cnt].lpArg = NULL;
			}
			g_req_cnt++;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-openmultiple", -1) == CSTR_EQUAL) {
			while (i + 1 < argc && g_req_cnt + 1 <= sizeof(g_requests)/sizeof(*g_requests)) {
				g_requests[g_req_cnt].type = REQ_OPEN;
				g_requests[g_req_cnt].lpPath = lk_strdup_w(argv[++i]);
				g_requests[g_req_cnt].lpFormat = NULL; // smart guess (emulate end-user usage behaviour)
				g_requests[g_req_cnt].lpArg = NULL;
				g_req_cnt++;
			}
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-background", -1) == CSTR_EQUAL) {
			g_showCmd = SW_HIDE;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-foreground", -1) == CSTR_EQUAL) {
			g_showCmd = SW_SHOW;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-unfocused", -1) == CSTR_EQUAL) {
			g_showCmd = SW_SHOWNOACTIVATE;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-minimized", -1) == CSTR_EQUAL) {
			g_showCmd = SW_MINIMIZE;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-maximized", -1) == CSTR_EQUAL) {
			g_showCmd = SW_MAXIMIZE;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-oneshot", -1) == CSTR_EQUAL) {
			g_bOnce = 1;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-oneshot-delayed", -1) == CSTR_EQUAL) {
			g_bOnce = 2;
		} else if (CompareStringW(lcid, flags, lpArg, -1, L"-keep", -1) == CSTR_EQUAL) {
			g_bOnce = 0;
		}
	}

	LocalFree(argv);

	OutputDebugString(TEXT("[I]parse_opts done!\r\n"));

	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd) {
	int res = -31;
	ATOM atomCls = 0;
	MSG msg;
	INITCOMMONCONTROLSEX iccex = {0};

	g_showCmd = nShowCmd;

	if (parse_opts() != 0)
		goto early_fail;

	iccex.dwSize = (DWORD) sizeof(iccex);
	iccex.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&iccex);

	if (atl_init() != 0)
		goto early_fail;

	if (!(atomCls = register_class())) {
		print_lasterr("RegisterClass");
		goto fail;
	}

	g_hWnd = CreateWindow(MAKEINTATOM(atomCls), TEXT("Dummy Fuzzer Window"),
		WS_OVERLAPPEDWINDOW & ~WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 640, 480,
		NULL, NULL, hInstance, NULL);

	if (g_hWnd == NULL) {
		print_lasterr("CreateWindow - Dummy Fuzzer Window");
		res = g_exitCode;
		goto fail;
	}

	ShowWindow(g_hWnd, g_showCmd);
	UpdateWindow(g_hWnd);

	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	res = (int) msg.wParam;

fail:
	if (atomCls != 0)
		UnregisterClass(MAKEINTATOM(atomCls), hInstance);

	atl_deinit();

early_fail:

	{
		size_t i;
		for (i = 0; i < g_req_cnt; i++) {
			const struct request *const req = &g_requests[i];
			if (req->lpPath != NULL)
				LocalFree(req->lpPath);
			if (req->lpFormat != NULL)
				LocalFree(req->lpFormat);
			if (req->lpArg != NULL)
				LocalFree(req->lpArg);
		}
	}

	// workaround for msvcr90 CRT collsion bug -- somehow __dlonexit hooks got screwed up
	TerminateProcess(GetCurrentProcess(), res);

	return res;
}

extern "C" {

void start(void) {
	int ret;
	STARTUPINFO si = {0};
	LPSTR lpCmdLine;
	si.cb = (DWORD) sizeof(STARTUPINFO);
	si.wShowWindow = SW_SHOW;

	HINSTANCE hInstance = GetModuleHandle(NULL);
	lpCmdLine = GetCommandLineA();
	GetStartupInfo(&si);

	ret = WinMain(hInstance, NULL, lpCmdLine, (int) si.wShowWindow);

	ExitProcess(ret);

	for (;;);
}

void _start(void) {
	start();

	for (;;);
}

}
