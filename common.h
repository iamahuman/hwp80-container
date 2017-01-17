#ifndef _LUKE1337_COMMON_H
#define _LUKE1337_COMMON_H

#define WIN32_LEAN_AND_MEAN // keep 'em simple stupid

#define WINVER 0x0501 // at least Windows XP
#define _WIN32_WINNT 0x0501

#define UNICODE
#define _UNICODE

#include <windows.h>
#include <oaidl.h>
#include <dispex.h>


/** Macro and inline utilities **/

#define SET_ARG_BSTR(vargs, idx, val) do { \
		(vargs)[idx].vt = ((val) == NULL ? VT_ERROR : VT_BSTR); \
		(vargs)[idx].bstrVal = ((val) == NULL ? NULL : SysAllocString((val))); \
	} while (0)

#define FREE_VARSTR(expr) if ((expr).vt == VT_BSTR && (expr).bstrVal != NULL) { \
		SysFreeString((expr).bstrVal); \
	}

inline static HRESULT ole_invoke_method(IDispatch *obj, DISPID dispId, VARIANT *result, VARIANTARG *vargs, UINT cArgs) {
	DISPPARAMS params = {0};
	params.rgvarg = vargs;
	params.cArgs = cArgs;
	params.cNamedArgs = 0;
	return obj->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, result, NULL, NULL);
}

/** dynamically loaded functions **/

typedef BOOL (WINAPI *ATLAXWININIT)();
typedef HRESULT (WINAPI *ATLAXATTACHCONTROLPROC)(IUnknown*, HWND, IUnknown**);  
typedef HRESULT (WINAPI *ATLAXCREATECONTROL)(LPCOLESTR, HWND, IStream*, IUnknown**);
typedef HRESULT (WINAPI *ATLAXGETCONTROL)(HWND, IUnknown**);
extern ATLAXWININIT g_lpfnAtlAxWinInit;
extern ATLAXATTACHCONTROLPROC g_lpfnAtlAxAttachControl;
extern ATLAXCREATECONTROL g_lpfnAtlAxCreateControl;
extern ATLAXGETCONTROL g_lpfnAtlAxGetControl;


/** global variables **/

extern HINSTANCE g_hInst;
extern HMODULE g_hATL;
extern HWND g_hWnd;


/** functions **/

extern void print_err_inner(LPCTSTR lpFile, int nLine, LPCTSTR lpRepr, DWORD code, HRESULT hresult);
#define print_err(repr, code) (print_err_inner(TEXT(__FILE__), __LINE__, TEXT(repr), (code), S_OK))
#define print_hresult(repr, code) (print_err_inner(TEXT(__FILE__), __LINE__, TEXT(repr), 0, (code)))
#define print_lasterr(repr) print_err(repr, GetLastError())

int atl_init();
int atl_deinit();

#endif
