#ifndef LUKE1337_HWPCTRL_H
#define LUKE1337_HWPCTRL_H
#include "common.h"

/*** Codes and interfaces specific to Hancom Office Hanword ActiveX control */

// static CLSID CLSID_HwpCtrl = {0xBD9C32DE, 0x3155, 0x4691, {0x89, 0x72, 0x09, 0x7D, 0x53, 0xB1, 0x00, 0x52}};

#define DIDX_HWPCTRL_REGISTER_MODULE 0
#define DIDX_HWPCTRL_OPEN 1
#define DIDX_HWPCTRL_SAVE 2
#define DIDX_HWPCTRL_SAVE_AS 3
#define DIDX_HWPCTRL_CLEAR 4
#define DIDX_HWPCTRL_SHOW_TOOL_BAR 5
#define DIDX_HWPCTRL_NR 6

typedef DISPID (DIDX_HWPCTRL[DIDX_HWPCTRL_NR]);

inline static HRESULT HwpCtrl_RegisterModule(IDispatch *obj, DISPID *table, VARIANT *result, LPCOLESTR ModuleType, LPCOLESTR ModuleData) {
	HRESULT hres;
	VARIANTARG vargs[2];
	SET_ARG_BSTR(vargs, 0, ModuleData); SET_ARG_BSTR(vargs, 1, ModuleType);
	hres = ole_invoke_method(obj, table[DIDX_HWPCTRL_REGISTER_MODULE], result, vargs, 2);
	FREE_VARSTR(vargs[0]); FREE_VARSTR(vargs[1]);
	return hres;
}

inline static HRESULT HwpCtrl_Open(IDispatch *obj, DISPID *table, VARIANT *result, LPCOLESTR Path, LPCOLESTR format, LPCOLESTR arg) {
	HRESULT hres;
	VARIANTARG vargs[3];
	SET_ARG_BSTR(vargs, 0, arg); SET_ARG_BSTR(vargs, 1, format); SET_ARG_BSTR(vargs, 2, Path);
	hres = ole_invoke_method(obj, table[DIDX_HWPCTRL_OPEN], result, vargs, 3);
	FREE_VARSTR(vargs[0]); FREE_VARSTR(vargs[1]); FREE_VARSTR(vargs[2]);
	return hres;
}

inline static HRESULT HwpCtrl_Save(IDispatch *obj, DISPID *table, VARIANT *result, VARIANT_BOOL save_if_dirty) {
	HRESULT hres;
	VARIANTARG vargs[1];
	vargs[0].vt = VT_BOOL;
	vargs[0].boolVal = save_if_dirty;
	hres = ole_invoke_method(obj, table[DIDX_HWPCTRL_SAVE], result, vargs, 1);
	return hres;
}

inline static HRESULT HwpCtrl_SaveAs(IDispatch *obj, DISPID *table, VARIANT *result, LPCOLESTR Path, LPCOLESTR format, LPCOLESTR arg) {
	HRESULT hres;
	VARIANTARG vargs[3];
	SET_ARG_BSTR(vargs, 0, arg); SET_ARG_BSTR(vargs, 1, format); SET_ARG_BSTR(vargs, 2, Path);
	hres = ole_invoke_method(obj, table[DIDX_HWPCTRL_SAVE_AS], result, vargs, 3);
	FREE_VARSTR(vargs[0]); FREE_VARSTR(vargs[1]); FREE_VARSTR(vargs[2]);
	return hres;
}

#define hwpAskSave 0
#define hwpDiscard 1
#define hwpSaveIfDirty 2
#define hwpSave 3

inline static HRESULT HwpCtrl_Clear(IDispatch *obj, DISPID *table, VARIANT *result, LONG option) {
	HRESULT hres;
	VARIANTARG vargs[1];
	vargs[0].vt = VT_I4;
	vargs[0].iVal = option;
	hres = ole_invoke_method(obj, table[DIDX_HWPCTRL_CLEAR], result, vargs, 1);
	return hres;
}

inline static HRESULT HwpCtrl_ShowToolBar(IDispatch *obj, DISPID *table, VARIANT *result, LONG bShow) {
	HRESULT hres;
	VARIANTARG vargs[1];
	vargs[0].vt = VT_I4;
	vargs[0].iVal = bShow;
	hres = ole_invoke_method(obj, table[DIDX_HWPCTRL_SAVE_AS], result, vargs, 1);
	return hres;
}

#endif
