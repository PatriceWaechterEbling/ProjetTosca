// (C) Patrice Waechter-Ebling 2022
//

#pragma once
#define WIN32_LEAN_AND_MEAN
#pragma comment(lib,"comctl32")


#define WIN32_LEAN_AND_MEAN 
#include <windows.h>
#include <crtdbg.h>
#include <tchar.h>
#include <commctrl.h>
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>
//#include <oleidl.h> 
#include <mshtml.h>
#include <mshtmhst.h>
#include "Interfaces.h"
#include "resource.h"


#define ProjetTosca ""

HWND  MainWindow;
HFONT  FontHandle;
HINSTANCE hMain;
static const TCHAR ClassName[] = "ProjetTosca";
static const TCHAR BrowserClassName[] = "ObjetProjetTosca";
static const SAFEARRAYBOUND ArrayBound = {1, 0};
char chaines[261];
wchar_t Blank[] = {L"about:blank"};
const TCHAR TextesHTML[] = {"<FONT COLOR=GREEN>This is page 1.</FONT><P><A HREF=\"app:2\">Click here.</A>\0\
This is page 2.<P>You can:<P><A HREF=\"app:1\">Go back to previous page.</A><BR><A HREF=\"app:3\">Go to next page.</A>\0\
<FONT COLOR=RED>This is page 3.</FONT><P>You can:<P><A HREF=\"app:2\">Go back to previous page.</A><BR><A HREF=\"app:1\">Go to first page.</A>\0\
\0"};

typedef struct _DOCHOSTUIINFO { ULONG cbSize;  DWORD dwFlags, dwDoubleClick; OLECHAR* pchHostCss, pchHostNS; }DOCHOSTUIINFO;
typedef struct _tagOLECMD { ULONG cmdID; DWORD cmdf; }OLECMD;
typedef struct _tagOLECMDTEXT { DWORD cmdtextf; ULONG cwActual, cwBuf; wchar_t rgwz[1]; }OLECMDTEXT;
typedef struct { IOleInPlaceSite inplace; _IOleInPlaceFrameEx  frame; } _IOleInPlaceSiteEx;
typedef struct { IDocHostUIHandler  ui; } _IDocHostUIHandlerEx;
typedef struct { IOleClientSite client; _IOleInPlaceSiteEx  inplace; _IDocHostUIHandlerEx ui; } _IOleClientSiteEx;
typedef struct { IOleInPlaceFrame frame; HWND window; } _IOleInPlaceFrameEx;

typedef struct IOleInPlaceFrameVtbl {
	BEGIN_INTERFACE
		virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleInPlaceFrame* This, __RPC__in REFIID riid, _COM_Outptr_  void** ppvObject);
	virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleInPlaceFrame* This);
	virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleInPlaceFrame* This);
	virtuelExp(IOleWindow, GetWindow)HRESULT(STDMETHODCALLTYPE* GetWindow)(__RPC__in IOleInPlaceFrame* This, __RPC__deref_out_opt HWND* phwnd);
	virtuelExp(IOleWindow, ContextSensitiveHelp)HRESULT(STDMETHODCALLTYPE* ContextSensitiveHelp)(__RPC__in IOleInPlaceFrame* This, BOOL fEnterMode);
	virtuelExp(IOleInPlaceUIWindow, GetBorder)HRESULT(STDMETHODCALLTYPE* GetBorder)(__RPC__in IOleInPlaceFrame* This, __RPC__out LPRECT lprectBorder);
	virtuelExp(IOleInPlaceUIWindow, RequestBorderSpace)HRESULT(STDMETHODCALLTYPE* RequestBorderSpace)(__RPC__in IOleInPlaceFrame* This, __RPC__in_opt LPCBORDERWIDTHS pborderwidths);
	virtuelExp(IOleInPlaceUIWindow, SetBorderSpace)HRESULT(STDMETHODCALLTYPE* SetBorderSpace)(__RPC__in IOleInPlaceFrame* This, __RPC__in_opt LPCBORDERWIDTHS pborderwidths);
	virtuelExp(IOleInPlaceUIWindow, SetActiveObject)HRESULT(STDMETHODCALLTYPE* SetActiveObject)(__RPC__in IOleInPlaceFrame* This, __RPC__in_opt IOleInPlaceActiveObject* pActiveObject, __RPC__in_opt_string LPCOLESTR pszObjName);
	virtuelExp(IOleInPlaceFrame, InsertMenus)HRESULT(STDMETHODCALLTYPE* InsertMenus)(__RPC__in IOleInPlaceFrame* This, __RPC__in HMENU hmenuShared, __RPC__inout LPOLEMENUGROUPWIDTHS lpMenuWidths);
	virtuelExp(IOleInPlaceFrame, SetMenu)HRESULT(STDMETHODCALLTYPE* SetMenu)(__RPC__in IOleInPlaceFrame* This, __RPC__in HMENU hmenuShared, __RPC__in HOLEMENU holemenu, __RPC__in HWND hwndActiveObject);
	virtuelExp(IOleInPlaceFrame, RemoveMenus)HRESULT(STDMETHODCALLTYPE* RemoveMenus)(__RPC__in IOleInPlaceFrame* This, __RPC__in HMENU hmenuShared);
	virtuelExp(IOleInPlaceFrame, SetStatusText)HRESULT(STDMETHODCALLTYPE* SetStatusText)(__RPC__in IOleInPlaceFrame* This, __RPC__in_opt LPCOLESTR pszStatusText);
	virtuelExp(IOleInPlaceFrame, EnableModeless)HRESULT(STDMETHODCALLTYPE* EnableModeless)(__RPC__in IOleInPlaceFrame* This, BOOL fEnable);
	virtuelExp(IOleInPlaceFrame, TranslateAccelerator)HRESULT(STDMETHODCALLTYPE* TranslateAccelerator)(__RPC__in IOleInPlaceFrame* This, __RPC__in LPMSG lpmsg, WORD wID);
	END_INTERFACE
} IOleInPlaceFrameVtbl;
typedef struct IOleClientSiteVtbl {
	BEGIN_INTERFACE
		virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleClientSite* This, __RPC__in REFIID riid, _COM_Outptr_  void** ppvObject);
	virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleClientSite* This);
	virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleClientSite* This);
	virtuelExp(IOleClientSite, SaveObject)HRESULT(STDMETHODCALLTYPE* SaveObject)(__RPC__in IOleClientSite* This);
	virtuelExp(IOleClientSite, GetMoniker)HRESULT(STDMETHODCALLTYPE* GetMoniker)(__RPC__in IOleClientSite* This, DWORD dwAssign, DWORD dwWhichMoniker, __RPC__deref_out_opt IMoniker** ppmk);
	virtuelExp(IOleClientSite, GetContainer)HRESULT(STDMETHODCALLTYPE* GetContainer)(__RPC__in IOleClientSite* This, __RPC__deref_out_opt IOleContainer** ppContainer);
	virtuelExp(IOleClientSite, ShowObject)HRESULT(STDMETHODCALLTYPE* ShowObject)(__RPC__in IOleClientSite* This);
	virtuelExp(IOleClientSite, OnShowWindow)HRESULT(STDMETHODCALLTYPE* OnShowWindow)(__RPC__in IOleClientSite* This, BOOL fShow);
	virtuelExp(IOleClientSite, RequestNewObjectLayout)HRESULT(STDMETHODCALLTYPE* RequestNewObjectLayout)(__RPC__in IOleClientSite* This);
	END_INTERFACE
} IOleClientSiteVtbl;
typedef struct IDocHostUIHandlerVtbl {
	BEGIN_INTERFACE
		virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(IDocHostUIHandler* This, REFIID riid, _COM_Outptr_  void** ppvObject);
	virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(IDocHostUIHandler* This);
	virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(IDocHostUIHandler* This);
	virtuelExp(IDocHostUIHandler, ShowContextMenu)HRESULT(STDMETHODCALLTYPE* ShowContextMenu)(IDocHostUIHandler* This, _In_  DWORD dwID, _In_  POINT* ppt, _In_  IUnknown* pcmdtReserved, _In_  IDispatch* pdispReserved);
	virtuelExp(IDocHostUIHandler, GetHostInfo)HRESULT(STDMETHODCALLTYPE* GetHostInfo)(IDocHostUIHandler* This, _Inout_  DOCHOSTUIINFO* pInfo);
	virtuelExp(IDocHostUIHandler, ShowUI)HRESULT(STDMETHODCALLTYPE* ShowUI)(IDocHostUIHandler* This, _In_  DWORD dwID, _In_  IOleInPlaceActiveObject* pActiveObject, _In_  IOleCommandTarget* pCommandTarget, _In_  IOleInPlaceFrame* pFrame, _In_  IOleInPlaceUIWindow* pDoc);
	virtuelExp(IDocHostUIHandler, HideUI)HRESULT(STDMETHODCALLTYPE* HideUI)(IDocHostUIHandler* This);
	virtuelExp(IDocHostUIHandler, UpdateUI)HRESULT(STDMETHODCALLTYPE* UpdateUI)(IDocHostUIHandler* This);
	virtuelExp(IDocHostUIHandler, EnableModeless)HRESULT(STDMETHODCALLTYPE* EnableModeless)(IDocHostUIHandler* This, BOOL fEnable);
	virtuelExp(IDocHostUIHandler, OnDocWindowActivate)HRESULT(STDMETHODCALLTYPE* OnDocWindowActivate)(IDocHostUIHandler* This, BOOL fActivate);
	virtuelExp(IDocHostUIHandler, OnFrameWindowActivate)HRESULT(STDMETHODCALLTYPE* OnFrameWindowActivate)(IDocHostUIHandler* This, BOOL fActivate);
	virtuelExp(IDocHostUIHandler, ResizeBorder)HRESULT(STDMETHODCALLTYPE* ResizeBorder)(IDocHostUIHandler* This, _In_  LPCRECT prcBorder, _In_  IOleInPlaceUIWindow* pUIWindow, _In_  BOOL fRameWindow);
	virtuelExp(IDocHostUIHandler, TranslateAccelerator)HRESULT(STDMETHODCALLTYPE* TranslateAccelerator)(IDocHostUIHandler* This, LPMSG lpMsg, const GUID* pguidCmdGroup, DWORD nCmdID);
	virtuelExp(IDocHostUIHandler, GetOptionKeyPath)HRESULT(STDMETHODCALLTYPE* GetOptionKeyPath)(IDocHostUIHandler* This, _Out_  LPOLESTR* pchKey, DWORD dw);
	virtuelExp(IDocHostUIHandler, GetDropTarget)HRESULT(STDMETHODCALLTYPE* GetDropTarget)(IDocHostUIHandler* This, _In_  IDropTarget* pDropTarget, _Outptr_  IDropTarget** ppDropTarget);
	virtuelExp(IDocHostUIHandler, GetExternal)HRESULT(STDMETHODCALLTYPE* GetExternal)(_Outptr_result_maybenull_  IDispatch** ppDispatch);
	virtuelExp(IDocHostUIHandler, TranslateUrl)HRESULT(STDMETHODCALLTYPE* TranslateUrl)(IDocHostUIHandler* This, DWORD dwTranslate, _In_  LPWSTR pchURLIn, _Outptr_  LPWSTR* ppchURLOut);
	virtuelExp(IDocHostUIHandler, FilterDataObject)HRESULT(STDMETHODCALLTYPE* FilterDataObject)(IDocHostUIHandler* This, _In_  IDataObject* pDO, _Outptr_result_maybenull_  IDataObject** ppDORet);
	END_INTERFACE
} IDocHostUIHandlerVtbl;
typedef struct IOleCommandTargetVtbl {
	BEGIN_INTERFACE
		virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleCommandTarget* This, __RPC__in REFIID riid, _COM_Outptr_  void** ppvObject);
	virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleCommandTarget* This);
	virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleCommandTarget* This);
	virtuelExp(IOleCommandTarget, QueryStatus)HRESULT(STDMETHODCALLTYPE* QueryStatus)(__RPC__in IOleCommandTarget* This, __RPC__in_opt const GUID* pguidCmdGroup, ULONG cCmds, __RPC__inout_ecount_full(cCmds) OLECMD prgCmds[], __RPC__inout_opt OLECMDTEXT* pCmdText);
	virtuelExp(IOleCommandTarget, Exec)HRESULT(STDMETHODCALLTYPE* Exec)(__RPC__in IOleCommandTarget* This, __RPC__in_opt const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, __RPC__in_opt VARIANT* pvaIn, __RPC__inout_opt VARIANT* pvaOut);
	END_INTERFACE
} IOleCommandTargetVtbl;
typedef struct IDocHostUIHandlerVtbl {
	BEGIN_INTERFACE
		virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(IDocHostUIHandler* This, REFIID riid, _COM_Outptr_  void** ppvObject);
	virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(IDocHostUIHandler* This);
	virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(IDocHostUIHandler* This);
	virtuelExp(IDocHostUIHandler, ShowContextMenu)HRESULT(STDMETHODCALLTYPE* ShowContextMenu)(IDocHostUIHandler* This, _In_  DWORD dwID, _In_  POINT* ppt, _In_  IUnknown* pcmdtReserved, _In_  IDispatch* pdispReserved);
	virtuelExp(IDocHostUIHandler, GetHostInfo)HRESULT(STDMETHODCALLTYPE* GetHostInfo)(IDocHostUIHandler* This, _Inout_  DOCHOSTUIINFO* pInfo);
	virtuelExp(IDocHostUIHandler, ShowUI)HRESULT(STDMETHODCALLTYPE* ShowUI)(IDocHostUIHandler* This, _In_  DWORD dwID, _In_  IOleInPlaceActiveObject* pActiveObject, _In_  IOleCommandTarget* pCommandTarget, _In_  IOleInPlaceFrame* pFrame, _In_  IOleInPlaceUIWindow* pDoc);
	virtuelExp(IDocHostUIHandler, HideUI)HRESULT(STDMETHODCALLTYPE* HideUI)(IDocHostUIHandler* This);
	virtuelExp(IDocHostUIHandler, UpdateUI)HRESULT(STDMETHODCALLTYPE* UpdateUI)(IDocHostUIHandler* This);
	virtuelExp(IDocHostUIHandler, EnableModeless)HRESULT(STDMETHODCALLTYPE* EnableModeless)(IDocHostUIHandler* This, BOOL fEnable);
	virtuelExp(IDocHostUIHandler, OnDocWindowActivate)HRESULT(STDMETHODCALLTYPE* OnDocWindowActivate)(IDocHostUIHandler* This, BOOL fActivate);
	virtuelExp(IDocHostUIHandler, OnFrameWindowActivate)HRESULT(STDMETHODCALLTYPE* OnFrameWindowActivate)(IDocHostUIHandler* This, BOOL fActivate);
	virtuelExp(IDocHostUIHandler, ResizeBorder)HRESULT(STDMETHODCALLTYPE* ResizeBorder)(IDocHostUIHandler* This, _In_  LPCRECT prcBorder, _In_  IOleInPlaceUIWindow* pUIWindow, _In_  BOOL fRameWindow);
	virtuelExp(IDocHostUIHandler, TranslateAccelerator)HRESULT(STDMETHODCALLTYPE* TranslateAccelerator)(IDocHostUIHandler* This, LPMSG lpMsg, const GUID* pguidCmdGroup, DWORD nCmdID);
	virtuelExp(IDocHostUIHandler, GetOptionKeyPath)HRESULT(STDMETHODCALLTYPE* GetOptionKeyPath)(IDocHostUIHandler* This, _Out_  LPOLESTR* pchKey, WORD dw);
	virtuelExp(IDocHostUIHandler, GetDropTarget)HRESULT(STDMETHODCALLTYPE* GetDropTarget)(IDocHostUIHandler* This, _In_  IDropTarget* pDropTarget, _Outptr_  IDropTarget** ppDropTarget);
	virtuelExp(IDocHostUIHandler, GetExternal)HRESULT(STDMETHODCALLTYPE* GetExternal)(IDocHostUIHandler* This, _Outptr_result_maybenull_  IDispatch** ppDispatch);
	virtuelExp(IDocHostUIHandler, TranslateUrl)HRESULT(STDMETHODCALLTYPE* TranslateUrl)(IDocHostUIHandler* This, DWORD dwTranslate, _In_  LPWSTR pchURLIn, _Outptr_  LPWSTR* ppchURLOut);
	virtuelExp(IDocHostUIHandler, FilterDataObject)HRESULT(STDMETHODCALLTYPE* FilterDataObject)(IDocHostUIHandler* This, _In_  IDataObject* pDO, _Outptr_result_maybenull_  IDataObject** ppDORet);
	END_INTERFACE
} IDocHostUIHandlerVtbl;
typedef struct IOleInPlaceSiteVtbl {
	BEGIN_INTERFACE
		virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleInPlaceSite* This, __RPC__in REFIID riid, _COM_Outptr_  void** ppvObject);
	virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleWindow, GetWindow)HRESULT(STDMETHODCALLTYPE* GetWindow)(__RPC__in IOleInPlaceSite* This, __RPC__deref_out_opt HWND* phwnd);
	virtuelExp(IOleWindow, ContextSensitiveHelp)HRESULT(STDMETHODCALLTYPE* ContextSensitiveHelp)(__RPC__in IOleInPlaceSite* This, BOOL fEnterMode);
	virtuelExp(IOleInPlaceSite, CanInPlaceActivate)HRESULT(STDMETHODCALLTYPE* CanInPlaceActivate)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleInPlaceSite, OnInPlaceActivate)HRESULT(STDMETHODCALLTYPE* OnInPlaceActivate)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleInPlaceSite, OnUIActivate)HRESULT(STDMETHODCALLTYPE* OnUIActivate)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleInPlaceSite, GetWindowContext)HRESULT(STDMETHODCALLTYPE* GetWindowContext)(__RPC__in IOleInPlaceSite* This, __RPC__deref_out_opt IOleInPlaceFrame** ppFrame, __RPC__deref_out_opt IOleInPlaceUIWindow** ppDoc, __RPC__out LPRECT lprcPosRect, __RPC__out LPRECT lprcClipRect, __RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo);
	virtuelExp(IOleInPlaceSite, Scroll)HRESULT(STDMETHODCALLTYPE* Scroll)(__RPC__in IOleInPlaceSite* This, SIZE scrollExtant);
	virtuelExp(IOleInPlaceSite, OnUIDeactivate)HRESULT(STDMETHODCALLTYPE* OnUIDeactivate)(__RPC__in IOleInPlaceSite* This, BOOL fUndoable);
	virtuelExp(IOleInPlaceSite, OnInPlaceDeactivate)HRESULT(STDMETHODCALLTYPE* OnInPlaceDeactivate)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleInPlaceSite, DiscardUndoState)HRESULT(STDMETHODCALLTYPE* DiscardUndoState)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleInPlaceSite, DeactivateAndUndo)HRESULT(STDMETHODCALLTYPE* DeactivateAndUndo)(__RPC__in IOleInPlaceSite* This);
	virtuelExp(IOleInPlaceSite, OnPosRectChange)HRESULT(STDMETHODCALLTYPE* OnPosRectChange)(__RPC__in IOleInPlaceSite* This, __RPC__in LPCRECT lprcPosRect);
	END_INTERFACE
} IOleInPlaceSiteVtbl;

interface IDocHostUIHandler { CONST_VTBL struct IDocHostUIHandlerVtbl* lpVtbl; };
interface IOleClientSite { CONST_VTBL struct IOleClientSiteVtbl* lpVtbl; };
interface IOleInPlaceFrame { CONST_VTBL struct IOleInPlaceFrameVtbl* lpVtbl; };
interface IOleCommandTarget { CONST_VTBL struct IOleCommandTargetVtbl* lpVtbl; };
interface IOleInPlaceSite { CONST_VTBL struct IOleInPlaceSiteVtbl* lpVtbl; };

HRESULT STDMETHODCALLTYPE UI_QueryInterface(IDocHostUIHandler* This, REFIID riid, LPVOID* ppvObj) { return(Site_QueryInterface((IOleClientSite*)((char*)This - sizeof(IOleClientSite) - sizeof(_IOleInPlaceSiteEx)), riid, ppvObj)); }
HRESULT STDMETHODCALLTYPE UI_AddRef(IDocHostUIHandler* This) { return(1); }
HRESULT STDMETHODCALLTYPE UI_Release(IDocHostUIHandler* This) { return(1); }
HRESULT STDMETHODCALLTYPE UI_ShowContextMenu(IDocHostUIHandler* This, DWORD dwID, POINT* ppt, IUnknown* pcmdtReserved, IDispatch* pdispReserved) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_GetHostInfo(IDocHostUIHandler* This, DOCHOSTUIINFO* pInfo) { pInfo->cbSize = sizeof(DOCHOSTUIINFO); pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER; pInfo->dwDoubleClick = 0; return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_ShowUI(IDocHostUIHandler* This, DWORD dwID, IOleInPlaceActiveObject* pActiveObject, IOleCommandTarget* pCommandTarget, IOleInPlaceFrame* pFrame, IOleInPlaceUIWindow* pDoc) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_HideUI(IDocHostUIHandler* This) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_UpdateUI(IDocHostUIHandler* This) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_EnableModeless(IDocHostUIHandler* This, BOOL fEnable) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_OnDocWindowActivate(IDocHostUIHandler* This, BOOL fActivate) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_OnFrameWindowActivate(IDocHostUIHandler* This, BOOL fActivate) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_ResizeBorder(IDocHostUIHandler* This, LPCRECT prcBorder, IOleInPlaceUIWindow* pUIWindow, BOOL fRameWindow) { return(S_OK); }
HRESULT STDMETHODCALLTYPE UI_TranslateAccelerator(IDocHostUIHandler* This, LPMSG lpMsg, const GUID* pguidCmdGroup, DWORD nCmdID) { return(S_FALSE); }
HRESULT STDMETHODCALLTYPE UI_GetOptionKeyPath(IDocHostUIHandler* This, LPOLESTR* pchKey, DWORD dw) { return(S_FALSE); }
HRESULT STDMETHODCALLTYPE UI_GetDropTarget(IDocHostUIHandler* This, IDropTarget* pDropTarget, IDropTarget** ppDropTarget) { return(S_FALSE); }
HRESULT STDMETHODCALLTYPE UI_GetExternal(IDocHostUIHandler* This, IDispatch** ppDispatch) { *ppDispatch = 0; return(S_FALSE); }
HRESULT STDMETHODCALLTYPE UI_TranslateUrl(IDocHostUIHandler* This, DWORD dwTranslate, OLECHAR* pchURLIn, OLECHAR** ppchURLOut) {
	unsigned short* src;
	unsigned short* dest;
	DWORD len;
	src = pchURLIn;
	while ((*(src)++));
	--src;
	len = src - pchURLIn;
	if (len >= 4 && !_wcsnicmp(pchURLIn, (wchar_t*)&AppUrl[0], 4)) {
		if ((dest = (OLECHAR*)CoTaskMemAlloc(12 << 1))) {
			HWND hwnd;
			*ppchURLOut = dest;
			CopyMemory(dest, &Blank[0], 12 << 1);
			len = Asc2Int(pchURLIn + 4);
			hwnd = ((_IOleInPlaceSiteEx*)((char*)This - sizeof(_IOleInPlaceSiteEx)))->frame.window;
			PostMessage(hwnd, WM_APP + 50, (WPARAM)len, 0);
			return(S_OK);
		}
	}
	*ppchURLOut = 0;
	return(S_FALSE);
}
HRESULT STDMETHODCALLTYPE UI_FilterDataObject(IDocHostUIHandler* This, IDataObject* pDO, IDataObject** ppDORet) { *ppDORet = 0; return(S_FALSE); }
HRESULT STDMETHODCALLTYPE Site_QueryInterface(IOleClientSite* This, REFIID riid, void** ppvObject) {
	if (!memcmp(riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(riid, &IID_IOleClientSite, sizeof(GUID)))  *ppvObject = &((_IOleClientSiteEx*)This)->client;
	else if (!memcmp((LPCVOID)riid, &IID_IOleInPlaceSite, sizeof(GUID)))  *ppvObject = &((_IOleClientSiteEx*)This)->inplace;
	else if (!memcmp(riid, &IID_IDocHostUIHandler, sizeof(GUID))) *ppvObject = &((_IOleClientSiteEx*)This)->ui;
	else { *ppvObject = 0;  return(E_NOINTERFACE); }
	return(S_OK);
}
HRESULT STDMETHODCALLTYPE Site_AddRef(IOleClientSite* This) { return(1); }
HRESULT STDMETHODCALLTYPE Site_Release(IOleClientSite* This) { return(1); }
HRESULT STDMETHODCALLTYPE Site_SaveObject(IOleClientSite* This) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Site_GetMoniker(IOleClientSite* This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Site_GetContainer(IOleClientSite* This, LPOLECONTAINER* ppContainer) { *ppContainer = 0; return(E_NOINTERFACE); }
HRESULT STDMETHODCALLTYPE Site_ShowObject(IOleClientSite* This) { return(NOERROR); }
HRESULT STDMETHODCALLTYPE Site_OnShowWindow(IOleClientSite* This, BOOL fShow) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Site_RequestNewObjectLayout(IOleClientSite* This) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE InPlace_QueryInterface(IOleInPlaceSite* This, REFIID riid, LPVOID* ppvObj) { return(Site_QueryInterface((IOleClientSite*)((char*)This - sizeof(IOleClientSite)), riid, ppvObj)); }
HRESULT STDMETHODCALLTYPE InPlace_AddRef(IOleInPlaceSite* This) { return(1); }
HRESULT STDMETHODCALLTYPE InPlace_Release(IOleInPlaceSite* This) { return(1); }
HRESULT STDMETHODCALLTYPE InPlace_GetWindow(IOleInPlaceSite* This, HWND* lphwnd) { *lphwnd = ((_IOleInPlaceSiteEx*)This)->frame.window; return(S_OK); }
HRESULT STDMETHODCALLTYPE InPlace_ContextSensitiveHelp(IOleInPlaceSite* This, BOOL fEnterMode) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE InPlace_CanInPlaceActivate(IOleInPlaceSite* This) { return(S_OK); }
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceActivate(IOleInPlaceSite* This) { return(S_OK); }
HRESULT STDMETHODCALLTYPE InPlace_OnUIActivate(IOleInPlaceSite* This) { return(S_OK); }
HRESULT STDMETHODCALLTYPE InPlace_GetWindowContext(IOleInPlaceSite* This, LPOLEINPLACEFRAME* lplpFrame, LPOLEINPLACEUIWINDOW* lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo) {
	*lplpFrame = (LPOLEINPLACEFRAME) & ((_IOleInPlaceSiteEx*)This)->frame;
	*lplpDoc = 0;
	lpFrameInfo->fMDIApp = FALSE;
	lpFrameInfo->hwndFrame = ((_IOleInPlaceFrameEx*)*lplpFrame)->window;
	lpFrameInfo->haccel = 0;
	lpFrameInfo->cAccelEntries = 0;
	return(S_OK);
}
HRESULT STDMETHODCALLTYPE InPlace_Scroll(IOleInPlaceSite* This, SIZE scrollExtent) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE InPlace_OnUIDeactivate(IOleInPlaceSite* This, BOOL fUndoable) { return(S_OK); }
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceDeactivate(IOleInPlaceSite* This) { return(S_OK); }
HRESULT STDMETHODCALLTYPE InPlace_DiscardUndoState(IOleInPlaceSite* This) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE InPlace_DeactivateAndUndo(IOleInPlaceSite* This) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_QueryInterface(IOleInPlaceFrame* This, REFIID riid, LPVOID* ppvObj) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_AddRef(IOleInPlaceFrame* This) { return(1); }
HRESULT STDMETHODCALLTYPE Frame_Release(IOleInPlaceFrame* This) { return(1); }
HRESULT STDMETHODCALLTYPE Frame_GetWindow(IOleInPlaceFrame* This, HWND* lphwnd) { *lphwnd = ((_IOleInPlaceFrameEx*)This)->window; return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_ContextSensitiveHelp(IOleInPlaceFrame* This, BOOL fEnterMode) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_GetBorder(IOleInPlaceFrame* This, LPRECT lprectBorder) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_RequestBorderSpace(IOleInPlaceFrame* This, LPCBORDERWIDTHS pborderwidths) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_SetBorderSpace(IOleInPlaceFrame* This, LPCBORDERWIDTHS pborderwidths) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_SetActiveObject(IOleInPlaceFrame* This, IOleInPlaceActiveObject* pActiveObject, LPCOLESTR pszObjName) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_InsertMenus(IOleInPlaceFrame* This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_SetMenu(IOleInPlaceFrame* This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_RemoveMenus(IOleInPlaceFrame* This, HMENU hmenuShared) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE Frame_SetStatusText(IOleInPlaceFrame* This, LPCOLESTR pszStatusText) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_EnableModeless(IOleInPlaceFrame* This, BOOL fEnable) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_TranslateAccelerator(IOleInPlaceFrame* This, LPMSG lpmsg, WORD wID) { return(E_NOTIMPL); }
HRESULT STDMETHODCALLTYPE InPlace_OnPosRectChange(IOleInPlaceSite* This, LPCRECT lprcPosRect) {
	IOleObject* browserObject;
	IOleInPlaceObject* inplace = 0;
	browserObject = *((IOleObject**)((char*)This - sizeof(IOleObject*) - sizeof(IOleClientSite)));
	if (!browserObject->QueryInterface(browserObject, &IID_IOleInPlaceObject, (void**)&inplace)) { inplace->SetObjectRects(inplace, lprcPosRect, lprcPosRect); }
	return(S_OK);
}
IOleInPlaceFrameVtbl MyIOleInPlaceFrameTable = {Frame_QueryInterface,Frame_AddRef,Frame_Release,Frame_GetWindow,Frame_ContextSensitiveHelp,Frame_GetBorder,Frame_RequestBorderSpace,Frame_SetBorderSpace,Frame_SetActiveObject,Frame_InsertMenus,Frame_SetMenu,Frame_RemoveMenus,Frame_SetStatusText,Frame_EnableModeless,Frame_TranslateAccelerator};
IOleClientSiteVtbl MyIOleClientSiteTable = {Site_QueryInterface,Site_AddRef,Site_Release,Site_SaveObject,Site_GetMoniker,Site_GetContainer,Site_ShowObject,Site_OnShowWindow,Site_RequestNewObjectLayout};

void AjouteImagesMenu(int Logo,int Index);
void ChargementChaine(int ID);
DWORD Asc2Int(OLECHAR *val){
	OLECHAR chr;
	DWORD len;
	len = 0;
	while (*val == ' ' || *val == 0x09) val++;
	while (*val) {
		chr = *(val)++ - '0';
		if ((DWORD)chr > 9) break;
		len += (len + (len << 3) + chr);
	}
	return(len);
}
void MasquerObjet(HWND hwnd){
	IOleObject **browserHandle;
	IOleObject *browserObject;
	if ((browserHandle = (IOleObject **)GetWindowLong(hwnd, GWL_USERDATA))) {
		browserObject = *browserHandle;
		browserObject->lpVtbl->Close(browserObject, OLECLOSE_NOSAVE);
		browserObject->lpVtbl->Release(browserObject);
		SetWindowLong(hwnd, GWL_USERDATA, 0);
	}
}
void ActionsObjet(HWND hwnd, DWORD action){
	IWebBrowser2 *webBrowser2;
	IOleObject  *browserObject;
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));
	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2)) {
		switch (action){
			case GOBACK: { webBrowser2->lpVtbl->GoBack(webBrowser2);  break; }
			case GOFORWARD: { webBrowser2->lpVtbl->GoForward(webBrowser2); break; }
			case GOHOME: { webBrowser2->lpVtbl->GoHome(webBrowser2);  break; }
			case SEARCH:  { webBrowser2->lpVtbl->GoSearch(webBrowser2);  break; }
			case REFRESH:  { webBrowser2->lpVtbl->Refresh(webBrowser2); }
			case STOP: { webBrowser2->lpVtbl->Stop(webBrowser2); }
		}
		webBrowser2->lpVtbl->Release(webBrowser2);
	}
}
long AfficheTextesHTML(HWND hwnd, LPCTSTR string){
	IWebBrowser2 *webBrowser2;
	LPDISPATCH  lpDispatch;
	IHTMLDocument2 *htmlDoc2;
	IOleObject  *browserObject;
	SAFEARRAY  *sfArray;
	VARIANT myURL;
	VARIANT *pVar;
	BSTR bstr;
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));
	bstr = 0;
	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2)) {
		lpDispatch = 0;
		webBrowser2->lpVtbl->get_Document(webBrowser2, &lpDispatch);
		if (!lpDispatch)  {
			VariantInit(&myURL);
			myURL.vt = VT_BSTR;
			myURL.bstrVal = SysAllocString(&Blank[0]);
			webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, 0, 0, 0, 0);
			VariantClear(&myURL);
		}
		if (!webBrowser2->lpVtbl->get_Document(webBrowser2, &lpDispatch) && lpDispatch)  {
			if (!lpDispatch->lpVtbl->QueryInterface(lpDispatch, &IID_IHTMLDocument2, (void**)&htmlDoc2)) {
				if ((sfArray = SafeArrayCreate(VT_VARIANT, 1, (SAFEARRAYBOUND *)&ArrayBound))){
					if (!SafeArrayAccessData(sfArray, (void**)&pVar)){
						pVar->vt = VT_BSTR;
						#ifndef UNICODE
						{
							wchar_t  *buffer;
							DWORD  size;
							size = MultiByteToWideChar(CP_ACP, 0, string, -1, 0, 0);
							if (!(buffer = (wchar_t *)GlobalAlloc(GMEM_FIXED, sizeof(wchar_t) * size))) goto bad;
							MultiByteToWideChar(CP_ACP, 0, string, -1, buffer, size);
							bstr = SysAllocString(buffer);
							GlobalFree(buffer);
						}
						#else
							bstr = SysAllocString(string);
						#endif
						if ((pVar->bstrVal = bstr)) {htmlDoc2->lpVtbl->write(htmlDoc2, sfArray); htmlDoc2->lpVtbl->close(htmlDoc2);}
					}
					SafeArrayDestroy(sfArray);
				}
				bad: htmlDoc2->lpVtbl->Release(htmlDoc2);
			}
			lpDispatch->lpVtbl->Release(lpDispatch);
		}
		webBrowser2->lpVtbl->Release(webBrowser2);
	}
	if (bstr) return(0);
	return(-1);
}
long AfficherPageHTML(HWND hwnd, LPTSTR webPageName)
{
	IWebBrowser2 *webBrowser2;
	VARIANT myURL;
	IOleObject  *browserObject;
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));
	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2))
	{
		VariantInit(&myURL);
		myURL.vt = VT_BSTR;
	}
#ifndef UNICODE
{
	wchar_t  *buffer;
	DWORD  size;
	size = MultiByteToWideChar(CP_ACP, 0, webPageName, -1, 0, 0);
	if (!(buffer = (wchar_t *)GlobalAlloc(GMEM_FIXED, sizeof(wchar_t) * size))) goto badalloc;
	MultiByteToWideChar(CP_ACP, 0, webPageName, -1, buffer, size);
	myURL.bstrVal = SysAllocString(buffer);
	GlobalFree(buffer);
}
#else
	myURL.bstrVal = SysAllocString(webPageName);
#endif
	{
		if (!myURL.bstrVal)	{	badalloc: webBrowser2->lpVtbl->Release(webBrowser2);	return(-6);	}
		webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, 0, 0, 0, 0);
		VariantClear(&myURL);
		webBrowser2->lpVtbl->Release(webBrowser2);
		return(0);
	}
	return(-5);
}
void RedimensionnerObjet(HWND hwnd, DWORD width, DWORD height)
{
	IWebBrowser2 *webBrowser2;
	IOleObject  *browserObject;
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));
	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2)) {
		webBrowser2->lpVtbl->put_Top(webBrowser2, 0);
		webBrowser2->lpVtbl->put_Left(webBrowser2, 0);
		webBrowser2->lpVtbl->put_Width(webBrowser2, width);
		webBrowser2->lpVtbl->put_Height(webBrowser2, height);
		webBrowser2->lpVtbl->Release(webBrowser2);
	}
}
long AfficherObjet(HWND hwnd)
{
	LPCLASSFACTORY  pClassFactory;
	IOleObject *browserObject;
	IWebBrowser2  *webBrowser2;
	RECT  rect;
	char  *ptr;
	_IOleClientSiteEx *_iOleClientSiteEx;
	if (!(ptr = (char *)GlobalAlloc(GMEM_FIXED, sizeof(_IOleClientSiteEx) + sizeof(IOleObject *))))  return(-1);
	_iOleClientSiteEx = (_IOleClientSiteEx *)(ptr + sizeof(IOleObject *));
	_iOleClientSiteEx->client.lpVtbl = &MyIOleClientSiteTable;
	_iOleClientSiteEx->inplace.inplace.lpVtbl = &MyIOleInPlaceSiteTable;
	_iOleClientSiteEx->inplace.frame.frame.lpVtbl = &MyIOleInPlaceFrameTable;
	_iOleClientSiteEx->inplace.frame.window = hwnd;
	_iOleClientSiteEx->ui.ui.lpVtbl = &MyIDocHostUIHandlerTable;
	pClassFactory = 0;
	if (!CoGetClassObject(&CLSID_WebBrowser, 3, NULL, &IID_IClassFactory, (void **)&pClassFactory) && pClassFactory) {
		if (!pClassFactory->lpVtbl->CreateInstance(pClassFactory, 0, &IID_IOleObject, &browserObject)){
			pClassFactory->lpVtbl->Release(pClassFactory);
			*((IOleObject **)ptr) = browserObject;
			SetWindowLong(hwnd, GWL_USERDATA, (LONG)ptr);
			if (!browserObject->lpVtbl->SetClientSite(browserObject, (IOleClientSite *)_iOleClientSiteEx))	{
				browserObject->lpVtbl->SetHostNames(browserObject, L"My Host Name", 0);
				GetClientRect(hwnd, &rect);
				if (!OleSetContainedObject((struct IUnknown *)browserObject, TRUE) && !browserObject->lpVtbl->DoVerb(browserObject, OLEIVERB_SHOW, NULL, (IOleClientSite *)_iOleClientSiteEx, -1, hwnd, &rect) &&  !browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2))  {
					webBrowser2->lpVtbl->put_Left(webBrowser2, 0);
					webBrowser2->lpVtbl->put_Top(webBrowser2, 0);
					webBrowser2->lpVtbl->put_Width(webBrowser2, rect.right);
					webBrowser2->lpVtbl->put_Height(webBrowser2, rect.bottom);
					webBrowser2->lpVtbl->Release(webBrowser2);
					return(0);
				}
			}
			MasquerObjet(hwnd); return(-4);
		}
		pClassFactory->lpVtbl->Release(pClassFactory);
		GlobalFree(ptr);  return(-3);
	}
	GlobalFree(ptr); return(-2);
}
void CreerPoliceObjet(void){
	NONCLIENTMETRICS nonClientMetrics;
	nonClientMetrics.cbSize = sizeof(NONCLIENTMETRICS);
	SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, &nonClientMetrics, 0);
	FontHandle = CreateFontIndirect(&nonClientMetrics.lfStatusFont);
}
LRESULT CALLBACK ProcedurePrincipale(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam){
	switch (uMsg) {
	  case WM_SIZE:  {MoveWindow(GetDlgItem(hwnd, 1000), 0, 0, LOWORD(lParam), HIWORD(lParam)-21, TRUE);return(0);}
	  case WM_COMMAND: {
		HWND browserHwnd = GetDlgItem(hwnd, 1000);
		if(LOWORD(wParam)>1007){SetDlgItemInt(hwnd,6000,LOWORD(wParam),0);
		if(LOWORD(wParam)>60000){ChargementChaine(LOWORD(wParam));}
		}
		switch (LOWORD(wParam)) {
			case 1001:{ActionsObjet(browserHwnd, GOBACK); break; }
			case 1002:{ActionsObjet(browserHwnd, GOFORWARD);break; }
			case 1003:{ActionsObjet(browserHwnd, STOP);  break; }
			case 1004:{AfficherPageHTML(browserHwnd, UrlTosca);  break; }
			case 1005:{ActionsObjet(browserHwnd, REFRESH);  break; }
			case 1006:{ActionsObjet(browserHwnd, STOP);  break; }
		} break;
	}
	case WM_CREATE: { return(0);  }
	case WM_ERASEBKGND:  {
		HBRUSH hBrush;
		HGDIOBJ orig;
		RECT rec;
		hBrush = GetStockObject(WHITE_BRUSH);
		GetClientRect(hwnd, &rec);
		orig = SelectObject((HDC)wParam, hBrush);
		Rectangle((HDC)wParam, 0, 0, rec.right, 40);
		SelectObject((HDC)wParam, orig);
		return(TRUE);
	}
	case WM_GETMINMAXINFO: {
		LPMINMAXINFO  lpmmi;
		NONCLIENTMETRICS nonClientMetrics;
		lpmmi = (LPMINMAXINFO)lParam;
		nonClientMetrics.cbSize = sizeof(NONCLIENTMETRICS);
		SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, &nonClientMetrics, 0);
		lpmmi->ptMinTrackSize.y = nonClientMetrics.iCaptionHeight + (nonClientMetrics.iBorderWidth << 1) + 80;
		}return(0);
	case WM_CLOSE:  { DestroyWindow(hwnd); return(1);  }
	case WM_DESTROY:  { PostQuitMessage(0); return(1);  }
	}
	return(DefWindowProc(hwnd, uMsg, wParam, lParam));
}
void AjouteImagesMenu(int Logo,int Index){ HBITMAP bmp=LoadBitmap(hMain,(LPCTSTR)Logo); SetMenuItemBitmaps( GetSubMenu(GetMenu(MainWindow),1),Index,MF_BYCOMMAND  ,bmp,bmp);}
LRESULT CALLBACK ProcedureObjet(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam){
	switch (uMsg) {
		case WM_SIZE:{  RedimensionnerObjet(hwnd, LOWORD(lParam), HIWORD(lParam));  return(0);}
		case WM_APP:{
			LPCTSTR  str;
			str = &TextesHTML[0];
			while (--wParam){if (!(*str)) return(0);  while (*(str)++); }
			AfficheTextesHTML(hwnd, str);
			return(0);
		}
		case WM_CREATE:{if (AfficherObjet(hwnd)) return(-1); return(0);  }
		case WM_DESTROY:{MasquerObjet(hwnd); return(TRUE);  }
	}
	return(DefWindowProc(hwnd, uMsg, wParam, lParam));
}
void ChargementChaine(int ID){
	char url[260],clef[8];
	wsprintf(clef,"%d",ID);
	GetPrivateProfileString("ID",clef,"",url,sizeof(url),chaines);
	SetDlgItemText(MainWindow,6000,url);
	AfficherPageHTML(GetDlgItem(MainWindow, 1000),url);
}
int CALLBACK WinMain(HINSTANCE hInstance, HINSTANCE hInstNULL, LPSTR lpszCmdLine, int nCmdShow)
{
	SYSTEMTIME st;
	MSG msg;
	RECT rc;
	int a=0;
	hMain=hInstance;
	GetClientRect(GetDesktopWindow(),&rc);
	GetSystemTime(&st);
	if (OleInitialize(NULL) == S_OK){
		WNDCLASSEX  wc;
		ZeroMemory(&wc, sizeof(WNDCLASSEX));
		wc.cbSize = sizeof(WNDCLASSEX);
		wc.hInstance = hInstance;
		wc.lpfnWndProc = ProcedurePrincipale;
		wc.lpszClassName = &ClassName[0];
		wc.style = CS_CLASSDC|CS_HREDRAW|CS_VREDRAW|CS_PARENTDC|CS_BYTEALIGNCLIENT|CS_DBLCLKS;
		wc.hIcon=LoadIcon(hInstance,(LPCTSTR)100);
		wc.lpszMenuName=(LPCTSTR)100;
		RegisterClassEx(&wc);
		wc.lpfnWndProc = ProcedureObjet;
		wc.lpszClassName = &BrowserClassName[0];
		wc.style = CS_HREDRAW|CS_VREDRAW;
		RegisterClassEx(&wc);
		CreerPoliceObjet();
		if ((MainWindow = CreateWindowEx(WS_EX_TRANSPARENT, &ClassName[0], "ProjetTosca", WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU| WS_THICKFRAME,0, 0, rc.right,rc.bottom-50,GetDesktopWindow(), NULL, hInstance, 0))){
			if ((msg.hwnd = CreateWindowEx(WS_EX_DLGMODALFRAME|WS_EX_CLIENTEDGE, &BrowserClassName[0], 0, WS_CHILD|WS_VISIBLE,0, 0, 0, 0,MainWindow, (HMENU)1000, hInstance, 0))) {
				CreateStatusWindow(WS_VISIBLE|WS_CHILD|WS_DLGFRAME|WS_DISABLED,url,MainWindow,6000);
				AfficherPageHTML(msg.hwnd,ProjetTosca);
				ShowWindow(MainWindow, nCmdShow);
				for(a=0; a<26; a++){AjouteImagesMenu(200+a,64000+a);}
				UpdateWindow(MainWindow);
				while (GetMessage(&msg, 0, 0, 0) == 1)  {
				TranslateMessage(&msg);
				DispatchMessage(&msg);
				}
			}else{
				MessageBox(MainWindow, "Échec a la création de l'objet", "Erreur", MB_OK);
				DestroyWindow(MainWindow);
			}
		}
		if (FontHandle) DeleteObject(FontHandle);OleUninitialize();return(0);
	}
	MessageBox(0, "Plantage OLE!", "Erreur OLE", MB_OK);return(-1);
}