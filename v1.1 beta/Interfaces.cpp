#include <windows.h>
#include "Interfaces.h"
typedef struct _DOCHOSTUIINFO{ULONG cbSize;  DWORD dwFlags,dwDoubleClick;OLECHAR *pchHostCss,pchHostNS;}DOCHOSTUIINFO;
typedef struct _tagOLECMD{ULONG cmdID;DWORD cmdf;}OLECMD;
typedef struct _tagOLECMDTEXT{DWORD cmdtextf; ULONG cwActual,cwBuf;wchar_t rgwz[1];}OLECMDTEXT;
typedef struct {IOleInPlaceSite inplace; _IOleInPlaceFrameEx  frame; } _IOleInPlaceSiteEx;
typedef struct {IDocHostUIHandler  ui; } _IDocHostUIHandlerEx;
typedef struct {IOleClientSite client; _IOleInPlaceSiteEx  inplace; _IDocHostUIHandlerEx ui; } _IOleClientSiteEx;
typedef struct {IOleInPlaceFrame frame; HWND window;} _IOleInPlaceFrameEx;

interface IDocHostUIHandler{CONST_VTBL struct IDocHostUIHandlerVtbl* lpVtbl;};
interface IOleClientSite{CONST_VTBL struct IOleClientSiteVtbl* lpVtbl;};
interface IOleInPlaceFrame{CONST_VTBL struct IOleInPlaceFrameVtbl* lpVtbl;};
interface IOleCommandTarget{CONST_VTBL struct IOleCommandTargetVtbl* lpVtbl;};
interface IOleInPlaceSite{CONST_VTBL struct IOleInPlaceSiteVtbl* lpVtbl;};


typedef struct IOleInPlaceFrameVtbl{
    BEGIN_INTERFACE
        virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleInPlaceFrame* This,__RPC__in REFIID riid,_COM_Outptr_  void** ppvObject);
        virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleInPlaceFrame* This);
        virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleInPlaceFrame* This);
        virtuelExp(IOleWindow, GetWindow)HRESULT(STDMETHODCALLTYPE* GetWindow)(__RPC__in IOleInPlaceFrame* This,__RPC__deref_out_opt HWND* phwnd);
        virtuelExp(IOleWindow, ContextSensitiveHelp)HRESULT(STDMETHODCALLTYPE* ContextSensitiveHelp)(__RPC__in IOleInPlaceFrame* This,BOOL fEnterMode);
        virtuelExp(IOleInPlaceUIWindow, GetBorder)HRESULT(STDMETHODCALLTYPE* GetBorder)(__RPC__in IOleInPlaceFrame* This,__RPC__out LPRECT lprectBorder);
        virtuelExp(IOleInPlaceUIWindow, RequestBorderSpace)HRESULT(STDMETHODCALLTYPE* RequestBorderSpace)(__RPC__in IOleInPlaceFrame* This, __RPC__in_opt LPCBORDERWIDTHS pborderwidths);
        virtuelExp(IOleInPlaceUIWindow, SetBorderSpace)HRESULT(STDMETHODCALLTYPE* SetBorderSpace)(__RPC__in IOleInPlaceFrame* This,__RPC__in_opt LPCBORDERWIDTHS pborderwidths);
        virtuelExp(IOleInPlaceUIWindow, SetActiveObject)HRESULT(STDMETHODCALLTYPE* SetActiveObject)(__RPC__in IOleInPlaceFrame* This, __RPC__in_opt IOleInPlaceActiveObject* pActiveObject, __RPC__in_opt_string LPCOLESTR pszObjName);
        virtuelExp(IOleInPlaceFrame, InsertMenus)HRESULT(STDMETHODCALLTYPE* InsertMenus)(__RPC__in IOleInPlaceFrame* This,__RPC__in HMENU hmenuShared,__RPC__inout LPOLEMENUGROUPWIDTHS lpMenuWidths);
        virtuelExp(IOleInPlaceFrame, SetMenu)HRESULT(STDMETHODCALLTYPE* SetMenu)(__RPC__in IOleInPlaceFrame* This,__RPC__in HMENU hmenuShared,__RPC__in HOLEMENU holemenu,__RPC__in HWND hwndActiveObject);
        virtuelExp(IOleInPlaceFrame, RemoveMenus)HRESULT(STDMETHODCALLTYPE* RemoveMenus)(__RPC__in IOleInPlaceFrame* This,__RPC__in HMENU hmenuShared);
        virtuelExp(IOleInPlaceFrame, SetStatusText)HRESULT(STDMETHODCALLTYPE* SetStatusText)(__RPC__in IOleInPlaceFrame* This,__RPC__in_opt LPCOLESTR pszStatusText);
        virtuelExp(IOleInPlaceFrame, EnableModeless)HRESULT(STDMETHODCALLTYPE* EnableModeless)(__RPC__in IOleInPlaceFrame* This,BOOL fEnable);
        virtuelExp(IOleInPlaceFrame, TranslateAccelerator)HRESULT(STDMETHODCALLTYPE* TranslateAccelerator)(__RPC__in IOleInPlaceFrame* This,__RPC__in LPMSG lpmsg,WORD wID);
    END_INTERFACE
} IOleInPlaceFrameVtbl;
typedef struct IOleClientSiteVtbl{
    BEGIN_INTERFACE
        virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleClientSite* This,__RPC__in REFIID riid,_COM_Outptr_  void** ppvObject);
        virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleClientSite* This);
        virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleClientSite* This);
        virtuelExp(IOleClientSite, SaveObject)HRESULT(STDMETHODCALLTYPE* SaveObject)(__RPC__in IOleClientSite* This);
        virtuelExp(IOleClientSite, GetMoniker)HRESULT(STDMETHODCALLTYPE* GetMoniker)(__RPC__in IOleClientSite* This,DWORD dwAssign,DWORD dwWhichMoniker,__RPC__deref_out_opt IMoniker** ppmk);
        virtuelExp(IOleClientSite, GetContainer)HRESULT(STDMETHODCALLTYPE* GetContainer)(__RPC__in IOleClientSite* This,__RPC__deref_out_opt IOleContainer** ppContainer);
        virtuelExp(IOleClientSite, ShowObject)HRESULT(STDMETHODCALLTYPE* ShowObject)(__RPC__in IOleClientSite* This);
        virtuelExp(IOleClientSite, OnShowWindow)HRESULT(STDMETHODCALLTYPE* OnShowWindow)(__RPC__in IOleClientSite* This,BOOL fShow);
        virtuelExp(IOleClientSite, RequestNewObjectLayout)HRESULT(STDMETHODCALLTYPE* RequestNewObjectLayout)(__RPC__in IOleClientSite* This);
    END_INTERFACE
} IOleClientSiteVtbl;
typedef struct IDocHostUIHandlerVtbl{
    BEGIN_INTERFACE
        virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(IDocHostUIHandler* This,  REFIID riid, _COM_Outptr_  void** ppvObject);
        virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)( IDocHostUIHandler* This);
        virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)( IDocHostUIHandler* This);
        virtuelExp(IDocHostUIHandler, ShowContextMenu)HRESULT(STDMETHODCALLTYPE* ShowContextMenu)( IDocHostUIHandler* This, _In_  DWORD dwID,_In_  POINT* ppt, _In_  IUnknown* pcmdtReserved, _In_  IDispatch* pdispReserved);
        virtuelExp(IDocHostUIHandler, GetHostInfo)HRESULT(STDMETHODCALLTYPE* GetHostInfo)(IDocHostUIHandler* This,_Inout_  DOCHOSTUIINFO* pInfo);
        virtuelExp(IDocHostUIHandler, ShowUI)HRESULT(STDMETHODCALLTYPE* ShowUI)( IDocHostUIHandler* This, _In_  DWORD dwID, _In_  IOleInPlaceActiveObject* pActiveObject, _In_  IOleCommandTarget* pCommandTarget,_In_  IOleInPlaceFrame* pFrame, _In_  IOleInPlaceUIWindow* pDoc);
        virtuelExp(IDocHostUIHandler, HideUI)HRESULT(STDMETHODCALLTYPE* HideUI)( IDocHostUIHandler* This);  
        virtuelExp(IDocHostUIHandler, UpdateUI)HRESULT(STDMETHODCALLTYPE* UpdateUI)( IDocHostUIHandler* This);
        virtuelExp(IDocHostUIHandler, EnableModeless)HRESULT(STDMETHODCALLTYPE* EnableModeless)( IDocHostUIHandler* This,  BOOL fEnable);
        virtuelExp(IDocHostUIHandler, OnDocWindowActivate)HRESULT(STDMETHODCALLTYPE* OnDocWindowActivate)( IDocHostUIHandler* This,  BOOL fActivate);
        virtuelExp(IDocHostUIHandler, OnFrameWindowActivate)HRESULT(STDMETHODCALLTYPE* OnFrameWindowActivate)( IDocHostUIHandler* This,  BOOL fActivate);
        virtuelExp(IDocHostUIHandler, ResizeBorder)HRESULT(STDMETHODCALLTYPE* ResizeBorder)( IDocHostUIHandler* This, _In_  LPCRECT prcBorder, _In_  IOleInPlaceUIWindow* pUIWindow, _In_  BOOL fRameWindow);
        virtuelExp(IDocHostUIHandler, TranslateAccelerator)HRESULT(STDMETHODCALLTYPE* TranslateAccelerator)( IDocHostUIHandler* This,  LPMSG lpMsg,  const GUID* pguidCmdGroup,  DWORD nCmdID);
        virtuelExp(IDocHostUIHandler, GetOptionKeyPath)HRESULT(STDMETHODCALLTYPE* GetOptionKeyPath)( IDocHostUIHandler* This,  _Out_  LPOLESTR* pchKey, DWORD dw);
        virtuelExp(IDocHostUIHandler, GetDropTarget)HRESULT(STDMETHODCALLTYPE* GetDropTarget)( IDocHostUIHandler* This, _In_  IDropTarget* pDropTarget, _Outptr_  IDropTarget** ppDropTarget);
        virtuelExp(IDocHostUIHandler, GetExternal)HRESULT(STDMETHODCALLTYPE* GetExternal)(_Outptr_result_maybenull_  IDispatch** ppDispatch);
        virtuelExp(IDocHostUIHandler, TranslateUrl)HRESULT(STDMETHODCALLTYPE* TranslateUrl)(IDocHostUIHandler* This, DWORD dwTranslate, _In_  LPWSTR pchURLIn, _Outptr_  LPWSTR* ppchURLOut);
        virtuelExp(IDocHostUIHandler, FilterDataObject)HRESULT(STDMETHODCALLTYPE* FilterDataObject)( IDocHostUIHandler* This, _In_  IDataObject* pDO, _Outptr_result_maybenull_  IDataObject** ppDORet);
    END_INTERFACE
} IDocHostUIHandlerVtbl;
typedef struct IOleCommandTargetVtbl{
    BEGIN_INTERFACE
        virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleCommandTarget* This,__RPC__in REFIID riid,_COM_Outptr_  void** ppvObject);
        virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleCommandTarget* This);
        virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(__RPC__in IOleCommandTarget* This);
        virtuelExp(IOleCommandTarget, QueryStatus)HRESULT(STDMETHODCALLTYPE* QueryStatus)(__RPC__in IOleCommandTarget* This,__RPC__in_opt const GUID* pguidCmdGroup,ULONG cCmds,__RPC__inout_ecount_full(cCmds) OLECMD prgCmds[],__RPC__inout_opt OLECMDTEXT* pCmdText);
        virtuelExp(IOleCommandTarget, Exec)HRESULT(STDMETHODCALLTYPE* Exec)(__RPC__in IOleCommandTarget* This,__RPC__in_opt const GUID* pguidCmdGroup,DWORD nCmdID,DWORD nCmdexecopt,__RPC__in_opt VARIANT* pvaIn,__RPC__inout_opt VARIANT* pvaOut);
    END_INTERFACE
} IOleCommandTargetVtbl;
typedef struct IDocHostUIHandlerVtbl{
    BEGIN_INTERFACE
        virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(IDocHostUIHandler* This,REFIID riid, _COM_Outptr_  void** ppvObject);
        virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(IDocHostUIHandler* This);
        virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)(IDocHostUIHandler* This);
        virtuelExp(IDocHostUIHandler, ShowContextMenu)HRESULT(STDMETHODCALLTYPE* ShowContextMenu)(IDocHostUIHandler* This,_In_  DWORD dwID,_In_  POINT* ppt,_In_  IUnknown* pcmdtReserved,_In_  IDispatch* pdispReserved);
        virtuelExp(IDocHostUIHandler, GetHostInfo)HRESULT(STDMETHODCALLTYPE* GetHostInfo)(IDocHostUIHandler* This,_Inout_  DOCHOSTUIINFO* pInfo);
        virtuelExp(IDocHostUIHandler, ShowUI)HRESULT(STDMETHODCALLTYPE* ShowUI)(IDocHostUIHandler* This,_In_  DWORD dwID,_In_  IOleInPlaceActiveObject* pActiveObject,_In_  IOleCommandTarget* pCommandTarget,_In_  IOleInPlaceFrame* pFrame,_In_  IOleInPlaceUIWindow* pDoc);
        virtuelExp(IDocHostUIHandler, HideUI)HRESULT(STDMETHODCALLTYPE* HideUI)(IDocHostUIHandler* This);
        virtuelExp(IDocHostUIHandler, UpdateUI)HRESULT(STDMETHODCALLTYPE* UpdateUI)(IDocHostUIHandler* This);
        virtuelExp(IDocHostUIHandler, EnableModeless)HRESULT(STDMETHODCALLTYPE* EnableModeless)(IDocHostUIHandler* This,BOOL fEnable);
        virtuelExp(IDocHostUIHandler, OnDocWindowActivate)HRESULT(STDMETHODCALLTYPE* OnDocWindowActivate)(IDocHostUIHandler* This,BOOL fActivate);
        virtuelExp(IDocHostUIHandler, OnFrameWindowActivate)HRESULT(STDMETHODCALLTYPE* OnFrameWindowActivate)(IDocHostUIHandler* This,BOOL fActivate);
        virtuelExp(IDocHostUIHandler, ResizeBorder)HRESULT(STDMETHODCALLTYPE* ResizeBorder)(IDocHostUIHandler* This,_In_  LPCRECT prcBorder,_In_  IOleInPlaceUIWindow* pUIWindow,_In_  BOOL fRameWindow);
        virtuelExp(IDocHostUIHandler, TranslateAccelerator)HRESULT(STDMETHODCALLTYPE* TranslateAccelerator)(IDocHostUIHandler* This,LPMSG lpMsg,const GUID* pguidCmdGroup,DWORD nCmdID);
        virtuelExp(IDocHostUIHandler, GetOptionKeyPath)HRESULT(STDMETHODCALLTYPE* GetOptionKeyPath)(IDocHostUIHandler* This,_Out_  LPOLESTR* pchKey,WORD dw);
        virtuelExp(IDocHostUIHandler, GetDropTarget)HRESULT(STDMETHODCALLTYPE* GetDropTarget)(IDocHostUIHandler* This,_In_  IDropTarget* pDropTarget,_Outptr_  IDropTarget** ppDropTarget);
        virtuelExp(IDocHostUIHandler, GetExternal)HRESULT(STDMETHODCALLTYPE* GetExternal)(IDocHostUIHandler* This,_Outptr_result_maybenull_  IDispatch** ppDispatch);
        virtuelExp(IDocHostUIHandler, TranslateUrl)HRESULT(STDMETHODCALLTYPE* TranslateUrl)(IDocHostUIHandler* This,DWORD dwTranslate, _In_  LPWSTR pchURLIn,_Outptr_  LPWSTR* ppchURLOut);
        virtuelExp(IDocHostUIHandler, FilterDataObject)HRESULT(STDMETHODCALLTYPE* FilterDataObject)(IDocHostUIHandler* This,_In_  IDataObject* pDO,_Outptr_result_maybenull_  IDataObject** ppDORet);
    END_INTERFACE
} IDocHostUIHandlerVtbl;
typedef struct IOleInPlaceSiteVtbl{
    BEGIN_INTERFACE
        virtuelExp(IUnknown, QueryInterface)HRESULT(STDMETHODCALLTYPE* QueryInterface)(__RPC__in IOleInPlaceSite* This,__RPC__in REFIID riid,_COM_Outptr_  void** ppvObject);
        virtuelExp(IUnknown, AddRef)ULONG(STDMETHODCALLTYPE* AddRef)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IUnknown, Release)ULONG(STDMETHODCALLTYPE* Release)( __RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleWindow, GetWindow)HRESULT(STDMETHODCALLTYPE* GetWindow)(__RPC__in IOleInPlaceSite* This,__RPC__deref_out_opt HWND* phwnd);
        virtuelExp(IOleWindow, ContextSensitiveHelp)HRESULT(STDMETHODCALLTYPE* ContextSensitiveHelp)(__RPC__in IOleInPlaceSite* This,BOOL fEnterMode);
        virtuelExp(IOleInPlaceSite, CanInPlaceActivate)HRESULT(STDMETHODCALLTYPE* CanInPlaceActivate)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleInPlaceSite, OnInPlaceActivate)HRESULT(STDMETHODCALLTYPE* OnInPlaceActivate)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleInPlaceSite, OnUIActivate)HRESULT(STDMETHODCALLTYPE* OnUIActivate)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleInPlaceSite, GetWindowContext)HRESULT(STDMETHODCALLTYPE* GetWindowContext)(__RPC__in IOleInPlaceSite* This,__RPC__deref_out_opt IOleInPlaceFrame** ppFrame,__RPC__deref_out_opt IOleInPlaceUIWindow** ppDoc,__RPC__out LPRECT lprcPosRect,__RPC__out LPRECT lprcClipRect,__RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo);
        virtuelExp(IOleInPlaceSite, Scroll)HRESULT(STDMETHODCALLTYPE* Scroll)(__RPC__in IOleInPlaceSite* This,SIZE scrollExtant);
        virtuelExp(IOleInPlaceSite, OnUIDeactivate)HRESULT(STDMETHODCALLTYPE* OnUIDeactivate)(__RPC__in IOleInPlaceSite* This,BOOL fUndoable);
        virtuelExp(IOleInPlaceSite, OnInPlaceDeactivate)HRESULT(STDMETHODCALLTYPE* OnInPlaceDeactivate)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleInPlaceSite, DiscardUndoState)HRESULT(STDMETHODCALLTYPE* DiscardUndoState)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleInPlaceSite, DeactivateAndUndo)HRESULT(STDMETHODCALLTYPE* DeactivateAndUndo)(__RPC__in IOleInPlaceSite* This);
        virtuelExp(IOleInPlaceSite, OnPosRectChange)HRESULT(STDMETHODCALLTYPE* OnPosRectChange)(__RPC__in IOleInPlaceSite* This,__RPC__in LPCRECT lprcPosRect);
    END_INTERFACE
} IOleInPlaceSiteVtbl;



HRESULT STDMETHODCALLTYPE UI_QueryInterface(IDocHostUIHandler *This, REFIID riid, LPVOID *ppvObj){return(Site_QueryInterface((IOleClientSite *)((char *)This - sizeof(IOleClientSite) - sizeof(_IOleInPlaceSiteEx)), riid, ppvObj));}
HRESULT STDMETHODCALLTYPE UI_AddRef(IDocHostUIHandler *This){return(1);}
HRESULT STDMETHODCALLTYPE UI_Release(IDocHostUIHandler *This){return(1);}
HRESULT STDMETHODCALLTYPE UI_ShowContextMenu(IDocHostUIHandler *This, DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_GetHostInfo(IDocHostUIHandler * This, DOCHOSTUIINFO *pInfo){pInfo->cbSize = sizeof(DOCHOSTUIINFO); pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER; pInfo->dwDoubleClick = 0; return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_ShowUI(IDocHostUIHandler *This, DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_HideUI(IDocHostUIHandler *This){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_UpdateUI(IDocHostUIHandler *This){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_EnableModeless(IDocHostUIHandler *This, BOOL fEnable){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_OnDocWindowActivate(IDocHostUIHandler *This, BOOL fActivate){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_OnFrameWindowActivate(IDocHostUIHandler *This, BOOL fActivate){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_ResizeBorder(IDocHostUIHandler *This, LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow){return(S_OK);}
HRESULT STDMETHODCALLTYPE UI_TranslateAccelerator(IDocHostUIHandler *This, LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID){return(S_FALSE);}
HRESULT STDMETHODCALLTYPE UI_GetOptionKeyPath(IDocHostUIHandler *This, LPOLESTR *pchKey, DWORD dw){return(S_FALSE);}
HRESULT STDMETHODCALLTYPE UI_GetDropTarget(IDocHostUIHandler *This, IDropTarget *pDropTarget, IDropTarget **ppDropTarget){return(S_FALSE);}
HRESULT STDMETHODCALLTYPE UI_GetExternal(IDocHostUIHandler *This, IDispatch **ppDispatch){*ppDispatch = 0; return(S_FALSE);}
HRESULT STDMETHODCALLTYPE UI_TranslateUrl(IDocHostUIHandler *This, DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut){
	unsigned short *src;
	unsigned short *dest;
	DWORD len;
	src = pchURLIn;
	while ((*(src)++));
	--src;
	len = src - pchURLIn; 
	if (len >= 4 && !_wcsnicmp(pchURLIn, (wchar_t *)&AppUrl[0], 4)){
		if ((dest = (OLECHAR *)CoTaskMemAlloc(12<<1))){
		HWND hwnd;
		*ppchURLOut = dest;
		CopyMemory(dest, &Blank[0], 12<<1);
		len = Asc2Int(pchURLIn + 4);
		hwnd = ((_IOleInPlaceSiteEx *)((char *)This - sizeof(_IOleInPlaceSiteEx)))->frame.window;
		PostMessage(hwnd, WM_APP + 50, (WPARAM)len, 0);
		return(S_OK);
		}
	}
	*ppchURLOut = 0;
	return(S_FALSE);
}
HRESULT STDMETHODCALLTYPE UI_FilterDataObject(IDocHostUIHandler *This, IDataObject *pDO, IDataObject **ppDORet){*ppDORet = 0;return(S_FALSE);}
HRESULT STDMETHODCALLTYPE Site_QueryInterface(IOleClientSite *This, REFIID riid, void ** ppvObject){
	if (!memcmp(riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(riid, &IID_IOleClientSite, sizeof(GUID)))  *ppvObject = &((_IOleClientSiteEx *)This)->client;
	else if (!memcmp(riid, &IID_IOleInPlaceSite, sizeof(GUID)))  *ppvObject = &((_IOleClientSiteEx *)This)->inplace;
	else if (!memcmp(riid, &IID_IDocHostUIHandler, sizeof(GUID))) *ppvObject = &((_IOleClientSiteEx *)This)->ui;
	else { *ppvObject = 0;  return(E_NOINTERFACE); }
	return(S_OK);
}
HRESULT STDMETHODCALLTYPE Site_AddRef(IOleClientSite *This){return(1);}
HRESULT STDMETHODCALLTYPE Site_Release(IOleClientSite *This){return(1);}
HRESULT STDMETHODCALLTYPE Site_SaveObject(IOleClientSite *This){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Site_GetMoniker(IOleClientSite *This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker ** ppmk){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Site_GetContainer(IOleClientSite *This, LPOLECONTAINER *ppContainer){*ppContainer = 0; return(E_NOINTERFACE);}
HRESULT STDMETHODCALLTYPE Site_ShowObject(IOleClientSite *This){return(NOERROR);}
HRESULT STDMETHODCALLTYPE Site_OnShowWindow(IOleClientSite *This, BOOL fShow){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Site_RequestNewObjectLayout(IOleClientSite *This){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE InPlace_QueryInterface(IOleInPlaceSite *This, REFIID riid, LPVOID *ppvObj){return(Site_QueryInterface((IOleClientSite *)((char *)This - sizeof(IOleClientSite)), riid, ppvObj));}
HRESULT STDMETHODCALLTYPE InPlace_AddRef(IOleInPlaceSite *This){return(1);}
HRESULT STDMETHODCALLTYPE InPlace_Release(IOleInPlaceSite *This){return(1);}
HRESULT STDMETHODCALLTYPE InPlace_GetWindow(IOleInPlaceSite *This, HWND *lphwnd){*lphwnd = ((_IOleInPlaceSiteEx *)This)->frame.window; return(S_OK);}
HRESULT STDMETHODCALLTYPE InPlace_ContextSensitiveHelp(IOleInPlaceSite *This, BOOL fEnterMode){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE InPlace_CanInPlaceActivate(IOleInPlaceSite *This){return(S_OK);}
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceActivate(IOleInPlaceSite *This){return(S_OK);}
HRESULT STDMETHODCALLTYPE InPlace_OnUIActivate(IOleInPlaceSite *This){return(S_OK);}
HRESULT STDMETHODCALLTYPE InPlace_GetWindowContext(IOleInPlaceSite *This, LPOLEINPLACEFRAME *lplpFrame, LPOLEINPLACEUIWINDOW *lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo){
	*lplpFrame = (LPOLEINPLACEFRAME)&((_IOleInPlaceSiteEx *)This)->frame;
	*lplpDoc = 0;
	lpFrameInfo->fMDIApp = FALSE;
	lpFrameInfo->hwndFrame = ((_IOleInPlaceFrameEx *)*lplpFrame)->window;
	lpFrameInfo->haccel = 0;
	lpFrameInfo->cAccelEntries = 0;
	return(S_OK);
}
HRESULT STDMETHODCALLTYPE InPlace_Scroll(IOleInPlaceSite *This, SIZE scrollExtent){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE InPlace_OnUIDeactivate(IOleInPlaceSite *This, BOOL fUndoable){return(S_OK);}
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceDeactivate(IOleInPlaceSite *This){return(S_OK);}
HRESULT STDMETHODCALLTYPE InPlace_DiscardUndoState(IOleInPlaceSite *This){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE InPlace_DeactivateAndUndo(IOleInPlaceSite *This){return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_QueryInterface(IOleInPlaceFrame *This, REFIID riid, LPVOID *ppvObj){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_AddRef(IOleInPlaceFrame *This){ return(1);}
HRESULT STDMETHODCALLTYPE Frame_Release(IOleInPlaceFrame *This){ return(1);}
HRESULT STDMETHODCALLTYPE Frame_GetWindow(IOleInPlaceFrame *This, HWND *lphwnd){*lphwnd = ((_IOleInPlaceFrameEx *)This)->window; return(S_OK);}
HRESULT STDMETHODCALLTYPE Frame_ContextSensitiveHelp(IOleInPlaceFrame *This, BOOL fEnterMode){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_GetBorder(IOleInPlaceFrame *This, LPRECT lprectBorder){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_RequestBorderSpace(IOleInPlaceFrame *This, LPCBORDERWIDTHS pborderwidths){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_SetBorderSpace(IOleInPlaceFrame *This, LPCBORDERWIDTHS pborderwidths){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_SetActiveObject(IOleInPlaceFrame *This, IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName){ return(S_OK);}
HRESULT STDMETHODCALLTYPE Frame_InsertMenus(IOleInPlaceFrame *This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_SetMenu(IOleInPlaceFrame* This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_RemoveMenus(IOleInPlaceFrame *This, HMENU hmenuShared){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE Frame_SetStatusText(IOleInPlaceFrame* This, LPCOLESTR pszStatusText) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_EnableModeless(IOleInPlaceFrame* This, BOOL fEnable) { return(S_OK); }
HRESULT STDMETHODCALLTYPE Frame_TranslateAccelerator(IOleInPlaceFrame *This, LPMSG lpmsg, WORD wID){ return(E_NOTIMPL);}
HRESULT STDMETHODCALLTYPE InPlace_OnPosRectChange(IOleInPlaceSite *This, LPCRECT lprcPosRect){
	IOleObject *browserObject;
	IOleInPlaceObject *inplace=0;
	browserObject = *((IOleObject **)((char *)This - sizeof(IOleObject *) - sizeof(IOleClientSite)));
	if (!browserObject->QueryInterface(browserObject, &IID_IOleInPlaceObject, (void**)&inplace)) { inplace->SetObjectRects(inplace, lprcPosRect, lprcPosRect); }
	return(S_OK);
}

