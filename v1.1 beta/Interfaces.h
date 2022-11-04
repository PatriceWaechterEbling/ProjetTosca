#pragma once
#include <windows.h>

#define NOTIMPLEMENTED _ASSERT(0); return(E_NOTIMPL)
static unsigned char AppUrl[] = {'a', 0, 'p', 0, 'p', 0, ':', 0};
#define GOBACK  0
#define GOFORWARD 1
#define GOHOME  2
#define SEARCH  3
#define REFRESH  4
#define STOP  5

#define virtuelExp(base, func) DECLSPEC_XFGVIRT(base, func)
#define IOleInPlaceFrame_QueryInterface(This,riid,ppvObject)    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 
#define IOleInPlaceFrame_AddRef(This)( (This)->lpVtbl -> AddRef(This) ) 
#define IOleInPlaceFrame_Release(This)( (This)->lpVtbl -> Release(This) ) 
#define IOleInPlaceFrame_GetWindow(This,phwnd)( (This)->lpVtbl -> GetWindow(This,phwnd) ) 
#define IOleInPlaceFrame_ContextSensitiveHelp(This,fEnterMode)( (This)->lpVtbl -> ContextSensitiveHelp(This,fEnterMode) ) 
#define IOleInPlaceFrame_GetBorder(This,lprectBorder)( (This)->lpVtbl -> GetBorder(This,lprectBorder) ) 
#define IOleInPlaceFrame_RequestBorderSpace(This,pborderwidths)( (This)->lpVtbl -> RequestBorderSpace(This,pborderwidths) ) 
#define IOleInPlaceFrame_SetBorderSpace(This,pborderwidths)( (This)->lpVtbl -> SetBorderSpace(This,pborderwidths) ) 
#define IOleInPlaceFrame_SetActiveObject(This,pActiveObject,pszObjName)	   ( (This)->lpVtbl -> SetActiveObject(This,pActiveObject,pszObjName) ) 
#define IOleInPlaceFrame_InsertMenus(This,hmenuShared,lpMenuWidths)( (This)->lpVtbl -> InsertMenus(This,hmenuShared,lpMenuWidths) ) 
#define IOleInPlaceFrame_SetMenu(This,hmenuShared,holemenu,hwndActiveObject)( (This)->lpVtbl -> SetMenu(This,hmenuShared,holemenu,hwndActiveObject) ) 
#define IOleInPlaceFrame_RemoveMenus(This,hmenuShared)( (This)->lpVtbl -> RemoveMenus(This,hmenuShared) ) 
#define IOleInPlaceFrame_SetStatusText(This,pszStatusText)( (This)->lpVtbl -> SetStatusText(This,pszStatusText) ) 
#define IOleInPlaceFrame_EnableModeless(This,fEnable)( (This)->lpVtbl -> EnableModeless(This,fEnable) ) 
#define IOleInPlaceFrame_TranslateAccelerator(This,lpmsg,wID)( (This)->lpVtbl -> TranslateAccelerator(This,lpmsg,wID) ) 
#define IOleClientSite_QueryInterface(This,riid,ppvObject)( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 
#define IOleClientSite_AddRef(This)( (This)->lpVtbl -> AddRef(This) ) 
#define IOleClientSite_Release(This)( (This)->lpVtbl -> Release(This) ) 
#define IOleClientSite_SaveObject(This)( (This)->lpVtbl -> SaveObject(This) ) 
#define IOleClientSite_GetMoniker(This,dwAssign,dwWhichMoniker,ppmk)( (This)->lpVtbl -> GetMoniker(This,dwAssign,dwWhichMoniker,ppmk) ) 
#define IOleClientSite_GetContainer(This,ppContainer)( (This)->lpVtbl -> GetContainer(This,ppContainer) ) 
#define IOleClientSite_ShowObject(This)( (This)->lpVtbl -> ShowObject(This) ) 
#define IOleClientSite_OnShowWindow(This,fShow)( (This)->lpVtbl -> OnShowWindow(This,fShow) ) 
#define IOleClientSite_RequestNewObjectLayout(This)( (This)->lpVtbl -> RequestNewObjectLayout(This) ) 
#define IOleCommandTarget_QueryInterface(This,riid,ppvObject)( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 
#define IOleCommandTarget_AddRef(This)( (This)->lpVtbl -> AddRef(This) ) 
#define IOleCommandTarget_Release(This)( (This)->lpVtbl -> Release(This) ) 
#define IOleCommandTarget_QueryStatus(This,pguidCmdGroup,cCmds,prgCmds,pCmdText)( (This)->lpVtbl -> QueryStatus(This,pguidCmdGroup,cCmds,prgCmds,pCmdText) ) 
#define IOleCommandTarget_Exec(This,pguidCmdGroup,nCmdID,nCmdexecopt,pvaIn,pvaOut)( (This)->lpVtbl -> Exec(This,pguidCmdGroup,nCmdID,nCmdexecopt,pvaIn,pvaOut) ) 
#define IOleInPlaceSite_QueryInterface(This,riid,ppvObject)	( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 
#define IOleInPlaceSite_AddRef(This)	( (This)->lpVtbl -> AddRef(This) ) 
#define IOleInPlaceSite_Release(This)	( (This)->lpVtbl -> Release(This) ) 
#define IOleInPlaceSite_GetWindow(This,phwnd)	( (This)->lpVtbl -> GetWindow(This,phwnd) ) 
#define IOleInPlaceSite_ContextSensitiveHelp(This,fEnterMode)	( (This)->lpVtbl -> ContextSensitiveHelp(This,fEnterMode) ) 
#define IOleInPlaceSite_CanInPlaceActivate(This)	( (This)->lpVtbl -> CanInPlaceActivate(This) ) 
#define IOleInPlaceSite_OnInPlaceActivate(This)	( (This)->lpVtbl -> OnInPlaceActivate(This) ) 
#define IOleInPlaceSite_OnUIActivate(This)	( (This)->lpVtbl -> OnUIActivate(This) ) 
#define IOleInPlaceSite_GetWindowContext(This,ppFrame,ppDoc,lprcPosRect,lprcClipRect,lpFrameInfo)	( (This)->lpVtbl -> GetWindowContext(This,ppFrame,ppDoc,lprcPosRect,lprcClipRect,lpFrameInfo) ) 
#define IOleInPlaceSite_Scroll(This,scrollExtant)	( (This)->lpVtbl -> Scroll(This,scrollExtant) ) 
#define IOleInPlaceSite_OnUIDeactivate(This,fUndoable)	( (This)->lpVtbl -> OnUIDeactivate(This,fUndoable) ) 
#define IOleInPlaceSite_OnInPlaceDeactivate(This)	( (This)->lpVtbl -> OnInPlaceDeactivate(This) ) 
#define IOleInPlaceSite_DiscardUndoState(This)	( (This)->lpVtbl -> DiscardUndoState(This) ) 
#define IOleInPlaceSite_DeactivateAndUndo(This)	( (This)->lpVtbl -> DeactivateAndUndo(This) ) 
#define IOleInPlaceSite_OnPosRectChange(This,lprcPosRect)	( (This)->lpVtbl -> OnPosRectChange(This,lprcPosRect) ) 

HRESULT STDMETHODCALLTYPE Frame_QueryInterface(IOleInPlaceFrame *This, REFIID riid, LPVOID *ppvObj);
HRESULT STDMETHODCALLTYPE Frame_AddRef(IOleInPlaceFrame *This);
HRESULT STDMETHODCALLTYPE Frame_Release(IOleInPlaceFrame *This);
HRESULT STDMETHODCALLTYPE Frame_GetWindow(IOleInPlaceFrame *This, HWND *lphwnd);
HRESULT STDMETHODCALLTYPE Frame_ContextSensitiveHelp(IOleInPlaceFrame *This, BOOL fEnterMode);
HRESULT STDMETHODCALLTYPE Frame_GetBorder(IOleInPlaceFrame *This, LPRECT lprectBorder);
HRESULT STDMETHODCALLTYPE Frame_RequestBorderSpace(IOleInPlaceFrame *This, LPCBORDERWIDTHS pborderwidths);
HRESULT STDMETHODCALLTYPE Frame_SetBorderSpace(IOleInPlaceFrame *This, LPCBORDERWIDTHS pborderwidths);
HRESULT STDMETHODCALLTYPE Frame_SetActiveObject(IOleInPlaceFrame *This, IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName);
HRESULT STDMETHODCALLTYPE Frame_InsertMenus(IOleInPlaceFrame *This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
HRESULT STDMETHODCALLTYPE Frame_SetMenu(IOleInPlaceFrame *This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject);
HRESULT STDMETHODCALLTYPE Frame_RemoveMenus(IOleInPlaceFrame *This, HMENU hmenuShared);
HRESULT STDMETHODCALLTYPE Frame_SetStatusText(IOleInPlaceFrame *This, LPCOLESTR pszStatusText);
HRESULT STDMETHODCALLTYPE Frame_EnableModeless(IOleInPlaceFrame *This, BOOL fEnable);
HRESULT STDMETHODCALLTYPE Frame_TranslateAccelerator(IOleInPlaceFrame *This, LPMSG lpmsg, WORD wID);
HRESULT STDMETHODCALLTYPE Site_QueryInterface(IOleClientSite *This, REFIID riid, void ** ppvObject);
HRESULT STDMETHODCALLTYPE Site_AddRef(IOleClientSite *This);
HRESULT STDMETHODCALLTYPE Site_Release(IOleClientSite *This);
HRESULT STDMETHODCALLTYPE Site_SaveObject(IOleClientSite *This);
HRESULT STDMETHODCALLTYPE Site_GetMoniker(IOleClientSite *This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker ** ppmk);
HRESULT STDMETHODCALLTYPE Site_GetContainer(IOleClientSite *This, LPOLECONTAINER *ppContainer);
HRESULT STDMETHODCALLTYPE Site_ShowObject(IOleClientSite *This);
HRESULT STDMETHODCALLTYPE Site_OnShowWindow(IOleClientSite *This, BOOL fShow);
HRESULT STDMETHODCALLTYPE Site_RequestNewObjectLayout(IOleClientSite *This);
HRESULT STDMETHODCALLTYPE UI_QueryInterface(IDocHostUIHandler *This, REFIID riid, void **ppvObject);
HRESULT STDMETHODCALLTYPE UI_AddRef(IDocHostUIHandler *This);
HRESULT STDMETHODCALLTYPE UI_Release(IDocHostUIHandler *This);
HRESULT STDMETHODCALLTYPE UI_ShowContextMenu(IDocHostUIHandler *This, DWORD dwID, POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved);
HRESULT STDMETHODCALLTYPE UI_GetHostInfo(IDocHostUIHandler *This, DOCHOSTUIINFO *pInfo);
HRESULT STDMETHODCALLTYPE UI_ShowUI(IDocHostUIHandler *This, DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc);
HRESULT STDMETHODCALLTYPE UI_HideUI(IDocHostUIHandler *This);
HRESULT STDMETHODCALLTYPE UI_UpdateUI(IDocHostUIHandler *This);
HRESULT STDMETHODCALLTYPE UI_EnableModeless(IDocHostUIHandler *This, BOOL fEnable);
HRESULT STDMETHODCALLTYPE UI_OnDocWindowActivate(IDocHostUIHandler *This, BOOL fActivate);
HRESULT STDMETHODCALLTYPE UI_OnFrameWindowActivate(IDocHostUIHandler *This, BOOL fActivate);
HRESULT STDMETHODCALLTYPE UI_ResizeBorder(IDocHostUIHandler *This, LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow);
HRESULT STDMETHODCALLTYPE UI_TranslateAccelerator(IDocHostUIHandler *This, LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID);
HRESULT STDMETHODCALLTYPE UI_GetOptionKeyPath(IDocHostUIHandler *This, LPOLESTR *pchKey, DWORD dw);
HRESULT STDMETHODCALLTYPE UI_GetDropTarget(IDocHostUIHandler *This, IDropTarget *pDropTarget, IDropTarget **ppDropTarget);
HRESULT STDMETHODCALLTYPE UI_GetExternal(IDocHostUIHandler *This, IDispatch **ppDispatch);
HRESULT STDMETHODCALLTYPE UI_TranslateUrl(IDocHostUIHandler *This, DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut);
HRESULT STDMETHODCALLTYPE UI_FilterDataObject(IDocHostUIHandler *This, IDataObject *pDO, IDataObject **ppDORet);
HRESULT STDMETHODCALLTYPE InPlace_QueryInterface(IOleInPlaceSite *This, REFIID riid, void **ppvObject);
HRESULT STDMETHODCALLTYPE InPlace_AddRef(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_Release(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_GetWindow(IOleInPlaceSite *This, HWND *lphwnd);
HRESULT STDMETHODCALLTYPE InPlace_ContextSensitiveHelp(IOleInPlaceSite * This, BOOL fEnterMode);
HRESULT STDMETHODCALLTYPE InPlace_CanInPlaceActivate(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceActivate(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_OnUIActivate(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_GetWindowContext(IOleInPlaceSite *This, LPOLEINPLACEFRAME *lplpFrame, LPOLEINPLACEUIWINDOW *lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
HRESULT STDMETHODCALLTYPE InPlace_Scroll(IOleInPlaceSite *This, SIZE scrollExtent);
HRESULT STDMETHODCALLTYPE InPlace_OnUIDeactivate(IOleInPlaceSite *This, BOOL fUndoable);
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceDeactivate(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_DiscardUndoState(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_DeactivateAndUndo(IOleInPlaceSite *This);
HRESULT STDMETHODCALLTYPE InPlace_OnPosRectChange(IOleInPlaceSite *This, LPCRECT lprcPosRect);
IDocHostUIHandlerVtbl MyIDocHostUIHandlerTable={UI_QueryInterface,UI_AddRef,UI_Release,UI_ShowContextMenu,UI_GetHostInfo,UI_ShowUI,UI_HideUI,UI_UpdateUI,UI_EnableModeless,UI_OnDocWindowActivate,UI_OnFrameWindowActivate,UI_ResizeBorder,UI_TranslateAccelerator,UI_GetOptionKeyPath,UI_GetDropTarget,UI_GetExternal,UI_TranslateUrl,UI_FilterDataObject};
IOleInPlaceSiteVtbl MyIOleInPlaceSiteTable =  {InPlace_QueryInterface,InPlace_AddRef,InPlace_Release,InPlace_GetWindow,InPlace_ContextSensitiveHelp,InPlace_CanInPlaceActivate,InPlace_OnInPlaceActivate,InPlace_OnUIActivate,InPlace_GetWindowContext,InPlace_Scroll,InPlace_OnUIDeactivate,InPlace_OnInPlaceDeactivate,InPlace_DiscardUndoState,InPlace_DeactivateAndUndo,InPlace_OnPosRectChange};
