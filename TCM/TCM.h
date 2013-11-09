#pragma once

#include "resource.h"
#include <ShObjIdl.h>
#include <ShellAPI.h>
#include <Shlobj.h>
#include <Shlwapi.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "Urlmon.lib") 
#pragma comment(linker,"/manifestdependency:\"type='win32' \
  name='Microsoft.Windows.Common-Controls' \
  version='6.0.0.0' \
  processorArchitecture='*' \
  publicKeyToken='6595b64144ccf1df' \
  language='*'\"") 

#define TCMT_CLOSE	1
#define TCMT_SHOW	2

class CtcmContextMenu : public IContextMenu
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IContextMenu
	STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
	STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);
	STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);

	CtcmContextMenu(IContextMenu *pContextMenu, HWND hwnd);
	~CtcmContextMenu();
public:
	IContextMenu *m_pContextMenu;
	IContextMenu2 *m_pContextMenu2;
	IContextMenu3 *m_pContextMenu3;
private:
	LONG	m_cRef;
	HWND	m_hwnd;
};

class CtcmDataObject : public IDataObject
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDataObject
	STDMETHODIMP GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium);
	STDMETHODIMP GetDataHere(FORMATETC *pformatetc, STGMEDIUM *pmedium);
	STDMETHODIMP QueryGetData(FORMATETC *pformatetc);
	STDMETHODIMP GetCanonicalFormatEtc(FORMATETC *pformatectIn, FORMATETC *pformatetcOut);
	STDMETHODIMP SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease);
	STDMETHODIMP EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc);
	STDMETHODIMP DAdvise(FORMATETC *pformatetc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection);
	STDMETHODIMP DUnadvise(DWORD dwConnection);
	STDMETHODIMP EnumDAdvise(IEnumSTATDATA **ppenumAdvise);

	CtcmDataObject(LPITEMIDLIST *apidl, int nArgs, HWND hwnd);
	~CtcmDataObject();

	HDROP GethDrop(int x, int y, BOOL fNC, BOOL bSpecial);
private:
	LPITEMIDLIST	*m_pidllist;
	LONG			m_cRef;
	long			m_nCount;
	HWND	m_hwnd;
};