// TCM.cpp
//
// Tablacus ContextMenu
// Copyright (c) 2011 Gaku
// Licensed under the MIT License:
// http://www.opensource.org/licenses/mit-license.php
//
// Visual C++ 2008 Express Edition
// Windows SDK v7.0 
//
// http://www.eonet.ne.jp/~gakana/tablacus/

#include "stdafx.h"
#include "TCM.h"

#define MAX_LOADSTRING 100
#define TCM_ABOUT	L"Tablacus ContextMenu 13.11.9 Gaku"
#define TCM_URL		L"http://www.eonet.ne.jp/~gakana/tablacus/"
// Global variables:
HINSTANCE hInst;								// current instance
TCHAR szWindowClass[MAX_LOADSTRING];
TCHAR szMessage[MAX_LOADSTRING];
FORMATETC HDROPFormat  = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};

HWND g_hWnd;
DWORD g_dwProcId;
int g_nCmdShow;
CtcmContextMenu *g_pCCM = NULL;
static UINT_PTR g_uTimerId = 0;
static UINT_PTR g_uTimerId2 = 0;
RECT g_rc;
int g_nCount = 0;
BSTR g_bsInvoke = NULL;
IDropTarget *g_pDropTarget = NULL;
DWORD g_grfKeyState = MK_RBUTTON;

int TCMFindVerb(HMENU hMenu)
{
	MENUITEMINFO mii;
	::ZeroMemory(&mii, sizeof(MENUITEMINFO));
	mii.cbSize = sizeof(MENUITEMINFO);

	int nCount = GetMenuItemCount(hMenu);
	for (int i = 0; i < nCount; i++) {
		mii.dwTypeData = NULL;
		mii.fMask  = MIIM_STRING | MIIM_SUBMENU;
		GetMenuItemInfo(hMenu, i, TRUE, &mii);
		if (mii.hSubMenu) {
			TCMFindVerb(mii.hSubMenu);
			DestroyMenu(mii.hSubMenu);
		}
		else if (mii.cch) {
			LPWSTR dwTypeData = new WCHAR[++mii.cch + 1];
			mii.dwTypeData = dwTypeData;
			mii.fMask  = MIIM_STRING | MIIM_ID;
			GetMenuItemInfo(hMenu, i, TRUE, &mii);
			dwTypeData[mii.cch] = NULL;
			BOOL bMatch = PathMatchSpec(dwTypeData, g_bsInvoke);
			delete [] dwTypeData;
			if (bMatch) {
				return mii.wID;
			}
		}
	}
	return 0;
}


// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
HWND				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);

BOOL PopupContextMenu(HWND hwnd, LPTSTR *lplpszArgs, int nArgs, int nStart)
{
	IShellFolder *pSF = NULL;
	LPCITEMIDLIST *apidl = NULL;
	LPITEMIDLIST *apidlFull = NULL;

	IShellFolder *pSF2 = NULL;
	LPCITEMIDLIST pidl = NULL;
	LPITEMIDLIST pidlFull = NULL;
	LPITEMIDLIST pidlParent = NULL;

	BOOL Result = false;
	POINT pt;
	int i, npidl;

	apidl = new LPCITEMIDLIST[nArgs];
	apidlFull = new LPITEMIDLIST[nArgs];
	npidl = 0;

	CMINVOKECOMMANDINFO cmi;
	ZeroMemory(&cmi, sizeof(CMINVOKECOMMANDINFO));

	GetCursorPos(&pt);
	cmi.nShow = SW_SHOWNORMAL;
	UINT uFlags = CMF_NORMAL;

	i = nStart;
	while (i < nArgs) {
		if (lstrcmpi(lplpszArgs[i], L"-x") == 0) {
			pt.x = StrToInt(lplpszArgs[++i]);
		}
		else if (StrCmpN(lplpszArgs[i], L"/x=", 3) == 0) {
			pt.x = StrToInt(lplpszArgs[i] + 3);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-y") == 0) {
			pt.y = StrToInt(lplpszArgs[++i]);
		}
		else if (StrCmpN(lplpszArgs[i], L"/y=", 3) == 0) {
			pt.y = StrToInt(lplpszArgs[i] + 3);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-fMask") == 0) {
			cmi.fMask = StrToInt(lplpszArgs[++i]);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Show") == 0) {
			cmi.fMask = StrToInt(lplpszArgs[++i]);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Directory") == 0) {
			cmi.lpDirectory = (LPCSTR)lplpszArgs[++i];
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Parameter") == 0) {
			cmi.lpParameters = (LPCSTR)lplpszArgs[++i];
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Flags") == 0) {
			uFlags = StrToInt(lplpszArgs[++i]);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Timer") == 0) {
			g_nCount = StrToInt(lplpszArgs[++i]);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-grfKeyState") == 0) {
			g_grfKeyState = StrToInt(lplpszArgs[++i]);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Invoke") == 0) {
			g_bsInvoke = ::SysAllocString(lplpszArgs[++i]);
		}
		else if (lstrcmpi(lplpszArgs[i], L"-Drop") == 0) {
			pidlFull = ::ILCreateFromPath(lplpszArgs[++i]);
			if (pidlFull) {
				if SUCCEEDED(SHBindToParent(pidlFull, IID_PPV_ARGS(&pSF2), &pidl)) {
					pSF2->GetUIObjectOf(hwnd, 1, &pidl, IID_IDropTarget, NULL, (LPVOID*)&g_pDropTarget);
					pSF2->Release();
				}
				::CoTaskMemFree(pidlFull);
			}
		}
		else {
			int n = lstrlen(lplpszArgs[i]);
			if (lplpszArgs[i][n - 1] == '"') {
				lplpszArgs[i][n - 1] = '\\';
			}
			apidlFull[npidl] = ILCreateFromPath(lplpszArgs[i]);
			if (apidlFull[npidl]) {
				if (!pSF) {
					SHBindToParent(apidlFull[npidl], IID_PPV_ARGS(&pSF), &apidl[npidl]);
					pidlParent = ILClone(apidlFull[npidl]);
					ILRemoveLastID(pidlParent);
				}
				else {
					apidl[npidl] = ILFindLastID(apidlFull[npidl]);
					if (pidlParent) {
						if (!ILIsParent(pidlParent, apidlFull[npidl], TRUE)) {
							::CoTaskMemFree(pidlParent);
							pidlParent = NULL;
						}
					}
				}
				npidl++;
			}
		}
		i++;
	}
	if (pSF) {
		if (g_pDropTarget) {
			IDataObject *pDataObj = NULL;
			if (!pidlParent || FAILED(pSF->GetUIObjectOf(hwnd, npidl, &apidl[0], IID_IDataObject, NULL, (LPVOID*)&pDataObj))) {
				pDataObj= new CtcmDataObject(apidlFull, npidl, hwnd);
			}
			Result = true;
			DWORD dwEffect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
			POINTL ptl;
			ptl.x = pt.x;
			ptl.y = pt.y;
			if SUCCEEDED(g_pDropTarget->DragEnter(pDataObj, g_grfKeyState, ptl, &dwEffect)) {
				dwEffect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
				if SUCCEEDED(g_pDropTarget->DragOver(g_grfKeyState, ptl, &dwEffect)) {
					dwEffect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
					g_pDropTarget->Drop(pDataObj, g_grfKeyState, ptl, &dwEffect);
				}
			}
			pDataObj->Release();
			g_pDropTarget->Release();
			g_uTimerId = SetTimer(hwnd, TCMT_CLOSE, 1000, NULL);
		}
		else {
			IContextMenu *pCM;
			if SUCCEEDED(pSF->GetUIObjectOf(hwnd, npidl, &apidl[0], IID_IContextMenu, NULL, (LPVOID*)&pCM)) {
				g_pCCM = new CtcmContextMenu(pCM, hwnd);
				pCM->Release();
				
				HMENU hMenu;
				hMenu = CreatePopupMenu();
				char szVerbA[MAX_PATH];
				if SUCCEEDED(g_pCCM->QueryContextMenu(hMenu, 0, 1, 0x7fff, uFlags)) {
					Result = true;
					SetForegroundWindow(hwnd);
					int nVerb = 0;
					LPCSTR lpVerb = NULL;
					if (g_bsInvoke) {
						if (lstrcmpi(g_bsInvoke, L"Default") == 0) {
							nVerb = GetMenuDefaultItem(hMenu, 0, 0);
						}
						else {
							nVerb = TCMFindVerb(hMenu);
							if (nVerb == 0) {
								ZeroMemory(szVerbA, sizeof(szVerbA));
								if (WideCharToMultiByte(CP_ACP, NULL, g_bsInvoke, -1, szVerbA, sizeof(szVerbA), NULL, NULL)) {
									lpVerb = (LPCSTR)szVerbA;
								}
							}
						}
					} else {
						nVerb = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD,
							pt.x, pt.y, hwnd, NULL);
					}
					if (nVerb) {
						lpVerb = (LPCSTR)MAKEINTRESOURCE(nVerb - 1);
					}
					if (lpVerb) {
						cmi.cbSize = sizeof(CMINVOKECOMMANDINFO);
						cmi.hwnd = hwnd;
						cmi.lpVerb = lpVerb;
						g_pCCM->InvokeCommand(&cmi);
					}
					if (g_bsInvoke) {
						::SysFreeString(g_bsInvoke);
						g_bsInvoke = NULL;
					}
				}
				DestroyMenu(hMenu);
				g_pCCM->Release();
			}
		}
		pSF->Release();
	}
	for (i = 0; i < npidl; i++) {
		CoTaskMemFree(apidlFull[i]);
	}
	if (pidlParent) {
		::CoTaskMemFree(pidlParent);
		pidlParent = NULL;
	}
	delete [] apidl;
	delete [] apidlFull;
	return Result;
}

BOOL CALLBACK EnumWindowsProc(HWND hwnd , LPARAM lParam)
{
	if (hwnd != g_hWnd && IsWindowVisible(hwnd)) {
		DWORD dwProcId;
		GetWindowThreadProcessId(hwnd, &dwProcId);
		if (g_dwProcId == dwProcId) {
			*((LPBOOL)lParam) = FALSE;
			return FALSE;
		}
	}
	return TRUE;
}

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	g_nCmdShow = nCmdShow;

 	// TODO: Place code here.
	::OleInitialize(NULL);
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDC_TCM, szWindowClass, MAX_LOADSTRING);
	LoadString(hInstance, IDC_MESSAGE, szMessage, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_TCM));

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
		try {
			::OleUninitialize();
	}
	catch (...) {
	}
	return (int) msg.wParam;
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TCM));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_TCM);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
HWND InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance; // Store instance handle in our global variable
	g_hWnd = CreateWindow(szWindowClass, TCM_ABOUT, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);
	GetWindowThreadProcessId(g_hWnd, &g_dwProcId);
	return g_hWnd;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message)
	{
	case WM_COMMAND: 
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// Parse the menu selections:
		switch (wmId) {
			case IDM_EXIT:
				DestroyWindow(hWnd);
				break;
			case IDM_URL:
				ShellExecute(hWnd, NULL, TCM_URL, NULL, NULL, SW_SHOWNORMAL);
				break;
			default:
				return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_CREATE:
		GetWindowRect(hWnd, &g_rc);
		MoveWindow(hWnd, -32768, -32768, 400, 120, TRUE);
		ShowWindow(hWnd, g_nCmdShow);
        DragAcceptFiles(hWnd, TRUE);
		 
		int nArgs;
		LPTSTR *lplpszArgs;
		lplpszArgs = CommandLineToArgvW(GetCommandLine(), &nArgs);
		if (!PopupContextMenu(hWnd, lplpszArgs, nArgs, 1)) {
			g_uTimerId2 = SetTimer(hWnd, TCMT_SHOW, 1000, NULL);
		}
		LocalFree(lplpszArgs);
		break;
	case WM_DROPFILES:
		UINT  i, uCount;
		HDROP hDrop;
		WCHAR szFile[MAX_PATH];
		
		hDrop = (HDROP)wParam;
		uCount = DragQueryFile(hDrop, (UINT)-1, NULL, 0);
		lplpszArgs = new LPTSTR[uCount];
		for (i = 0; i < uCount; i++) {
			DragQueryFile(hDrop, i, szFile, MAX_PATH);
			lplpszArgs[i] = SysAllocString(szFile);
		}
		DragFinish(hDrop);
		KillTimer(hWnd, g_uTimerId2);
		ShowWindow(hWnd, SW_HIDE);
		if (!PopupContextMenu(hWnd, lplpszArgs, uCount, 0)) {
			g_uTimerId2 = SetTimer(hWnd, TCMT_SHOW, 1000, NULL);
		}
		for (i = 0; i < uCount; i++) {
			SysFreeString(lplpszArgs[i]);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		RECT rc;
		GetClientRect(hWnd, &rc);
		rc.top += 8;
		rc.left += 8;
		DrawText(hdc, (LPCWSTR)&szMessage, -1, &rc, DT_LEFT);
		rc.top += 24;
		SetTextColor(hdc, 0xff0000);
		DrawText(hdc, TCM_URL, -1, &rc, DT_LEFT);
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	case WM_TIMER:
		if (wParam == TCMT_CLOSE) {
			if (g_nCount > 0) {
				g_nCount--;
			}
			else {
//				if (FindWindow(L"StubWindow32", NULL) == 0) {	//Against property
				BOOL b = TRUE;
				EnumWindows(EnumWindowsProc , (LPARAM)&b);
				if (b) {
					KillTimer(hWnd, g_uTimerId);
					DestroyWindow(hWnd);
				}
			}
		}
		if (wParam == TCMT_SHOW) {
			KillTimer(hWnd, g_uTimerId2);
			ShowWindow(hWnd, g_nCmdShow);
			MoveWindow(hWnd, g_rc.left, g_rc.top, 400, 120, TRUE);
			UpdateWindow(hWnd);
		}
		break;
	//Menu
	case WM_MEASUREITEM:
	case WM_INITMENUPOPUP:
	case WM_DRAWITEM:
	case WM_MENUSELECT:
	case WM_HELP:
	case WM_MENUCHAR:
	case WM_INITMENU:
		LRESULT lResult;
		lResult = 0;
		if (g_pCCM) {
			if (g_pCCM->m_pContextMenu3) {
				g_pCCM->m_pContextMenu3->HandleMenuMsg2(message, wParam, lParam, &lResult);
			}
			else if (g_pCCM->m_pContextMenu2) {
				g_pCCM->m_pContextMenu2->HandleMenuMsg(message, wParam, lParam);
			}
		}
		return DefWindowProc(hWnd, message, wParam, lParam);
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

//CtcmContextMenu

CtcmContextMenu::CtcmContextMenu(IContextMenu *pContextMenu, HWND hwnd)
{
	m_cRef = 1;
	m_pContextMenu = NULL;
	m_pContextMenu2 = NULL;
	m_pContextMenu3 = NULL;
	m_hwnd = hwnd;
	
	if (pContextMenu) {
		if SUCCEEDED(pContextMenu->QueryInterface(IID_PPV_ARGS(&m_pContextMenu))) {
			pContextMenu->QueryInterface(IID_PPV_ARGS(&m_pContextMenu2));
			pContextMenu->QueryInterface(IID_PPV_ARGS(&m_pContextMenu3));
		}
	}
}

CtcmContextMenu::~CtcmContextMenu()
{
	if (m_pContextMenu3) {
		m_pContextMenu3->Release();
	}
	if (m_pContextMenu2) {
		m_pContextMenu2->Release();
	}
	if (m_pContextMenu) {
		m_pContextMenu->Release();
	}
	g_pCCM = NULL;
	g_uTimerId = SetTimer(m_hwnd, TCMT_CLOSE, 1000, NULL);
}

STDMETHODIMP CtcmContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IContextMenu)) {
		*ppvObject = static_cast<IContextMenu *>(this);
	}
	else if (IsEqualIID(riid, IID_IContextMenu2)) {
		return m_pContextMenu->QueryInterface(IID_IContextMenu2, ppvObject);
	}
	else if (IsEqualIID(riid, IID_IContextMenu3)) {
		return m_pContextMenu->QueryInterface(IID_IContextMenu3, ppvObject);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CtcmContextMenu::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CtcmContextMenu::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CtcmContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	return m_pContextMenu->QueryContextMenu(hmenu, indexMenu, idCmdFirst, idCmdLast, uFlags);
}

STDMETHODIMP CtcmContextMenu::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	return m_pContextMenu->GetCommandString(idCmd, uFlags, pwReserved, pszName, cchMax);
}

STDMETHODIMP CtcmContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	return m_pContextMenu->InvokeCommand(pici);
}

// CtcmDataObject

CtcmDataObject::CtcmDataObject(LPITEMIDLIST *apidl, int nArgs, HWND hwnd)
{
	m_cRef = 1;
	m_pidllist = apidl;
	m_nCount = nArgs;
	m_hwnd = hwnd;
}
CtcmDataObject::~CtcmDataObject()
{
}

STDMETHODIMP CtcmDataObject::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDataObject)) {
		*ppvObject = reinterpret_cast<IDataObject *>(this);
	}
	else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP_(ULONG) CtcmDataObject::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CtcmDataObject::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

HDROP CtcmDataObject::GethDrop(int x, int y, BOOL fNC, BOOL bSpecial)
{
	BSTR *pbslist = new BSTR[m_nCount];
	UINT uSize = sizeof(WCHAR);
	for (int i = m_nCount - 1; i >= 0; i--) {
		WCHAR szPath[MAX_PATH];
		pbslist[i] = NULL;
		if (SHGetPathFromIDList(m_pidllist[i], szPath)) {
			if (szPath[0]) {
				pbslist[i] = SysAllocString(szPath);
				uSize += SysStringByteLen(pbslist[i]) + sizeof(WCHAR);
			}
		}
	}
	HDROP hDrop = (HDROP)GlobalAlloc(GHND, sizeof(DROPFILES) + uSize);
	LPDROPFILES lpDropFiles = (LPDROPFILES)GlobalLock(hDrop);
	try {
		lpDropFiles->pFiles = sizeof(DROPFILES);
		lpDropFiles->pt.x = x;
		lpDropFiles->pt.y = y;
		lpDropFiles->fNC = fNC;
		lpDropFiles->fWide = TRUE;

		LPWSTR lp = (LPWSTR)&lpDropFiles[1];
		for (int i = 0; i< m_nCount; i++) {
			if (pbslist[i]) {
				lstrcpy(lp, pbslist[i]);
				lp += (SysStringByteLen(pbslist[i]) / 2) + 1;
				::SysFreeString(pbslist[i]);
			}
		}
		*lp = 0;
		delete [] pbslist;
	}
	catch (...) {
	}
	GlobalUnlock(hDrop);
	return hDrop;
}

//IDataObject
STDMETHODIMP CtcmDataObject::GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
{
	if (pformatetcIn->cfFormat == CF_HDROP) {
		pmedium->tymed = TYMED_HGLOBAL;
		pmedium->hGlobal = (HGLOBAL)GethDrop(0, 0, FALSE, FALSE);
		pmedium->pUnkForRelease = NULL;
		return S_OK;
	}
	return DATA_E_FORMATETC;
}

STDMETHODIMP CtcmDataObject::GetDataHere(FORMATETC *pformatetc, STGMEDIUM *pmedium)
{
	return E_NOTIMPL;
}

STDMETHODIMP CtcmDataObject::QueryGetData(FORMATETC *pformatetc)
{
	if (pformatetc->cfFormat == CF_HDROP) {
		return S_OK;
	}
	return DATA_E_FORMATETC;
}

STDMETHODIMP CtcmDataObject::GetCanonicalFormatEtc(FORMATETC *pformatectIn, FORMATETC *pformatetcOut)
{
	return E_NOTIMPL;
}

STDMETHODIMP CtcmDataObject::SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease)
{
	return E_NOTIMPL;
}

STDMETHODIMP CtcmDataObject::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc)
{
	if (dwDirection == DATADIR_GET) {
		FORMATETC formats[] = { HDROPFormat };
		return CreateFormatEnumerator(1, formats, ppenumFormatEtc);
	}
	return E_NOTIMPL;
}

STDMETHODIMP CtcmDataObject::DAdvise(FORMATETC *pformatetc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	return E_NOTIMPL;
}

STDMETHODIMP CtcmDataObject::DUnadvise(DWORD dwConnection)
{
	return E_NOTIMPL;
}

STDMETHODIMP CtcmDataObject::EnumDAdvise(IEnumSTATDATA **ppenumAdvise)
{
	return E_NOTIMPL;
}
