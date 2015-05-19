#define UNICODE
#pragma comment(lib,"rpcrt4")
#pragma comment(lib,"shlwapi")
#pragma comment(linker,"/opt:nowin98")

#define WM_RUNTIMECHECK WM_APP
#define ID_BUTTON1 (100)

#include<windows.h>
#include<rpcdce.h>
#include<shlwapi.h>
#include"resource.h"

TCHAR szClassName[] = TEXT("Window");

struct RUNTIME_STRUCT
{
	BOOL bIsInstalled;
	INT nResourceID;
	LPTSTR lpszRuntimeName;
};

RUNTIME_STRUCT runtimes[] =
{
	{ FALSE, IDR_DLL_VC6, TEXT("vc6") },
	{ FALSE, IDR_DLL_VC2003, TEXT("vc2003") },
	{ FALSE, IDR_DLL_VC2005, TEXT("vc2005") },
	{ FALSE, IDR_DLL_VC2008, TEXT("vc2008") },
	{ FALSE, IDR_DLL_VC2010, TEXT("vc2010") },
	{ FALSE, IDR_DLL_VC2012, TEXT("vc2012") },
	{ FALSE, IDR_DLL_VC2013, TEXT("vc2013") },
};

BOOL CreateFileFromResource(IN LPCTSTR lpszResourceName, IN LPCTSTR lpszResourceType,
	IN LPCTSTR lpszResFileName)
{
	DWORD dwWritten;
	const HRSRC hRs = FindResource(0, lpszResourceName, lpszResourceType);
	if (!hRs) return FALSE;
	const DWORD dwResSize = SizeofResource(0, hRs);
	if (!dwResSize) return FALSE;
	const HGLOBAL hMem = LoadResource(0, hRs);
	if (!hMem) return FALSE;
	const LPBYTE lpByte = (BYTE *)LockResource(hMem);
	if (!lpByte) return FALSE;
	const HANDLE hFile = CreateFile(lpszResFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
		FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) return FALSE;
	BOOL bReturn = FALSE;
	if (0 != WriteFile(hFile, lpByte, dwResSize, &dwWritten, NULL))
	{
		bReturn = TRUE;
	}
	CloseHandle(hFile);
	return bReturn;
}

BOOL CreateGUID(OUT LPTSTR lpszGUID)
{
	GUID guid = GUID_NULL;
	HRESULT hr = UuidCreate(&guid);
	if (HRESULT_CODE(hr) != RPC_S_OK) return FALSE;
	if (guid == GUID_NULL) return FALSE;
	wsprintf(lpszGUID, TEXT("{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}"),
		guid.Data1, guid.Data2, guid.Data3,
		guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
		guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
	return TRUE;
}

BOOL CreateTempDirectory(OUT LPTSTR pszDir)
{
	TCHAR szGUID[39];
	if (GetTempPath(MAX_PATH, pszDir) == 0)return FALSE;
	if (CreateGUID(szGUID) == 0)return FALSE;
	if (PathAppend(pszDir, szGUID) == 0)return FALSE;
	if (CreateDirectory(pszDir, 0) == 0)return FALSE;
	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HFONT hFont;
	switch (msg)
	{
	case WM_CREATE:
		hFont = CreateFont(18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, TEXT("メイリオ"));
		CreateWindow(TEXT("BUTTON"), TEXT("更新(F5)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP,
			128, 8, 64, 26, hWnd, (HMENU)ID_BUTTON1, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendDlgItemMessage(hWnd, ID_BUTTON1, WM_SETFONT, (WPARAM)hFont, 0);
		SendMessage(hWnd, WM_RUNTIMECHECK, 0, 0);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == ID_BUTTON1)
		{
			SendMessage(hWnd, WM_RUNTIMECHECK, 0, 0);
		}
		break;
	case WM_RUNTIMECHECK:
		{
			const DWORD dwOldErrorMode = SetErrorMode(SEM_FAILCRITICALERRORS);
			TCHAR szTempDirectoryPath[MAX_PATH];
			for (int i = 0; i < sizeof runtimes / sizeof RUNTIME_STRUCT; i++)
			{
				if (CreateTempDirectory(szTempDirectoryPath))
				{
					PathAppend(szTempDirectoryPath, runtimes[i].lpszRuntimeName);
					PathAddExtension(szTempDirectoryPath, TEXT(".dll"));
					if (CreateFileFromResource(MAKEINTRESOURCE(runtimes[i].nResourceID),
						TEXT("DLL"), szTempDirectoryPath))
					{
						const HMODULE hModule = LoadLibrary(szTempDirectoryPath);
						if (hModule)
						{
							FreeLibrary(hModule);
							runtimes[i].bIsInstalled = TRUE;
						}
						DeleteFile(szTempDirectoryPath);
					}
					PathRemoveFileSpec(szTempDirectoryPath);
					RemoveDirectory(szTempDirectoryPath);
				}				
			}
			SetErrorMode(dwOldErrorMode);
			InvalidateRect(hWnd, 0, 1);
		}
		break;
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			const HDC hdc = BeginPaint(hWnd,&ps);
			TCHAR szText[256];
			const DWORD dwOldBkMode = SetBkMode(hdc,TRANSPARENT);
			const HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
			for (int i = 0; i < sizeof runtimes / sizeof RUNTIME_STRUCT; i++)
			{
				wsprintf(szText, TEXT("%s %s"), runtimes[i].lpszRuntimeName,
					runtimes[i].bIsInstalled ? TEXT("O") : TEXT("X"));
				TextOut(hdc, 10, 10+i*32, szText, lstrlen(szText));
			}
			SelectObject(hdc, hOldFont);
			SetBkMode(hdc, dwOldBkMode);
			EndPaint(hWnd, &ps);
		}
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return DefDlgProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	const WNDCLASS wndclass = {
		0,
		WndProc,
		0,
		DLGWINDOWEXTRA,
		hInstance,
		0,
		LoadCursor(0, IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);

	const DWORD dwStyle = WS_CAPTION | WS_SYSMENU |
		WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_CLIPCHILDREN;

	RECT rect = { 0, 0, 200, 10 + 32 * (sizeof runtimes / sizeof RUNTIME_STRUCT) };
	AdjustWindowRect(&rect, dwStyle, 0);

	const HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Visual C++ ランタイム チェッカー"),
		dwStyle,
		CW_USEDEFAULT,
		0,
		rect.right - rect.left,
		rect.bottom - rect.top,
		0,
		0,
		hInstance,
		0
		);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);

	ACCEL Accel[] = { { FVIRTKEY, VK_F5, ID_BUTTON1 } };
	HACCEL hAccel = CreateAcceleratorTable(Accel, sizeof Accel / sizeof ACCEL);

	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!TranslateAccelerator(hWnd, hAccel, &msg) && !IsDialogMessage(hWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	DestroyAcceleratorTable(hAccel);
	return msg.wParam;
}
