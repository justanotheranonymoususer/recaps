/******************************************************************************

Recaps - change language and keyboard layout using the CapsLock key.
Copyright (C) 2007 Eli Golovinsky

-------------------------------------------------------------------------------

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

*******************************************************************************/

#include "stdafx.h"
#include "resource.h"
#include "trayicon.h"
#include "fixlayouts.h"
#include "utils.h"

#define HELP_MESSAGE \
	L"Recaps allows you to quickly switch the current\n"\
	L"language using the Capslock key.\n"\
	L"\n"\
	L"Capslock changes between the chosen pair of keyboard laguanges.\n"\
	L"Alt+Capslock changes the chosen pair of keyboard languages.\n"\
	L"Ctrl+Capslock fixes text you typed in the wrong laguange.\n"\
	L"* If both Ctrl keys (left and right) are pressed, only the selected text will be fixed.\n"\
	L"Shift+Capslock is the old Capslock that lets you type in CAPITAL.\n"\
	L"\n"\
	L"http://www.gooli.org/blog/recaps\n\n"\
	L"Eli Golovinsky, Israel 2008"

#define HELP_TITLE \
    L"Recaps 0.7 alpha - Retake your Capslock!"

// General constants
#define MAXLEN 1024
#define MAX_LAYOUTS 256
#define MUTEX L"recaps-D3E743A3-E0F9-47f5-956A-CD15C6548789"
#define WINDOWCLASS_NAME L"RECAPS"
#define TITLE L"Recaps"

// Tray icon constants
#define ID_TRAYICON          1
#define APPWM_TRAYICON       WM_APP

// Our commands
#define ID_ABOUT             2000
#define ID_EXIT              2001
#define ID_MAIN_LANG         2002
#define ID_LANG              (2002 + MAX_LAYOUTS)

// Custom messages
#define APPWM_LANG_ACTION    (WM_APP + 1)

typedef enum
{
	LANG_ACTION_NONE,
	LANG_ACTION_SWITCH_LAYOUT,
	LANG_ACTION_SWITCH_PAIR,
	LANG_ACTION_CONVERT_ALL_TEXT,
	LANG_ACTION_CONVERT_SELECTED_TEXT,
} LangAction;

typedef struct
{
	WCHAR names[MAX_LAYOUTS][MAXLEN];
	HKL   hkls[MAX_LAYOUTS];
	UINT  count;
	UINT  main;
	UINT  paired;
} KeyboardLayoutInfo;

KeyboardLayoutInfo g_keyboardInfo;
BOOL g_bShowTrayIcon;
BOOL g_bModalShown;
HHOOK g_hKeyboardHook;
UINT g_uTaskbarRestart;
HWND g_hMainWnd;
HANDLE g_hKeyboardHookThread;
DWORD g_dwKeyboardThreadId;

LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
int OnTrayIcon(HWND hWnd, WPARAM wParam, LPARAM lParam);
int OnCommand(HWND hWnd, WORD wID, HWND hCtl);
BOOL ShowPopupMenu(HWND hWnd);

void GetKeyboardLayouts(KeyboardLayoutInfo* info);
void LoadConfiguration(KeyboardLayoutInfo* info);
void SaveConfiguration(const KeyboardLayoutInfo* info);

HWND RemoteGetFocus();
HKL GetWindowLayout(HWND hWnd);
HKL GetCurrentLayout();
HKL SwitchLayout(HWND hWnd, HKL hkl);
HKL SwitchToPairedLayout();
HKL SwitchPair();
void SwitchAndConvert(BOOL bOnlySelected);
BOOL KeyboardHookInit();
void KeyboardHookUninit();
DWORD WINAPI KeyboardHookThread(LPVOID pParameter);
LRESULT CALLBACK LowLevelKeyboardHookProc(int nCode, WPARAM wParam, LPARAM lParam);

///////////////////////////////////////////////////////////////////////////////
// Program's entry point
int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
	UNREFERENCED_PARAMETER(hInstance);
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	UNREFERENCED_PARAMETER(nCmdShow);

	// Prevent from two copies of Recaps from running at the same time
	HANDLE mutex = CreateMutex(NULL, FALSE, MUTEX);
	DWORD result = WaitForSingleObject(mutex, 0);
	if(result == WAIT_TIMEOUT)
	{
		CloseHandle(mutex);
		MessageBox(NULL, L"Recaps is already running.", L"Recaps", MB_OK | MB_ICONINFORMATION);
		return 1;
	}

	// Initialize
	GetKeyboardLayouts(&g_keyboardInfo);
	LoadConfiguration(&g_keyboardInfo);
	g_bShowTrayIcon = !DoesCmdLineSwitchExists(L"-no_icon");

	// Create a fake window to listen to events
	WNDCLASSEX wclx = { 0 };
	wclx.cbSize = sizeof(wclx);
	wclx.lpfnWndProc = WindowProc;
	wclx.hInstance = hInstance;
	wclx.lpszClassName = WINDOWCLASS_NAME;
	RegisterClassEx(&wclx);
	g_hMainWnd = CreateWindow(WINDOWCLASS_NAME, NULL, 0, 0, 0, 0, 0, NULL, 0, hInstance, NULL);

	// Handle messages
	MSG msg;
	while(GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	// Clean up
	UnregisterClass(WINDOWCLASS_NAME, hInstance);
	SaveConfiguration(&g_keyboardInfo);
	CloseHandle(mutex);

	return 0;
}

///////////////////////////////////////////////////////////////////////////////
// Handles events at the window (both hot key and from the tray icon)
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch(uMsg)
	{
	case WM_CREATE:
		g_uTaskbarRestart = RegisterWindowMessage(L"TaskbarCreated");
		if(g_bShowTrayIcon)
		{
			AddTrayIcon(hWnd, 0, APPWM_TRAYICON, IDI_MAINFRAME, TITLE);
		}

		// Set hook to capture CapsLock
		KeyboardHookInit();
		return 0;

	case APPWM_TRAYICON:
		return OnTrayIcon(hWnd, wParam, lParam);

	case WM_COMMAND:
		return OnCommand(hWnd, LOWORD(wParam), (HWND)lParam);

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		if(g_bShowTrayIcon)
		{
			RemoveTrayIcon(hWnd, 0);
		}
		KeyboardHookUninit();
		PostQuitMessage(0);
		return 0;

	case APPWM_LANG_ACTION:
		switch(wParam)
		{
		case LANG_ACTION_SWITCH_LAYOUT:
			SwitchToPairedLayout();
			break;

		case LANG_ACTION_SWITCH_PAIR:
			SwitchPair();
			break;

		case LANG_ACTION_CONVERT_ALL_TEXT:
			SwitchAndConvert(FALSE);
			break;

		case LANG_ACTION_CONVERT_SELECTED_TEXT:
			SwitchAndConvert(TRUE);
			break;
		}
		return 0;

	default:
		if(uMsg == g_uTaskbarRestart)
		{
			if(g_bShowTrayIcon)
			{
				AddTrayIcon(hWnd, 0, APPWM_TRAYICON, IDI_MAINFRAME, TITLE);
			}
			return 0;
		}

		return DefWindowProc(hWnd, uMsg, wParam, lParam);
	}
}

///////////////////////////////////////////////////////////////////////////////
// Create and display a popup menu when the user right-clicks on the icon
int OnTrayIcon(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(wParam);

	if(g_bModalShown == TRUE)
		return 0;

	switch(lParam)
	{
	case WM_RBUTTONUP:
		// Show the context menu
		ShowPopupMenu(hWnd);
		return 0;
	}
	return 0;
}

///////////////////////////////////////////////////////////////////////////////
// Handles user commands from the menu
int OnCommand(HWND hWnd, WORD wID, HWND hCtl)
{
	UNREFERENCED_PARAMETER(hCtl);

	// Have a look at the command and act accordingly
	if(wID == ID_EXIT)
	{
		PostMessage(hWnd, WM_CLOSE, 0, 0);
	}
	else if(wID == ID_ABOUT)
	{
		MessageBox(NULL, HELP_MESSAGE, HELP_TITLE, MB_OK | MB_ICONINFORMATION);
	}
	else if(wID >= ID_MAIN_LANG && wID < ID_MAIN_LANG + MAX_LAYOUTS)
	{
		UINT newMainLayout = wID - ID_MAIN_LANG;
		if(newMainLayout == g_keyboardInfo.paired)
			g_keyboardInfo.paired = g_keyboardInfo.main;

		g_keyboardInfo.main = newMainLayout;
		SaveConfiguration(&g_keyboardInfo);
	}
	else if(wID >= ID_LANG && wID < ID_LANG + MAX_LAYOUTS)
	{
		g_keyboardInfo.paired = wID - ID_LANG;
		SaveConfiguration(&g_keyboardInfo);
	}

	return 0;
}

///////////////////////////////////////////////////////////////////////////////
// Create and display a popup menu when the user right-clicks on the icon
BOOL ShowPopupMenu(HWND hWnd)
{
	// Create a submenu for the main locale
	HMENU hMainLocalePop = CreatePopupMenu();

	// Add items for the languages
	for(UINT layout = 0; layout < g_keyboardInfo.count; layout++)
	{
		AppendMenu(hMainLocalePop, MF_STRING, ID_MAIN_LANG + layout, g_keyboardInfo.names[layout]);
	}

	// Check the main language
	CheckMenuRadioItem(hMainLocalePop, ID_MAIN_LANG, ID_MAIN_LANG + g_keyboardInfo.count - 1, 
		ID_MAIN_LANG + g_keyboardInfo.main, MF_BYCOMMAND);

	// Create the main popup menu
	HMENU hPop = CreatePopupMenu();
	AppendMenu(hPop, MF_STRING, ID_ABOUT, L"Help...");
	AppendMenu(hPop, MF_SEPARATOR, 0, NULL);
	AppendMenu(hPop, MF_POPUP, (UINT_PTR)hMainLocalePop, L"Main language");
	AppendMenu(hPop, MF_SEPARATOR, 0, NULL);

	// Add pairs of items for the languages
	for(UINT layout = 0; layout < g_keyboardInfo.count; layout++)
	{
		if(layout == g_keyboardInfo.main)
			continue;

		WCHAR szBuffer[MAXLEN * 2 + 16];
		swprintf_s(szBuffer, sizeof(szBuffer) / sizeof(WCHAR), L"%s <=> %s",
			g_keyboardInfo.names[g_keyboardInfo.main], g_keyboardInfo.names[layout]);

		AppendMenu(hPop, MF_STRING, ID_LANG + layout, szBuffer);
	}

	// Check the paired language
	CheckMenuRadioItem(hPop, ID_LANG, ID_LANG + g_keyboardInfo.count - 1, 
		ID_LANG + g_keyboardInfo.paired, MF_BYCOMMAND);

	AppendMenu(hPop, MF_SEPARATOR, 0, NULL);
	AppendMenu(hPop, MF_STRING, ID_EXIT, L"Exit");

	// Show the menu

	// See http://support.microsoft.com/kb/135788 for the reasons 
	// for the SetForegroundWindow and Post Message trick.
	POINT curpos;
	GetCursorPos(&curpos);
	SetForegroundWindow(hWnd);
	g_bModalShown = TRUE;
	UINT cmd = TrackPopupMenu(
		hPop, TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY,
		curpos.x, curpos.y, 0, hWnd, NULL
		);
	PostMessage(hWnd, WM_NULL, 0, 0);
	g_bModalShown = FALSE;

	// Send a command message to the window to handle the menu item the user chose
	if(cmd)
		SendMessage(hWnd, WM_COMMAND, cmd, 0);

	DestroyMenu(hPop);

	return cmd != 0;
}

///////////////////////////////////////////////////////////////////////////////
// Fills ``info`` with the currently installed keyboard layouts
// Based on http://blogs.msdn.com/michkap/archive/2004/12/05/275231.aspx.
void GetKeyboardLayouts(KeyboardLayoutInfo* info)
{
	BOOL mainWasChosen = FALSE;
	memset(info, 0, sizeof(KeyboardLayoutInfo));
	info->count = GetKeyboardLayoutList(MAX_LAYOUTS, info->hkls);
	for(UINT i = 0; i < info->count; i++)
	{
		LANGID language = LOWORD(info->hkls[i]);
		LCID locale = MAKELCID(language, SORT_DEFAULT);
		GetLocaleInfo(locale, LOCALE_SLANGUAGE, info->names[i], MAXLEN);

		// Prefer English as the default main language
		if(!mainWasChosen && language == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US))
		{
			info->main = i;
			if(i == 0 && info->count >= 2)
				info->paired = 1;
			mainWasChosen = TRUE;
		}
	}

	if(!mainWasChosen && info->count >= 2)
		info->paired = 1;
}

///////////////////////////////////////////////////////////////////////////////
// Load currently active keyboard layouts from the registry
void LoadConfiguration(KeyboardLayoutInfo* info)
{
	HKEY hkey;
	LONG result;

	result = RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\Recaps", 0, KEY_QUERY_VALUE, &hkey);

	// Load main and paired language
	if(result == ERROR_SUCCESS)
	{
		WCHAR localeName[MAXLEN];
		DWORD length;

		length = sizeof(localeName);
		result = RegGetValue(hkey, NULL, L"main", RRF_RT_REG_SZ, NULL, localeName, &length);
		if(result == ERROR_SUCCESS)
		{
			for(UINT i = 0; i < info->count; i++)
			{
				if(wcscmp(localeName, info->names[i]) == 0)
				{
					info->main = i;
					break;
				}
			}
		}

		length = sizeof(localeName);
		result = RegGetValue(hkey, NULL, L"paired", RRF_RT_REG_SZ, NULL, localeName, &length);
		if(result == ERROR_SUCCESS)
		{
			for(UINT i = 0; i < info->count; i++)
			{
				if(wcscmp(localeName, info->names[i]) == 0)
				{
					info->paired = i;
					break;
				}
			}
		}

		if(info->main == info->paired && info->count >= 2)
			info->paired = (info->main == 0) ? 1 : 0;

		RegCloseKey(hkey);
	}
}

///////////////////////////////////////////////////////////////////////////////
// Saves currently active keyboard layouts to the registry
void SaveConfiguration(const KeyboardLayoutInfo* info)
{
	HKEY hkey;
	LONG result;

	result = RegCreateKeyEx(HKEY_CURRENT_USER, L"Software\\Recaps", 0, NULL, 0, KEY_SET_VALUE, NULL, &hkey, NULL);

	// Save main and paired language
	if(result == ERROR_SUCCESS)
	{
		const WCHAR *pLocaleName;

		pLocaleName = info->names[info->main];
		RegSetValueEx(hkey, L"main", 0, REG_SZ, (const BYTE *)(pLocaleName), (DWORD)((wcslen(pLocaleName) + 1) * sizeof(WCHAR)));

		pLocaleName = info->names[info->paired];
		RegSetValueEx(hkey, L"paired", 0, REG_SZ, (const BYTE *)(pLocaleName), (DWORD)((wcslen(pLocaleName) + 1) * sizeof(WCHAR)));

		RegCloseKey(hkey);
	}
}

///////////////////////////////////////////////////////////////////////////////
// Finds out which window has the focus
HWND RemoteGetFocus()
{
	GUITHREADINFO remoteThreadInfo;
	remoteThreadInfo.cbSize = sizeof(GUITHREADINFO);
	if(!GetGUIThreadInfo(0, &remoteThreadInfo))
	{
		return NULL;
	}

	return remoteThreadInfo.hwndFocus ? remoteThreadInfo.hwndFocus : remoteThreadInfo.hwndActive;
}

///////////////////////////////////////////////////////////////////////////////
// Returns the current layout in the active window
HKL GetWindowLayout(HWND hWnd)
{
	DWORD threadId = GetWindowThreadProcessId(hWnd, NULL);
	return GetKeyboardLayout(threadId);
}

///////////////////////////////////////////////////////////////////////////////
// Returns the current layout in the active window
HKL GetCurrentLayout()
{
	HWND hWnd = RemoteGetFocus();
	if(!hWnd)
		return NULL;

	return GetWindowLayout(hWnd);
}

///////////////////////////////////////////////////////////////////////////////
// Switches the current language
HKL SwitchLayout(HWND hWnd, HKL hkl)
{
	BOOL bBuggy = FALSE;

	HWND hRootOwnerWnd = GetAncestor(hWnd, GA_ROOTOWNER);

	WCHAR szClassName[256];
	if(hRootOwnerWnd && GetClassName(hRootOwnerWnd, szClassName, _countof(szClassName)))
	{
		// Skype and Word hang when posting WM_INPUTLANGCHANGEREQUEST.
		if(wcscmp(szClassName, L"tSkMainForm") == 0 ||
			wcscmp(szClassName, L"TConversationForm") == 0 ||
			wcscmp(szClassName, L"OpusApp") == 0)
		{
			bBuggy = TRUE;
		}
	}

	if(bBuggy)
	{
		// A workaround for apps which don't support WM_INPUTLANGCHANGEREQUEST.
		for(UINT i = 0; i < g_keyboardInfo.count; i++)
		{
			HKL currentLayout = GetWindowLayout(hWnd);
			if(currentLayout == hkl)
				break;

			// Change layout by simulating Alt+Shift.
			SendAltShift();

			// Wait for the change to apply.
			for(UINT j = 0; j < 10; j++)
			{
				if(GetWindowLayout(hWnd) != currentLayout)
					break;

				Sleep(30);
			}
		}
	}
	else
	{
		PostMessage(hWnd, WM_INPUTLANGCHANGEREQUEST, 0, (LPARAM)hkl);
	}

	return hkl;
}

///////////////////////////////////////////////////////////////////////////////
// Switches the current language to the other language of the current pair
HKL SwitchToPairedLayout()
{
	HWND hWnd = RemoteGetFocus();
	if(!hWnd)
		return NULL;

	HKL currentLayout = GetWindowLayout(hWnd);

	// Find the current keyboard layout's index
	UINT i;
	for(i = 0; i < g_keyboardInfo.count; i++)
	{
		if(g_keyboardInfo.hkls[i] == currentLayout)
			break;
	}
	UINT currentLanguageIndex = i;

	// Decide the new layout
	UINT newLanguage;
	if(currentLanguageIndex == g_keyboardInfo.main)
		newLanguage = g_keyboardInfo.paired;
	else
		newLanguage = g_keyboardInfo.main;

	// Activate the new language
	SwitchLayout(hWnd, g_keyboardInfo.hkls[newLanguage]);

#ifdef _DEBUG
	PrintDebugString("Language set to %S", g_keyboardInfo.names[newLanguage]);
#endif

	return g_keyboardInfo.hkls[newLanguage];
}

///////////////////////////////////////////////////////////////////////////////
// Switches the language pair
HKL SwitchPair()
{
	// Find the current keyboard layout's index
	UINT newPaired = g_keyboardInfo.paired;
	newPaired = (newPaired + 1) % g_keyboardInfo.count;
	if(newPaired == g_keyboardInfo.main)
	{
		newPaired = (newPaired + 1) % g_keyboardInfo.count;
	}

	g_keyboardInfo.paired = newPaired;

	HWND hWnd = RemoteGetFocus();
	if(hWnd)
	{
		SwitchLayout(hWnd, g_keyboardInfo.hkls[newPaired]);
	}

	SaveConfiguration(&g_keyboardInfo);

	return g_keyboardInfo.hkls[newPaired];
}

///////////////////////////////////////////////////////////////////////////////
// Selects the entire current line and converts it to the current keyboard layout
void SwitchAndConvert(BOOL bOnlySelected)
{
	if(!bOnlySelected)
	{
		SendKeyCombo('A', TRUE, FALSE, FALSE);
	}

	HKL sourceLayout = GetCurrentLayout();
	HKL targetLayout = SwitchToPairedLayout();
	if(sourceLayout && targetLayout)
	{
		ConvertSelectedTextInActiveWindow(sourceLayout, targetLayout);
	}
}

///////////////////////////////////////////////////////////////////////////////
// Creates a thread, and initializes the keyboard hook inside it
BOOL KeyboardHookInit()
{
	BOOL bSuccess = FALSE;

	HANDLE hThreadReadyEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
	if(hThreadReadyEvent)
	{
		g_hKeyboardHookThread = CreateThread(NULL, 0, KeyboardHookThread,
			(void*)hThreadReadyEvent, CREATE_SUSPENDED, &g_dwKeyboardThreadId);
		if(g_hKeyboardHookThread)
		{
			SetThreadPriority(g_hKeyboardHookThread, THREAD_PRIORITY_TIME_CRITICAL);
			ResumeThread(g_hKeyboardHookThread);

			WaitForSingleObject(hThreadReadyEvent, INFINITE);

			bSuccess = TRUE;
		}

		CloseHandle(hThreadReadyEvent);
	}

	return bSuccess;
}

///////////////////////////////////////////////////////////////////////////////
// Uninitializes the keyboard hook, and ends the thread that runs it
void KeyboardHookUninit()
{
	HANDLE hThread = InterlockedExchangePointer(&g_hKeyboardHookThread, NULL);
	if(hThread)
	{
		PostThreadMessage(g_dwKeyboardThreadId, WM_APP, 0, 0);
		WaitForSingleObject(hThread, INFINITE);
		CloseHandle(hThread);
	}
}

///////////////////////////////////////////////////////////////////////////////
// The keyboard hook thread
DWORD WINAPI KeyboardHookThread(LPVOID pParameter)
{
	HANDLE hThreadReadyEvent;
	MSG msg;
	BOOL bRet;
	HANDLE hThread;

	hThreadReadyEvent = (HANDLE)pParameter;
	PeekMessage(&msg, NULL, WM_USER, WM_USER, PM_NOREMOVE);
	SetEvent(hThreadReadyEvent);

	g_hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardHookProc, GetModuleHandle(NULL), 0);
	if(g_hKeyboardHook)
	{
		while((bRet = GetMessage(&msg, NULL, 0, 0)) != 0)
		{
			if(bRet == -1)
			{
				msg.wParam = 0;
				break;
			}

			if(msg.hwnd == NULL && msg.message == WM_APP)
			{
				PostQuitMessage(0);
				continue;
			}

			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}

		UnhookWindowsHookEx(g_hKeyboardHook);
	}
	else
		msg.wParam = 0;

	hThread = InterlockedExchangePointer(&g_hKeyboardHookThread, NULL);
	if(hThread)
		CloseHandle(hThread);

	return (DWORD)msg.wParam;
}

///////////////////////////////////////////////////////////////////////////////
// A LowLevelHookProc implementation that captures the CapsLock key
LRESULT CALLBACK LowLevelKeyboardHookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if(nCode != HC_ACTION)
		return CallNextHookEx(g_hKeyboardHook, nCode, wParam, lParam);

	KBDLLHOOKSTRUCT* data = (KBDLLHOOKSTRUCT*)lParam;
	BOOL caps = data->vkCode == VK_CAPITAL &&
		(wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN);

	// ignore injected keystrokes
	if(caps && (data->flags & LLKHF_INJECTED) == 0)
	{
		if(GetKeyState(VK_MENU) < 0)
		{
			// Handle Alt+CapsLock - switch current layout pair
			PostMessage(g_hMainWnd, APPWM_LANG_ACTION, LANG_ACTION_SWITCH_PAIR, 0);

			// This call of keybd_event is a workaround for the following issue:
			// Because we disable the WM_KEYDOWN-VK_CAPITAL message, the target
			// window might get the following sequence:
			// 1. WM_KEYDOWN-VK_MENU
			// 2. WM_KEYUP  -VK_MENU
			// 3. WM_KEYUP  -VK_CAPITAL
			// Between 1 and 2 there's the deleted message of WM_KEYDOWN-VK_CAPITAL.
			// Because of this sequence, the target window activates the menu, as
			// if the ALT button was pressed.
			// As a workaround, we send an additional WM_KEYUP-VK_CAPITAL message,
			// so it becomes:
			// 1. WM_KEYDOWN-VK_MENU
			// 2. WM_KEYUP  -VK_CAPITAL
			// 3. WM_KEYUP  -VK_MENU
			// 4. WM_KEYUP  -VK_CAPITAL
			keybd_event(VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0);
			return 1;
		}
		else if(GetKeyState(VK_CONTROL) < 0)
		{
			// Handle Ctrl+CapsLock - switch current layout and convert text in current field.
			// If both ctrl keys are pressed, don't select all text with Ctrl+A before converting.

			BOOL bBothControlsAreDown = GetKeyState(VK_LCONTROL) < 0 && GetKeyState(VK_RCONTROL) < 0;
			WPARAM wAction = bBothControlsAreDown ? LANG_ACTION_CONVERT_SELECTED_TEXT : LANG_ACTION_CONVERT_ALL_TEXT;
			PostMessage(g_hMainWnd, APPWM_LANG_ACTION, wAction, 0);

			return 1; // prevent windows from handling the keystroke
		}
		else if(GetKeyState(VK_SHIFT) < 0)
		{
			// Skip Shift+CapsLock
		}
		else
		{
			// Handle CapsLock - only switch current layout
			PostMessage(g_hMainWnd, APPWM_LANG_ACTION, LANG_ACTION_SWITCH_LAYOUT, 0);
			return 1;
		}
	}

	return CallNextHookEx(g_hKeyboardHook, nCode, wParam, lParam);
}
