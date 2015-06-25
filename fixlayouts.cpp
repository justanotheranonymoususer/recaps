#include "stdafx.h"
#include "fixlayouts.h"
#include "utils.h"

///////////////////////////////////////////////////////////////////////////////
// Converts the text in the active window from one keyboard layout to another
// using the clipboard.
void ConvertSelectedTextInActiveWindow(HKL hklSource, HKL hklTarget)
{
	WCHAR* sourceText = NULL;
	WCHAR* targetText = NULL;
	const WCHAR dummy[] = L"__RECAPS__";

	// store previous clipboard data and set clipboard to dummy string
	ClipboardData prevClipboardData;
	if(!StoreClipboardData(&prevClipboardData))
		return;

	if(!SetClipboardText(dummy))
	{
		// restore the original clipboard data
		RestoreClipboardData(&prevClipboardData);
		return;
	}

	// copy the selected text by simulating Ctrl-C
	SendKeyCombo(VK_CONTROL, 'C', FALSE);

	// wait until copy operation completes and get the copied data from the clipboard
	// this loop has the nice side effect of setting copyOK to FALSE if there's no
	// selected text, since nothing is copied and the clipboard still contains the
	// contents of `dummy`.
	BOOL copyOK = FALSE;
	for(int i = 0; i < 10; i++)
	{
		sourceText = GetClipboardText();
		if(sourceText && wcscmp(sourceText, dummy) != 0)
		{
			copyOK = TRUE;
			break;
		}
		else
		{
			free(sourceText);
			Sleep(30);
		}
	}

	if(copyOK)
	{
		// if the string only matches one particular layout, use it
		// otherwise use the provided layout
		int matches = 0;
		HKL hklDetected = DetectLayoutFromString(sourceText, &matches);
		if(matches == 1)
			hklSource = hklDetected;

		// convert the text between layouts
		size_t length = wcslen(sourceText);
		targetText = (WCHAR*)malloc(sizeof(WCHAR) * (length + 1));
		size_t converted = LayoutConvertString(sourceText, targetText, length + 1, hklSource, hklTarget);

		if(converted)
		{
			// put the converted string on the clipboard
			if(SetClipboardText(targetText))
			{
				// simulate Ctrl-V to paste the text, replacing the previous text
				SendKeyCombo(VK_CONTROL, 'V', FALSE);

				// let the application complete pasting before putting the old data back on the clipboard
				Sleep(REMOTE_APP_WAIT);
			}
		}

		// free allocated memory
		free(sourceText);
		free(targetText);
	}

	// restore the original clipboard data
	RestoreClipboardData(&prevClipboardData);
}

///////////////////////////////////////////////////////////////////////////////
// Converts a character from one keyboard layout to another
WCHAR LayoutConvertChar(WCHAR ch, HKL hklSource, HKL hklTarget)
{
	// special handling for some ambivalent characters in Hebrew layout
	if(LOWORD(hklSource) == MAKELANGID(LANG_HEBREW, SUBLANG_HEBREW_ISRAEL) &&
		LOWORD(hklTarget) == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US))
	{
		switch(ch)
		{
		case L'.':  return L'/';
		case L'/':  return L'q';
		case L'\'': return L'w';
		case L',':  return L'\'';
		}
	}
	else if(LOWORD(hklSource) == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US) &&
		LOWORD(hklTarget) == MAKELANGID(LANG_HEBREW, SUBLANG_HEBREW_ISRAEL))
	{
		switch(ch)
		{
		case L'/':  return L'.';
		case L'q':  return L'/';
		case L'w':  return L'\'';
		case L'\'': return L',';
		}
	}

	// get the virtual key code and the shift state using the character and the source keyboard layout
	SHORT vkAndShift = VkKeyScanEx(ch, hklSource);
	if(vkAndShift == -1)
		return 0; // no such char in source keyboard layout

	BYTE vk = LOBYTE(vkAndShift);
	BYTE shift = HIBYTE(vkAndShift);

	// convert the shift state returned from VkKeyScanEx to an array that represents the
	// key state usable with ToUnicodeEx that we'll be calling next
	BYTE keyState[256] = { 0 };
	if(shift & 1) keyState[VK_SHIFT] = 0x80;	// turn on high bit
	if(shift & 2) keyState[VK_CONTROL] = 0x80;
	if(shift & 4) keyState[VK_MENU] = 0x80;

	// convert virtual key and key state to a new character using the target keyboard layout
	WCHAR buffer[10] = { 0 };
	int result = ToUnicodeEx(vk, 0, keyState, buffer, 10, 0, hklTarget);

	// result can be more than 1 if the character in the source layout is represented by 
	// several characters in the target layout, but we ignore this to simplify the function.
	if(result == 1)
		return buffer[0];

	// conversion failed for some reason
	return 0;
}

///////////////////////////////////////////////////////////////////////////////
// Converts a string from one keyboard layout to another
size_t LayoutConvertString(const WCHAR* str, WCHAR* buffer, size_t size, HKL hklSource, HKL hklTarget)
{
	size_t i;
	for(i = 0; i < wcslen(str) && i < size - 1; i++)
	{
		WCHAR ch = LayoutConvertChar(str[i], hklSource, hklTarget);
		if(ch == 0)
			return 0;
		buffer[i] = ch;
	}
	buffer[i] = '\0';
	return i;
}

///////////////////////////////////////////////////////////////////////////////
// Goes through all the installed keyboard layouts and returns a layout that
// can generate the string. If not matching layout is found, returns NULL.
// If `multiple` isn't NULL it will be set to the number of matched layouts.
HKL DetectLayoutFromString(const WCHAR* str, int* pmatches)
{
	HKL result = NULL;
	HKL* hkls;
	UINT layoutCount;
	layoutCount = GetKeyboardLayoutList(0, NULL);
	hkls = (HKL*)malloc(sizeof(HKL) * layoutCount);
	GetKeyboardLayoutList(layoutCount, hkls);

	int matches = 0;
	for(size_t layout = 0; layout < layoutCount; layout++)
	{
		BOOL validLayout = TRUE;
		for(size_t i = 0; i < wcslen(str); i++)
		{
			UINT vk = VkKeyScanEx(str[i], hkls[layout]);
			if(vk == -1)
			{
				validLayout = FALSE;
				break;
			}
		}
		if(validLayout)
		{
			matches++;
			if(!result)
				result = hkls[layout];
		}
	}

	if(pmatches)
		*pmatches = matches;

	return result;
}

///////////////////////////////////////////////////////////////////////////////
// Stores the clipboard data in all its formats in `formats`.
// You must call FreeAllClipboardData on `formats` when it's no longer needed.
BOOL StoreClipboardData(ClipboardData* formats)
{
	if(!OpenClipboard(NULL))
		return FALSE;

	formats->count = CountClipboardFormats();
	if(formats->count == 0)
	{
		DWORD dwError = GetLastError();
		CloseClipboard();
		return dwError == ERROR_SUCCESS;
	}

	formats->dataArray = (ClipboardFormat*)malloc(sizeof(ClipboardData) * formats->count);
	ZeroMemory(formats->dataArray, sizeof(ClipboardData) * formats->count);
	int i = 0;

	UINT format = EnumClipboardFormats(0);
	while(format)
	{
		if(i > formats->count - 1)
			break;

		HANDLE dataHandle = GetClipboardData(format);
		if(!dataHandle)
			break;

		size_t size;
		if(!(GlobalFlags(dataHandle) & GMEM_DISCARDED))
		{
			size = GlobalSize(dataHandle);
			if(size == 0)
				break;
		}
		else
			size = 0;

		LPVOID source = GlobalLock(dataHandle);
		if(!source)
			break;

		BOOL bCopySucceeded = FALSE;

		formats->dataArray[i].format = format;
		formats->dataArray[i].dataHandle = GlobalAlloc(GHND, size);
		if(formats->dataArray[i].dataHandle)
		{
			if(size == 0)
			{
				bCopySucceeded = TRUE;
			}
			else
			{
				LPVOID dest = GlobalLock(formats->dataArray[i].dataHandle);
				if(dest)
				{
					CopyMemory(dest, source, size);
					GlobalUnlock(formats->dataArray[i].dataHandle);
					bCopySucceeded = TRUE;
				}
			}

			if(!bCopySucceeded)
				GlobalFree(formats->dataArray[i].dataHandle);
		}

		GlobalUnlock(dataHandle);

		if(!bCopySucceeded)
			break;

		i++;

		// next format
		format = EnumClipboardFormats(format);
	}

	CloseClipboard();

	if(format) // if failed before completion
	{
		for(int j = 0; j < i; j++)
		{
			GlobalFree(formats->dataArray[j].dataHandle);
		}

		free(formats->dataArray);
		return FALSE;
	}

	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////
// Restores the data in the clipboard from `formats` that was generated by
// StoreClipboardData, and frees allocated data.
BOOL RestoreClipboardData(ClipboardData* formats)
{
	if(!OpenClipboard(NULL))
		return FALSE;

	EmptyClipboard();
	for(int i = 0; i < formats->count; i++)
	{
		SetClipboardData(formats->dataArray[i].format, formats->dataArray[i].dataHandle);
	}

	CloseClipboard();
	free(formats->dataArray);
	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////
// Gets unicode text from the clipboard. 
// You must free the returned string when you don't need it anymore.
WCHAR* GetClipboardText()
{
	if(!OpenClipboard(NULL))
		return NULL;

	WCHAR* text = NULL;
	HANDLE handle = GetClipboardData(CF_UNICODETEXT);
	if(handle)
	{
		WCHAR* clipboardText = (WCHAR*)GlobalLock(handle);
		if(clipboardText)
		{
			size_t size = sizeof(WCHAR) * (wcslen(clipboardText) + 1);
			text = (WCHAR*)malloc(size);
			if(text)
				memcpy(text, clipboardText, size);

			GlobalUnlock(handle);
		}
	}

	CloseClipboard();
	return text;
}

///////////////////////////////////////////////////////////////////////////////
// Puts unicode text on the clipboard
BOOL SetClipboardText(const WCHAR* text)
{
	if(!OpenClipboard(NULL))
		return FALSE;

	BOOL bSucceeded = FALSE;

	size_t size = sizeof(WCHAR) * (wcslen(text) + 1);
	HANDLE handle = GlobalAlloc(GHND, size);
	if(handle)
	{
		WCHAR* clipboardText = (WCHAR*)GlobalLock(handle);
		if(clipboardText)
		{
			memcpy(clipboardText, text, size);
			GlobalUnlock(handle);
			bSucceeded = EmptyClipboard() &&
				SetClipboardData(CF_UNICODETEXT, handle);
		}

		if(!bSucceeded)
			GlobalFree(handle);
	}

	CloseClipboard();
	return bSucceeded;
}

///////////////////////////////////////////////////////////////////////////////
// Simulates a key press in the active window
void SendKey(BYTE vk, BOOL extended)
{
	keybd_event(vk, 0, extended ? KEYEVENTF_EXTENDEDKEY : 0, 0);
	keybd_event(vk, 0, KEYEVENTF_KEYUP | (extended ? KEYEVENTF_EXTENDEDKEY : 0), 0);
}

///////////////////////////////////////////////////////////////////////////////
// Simulates a key combination (such as Ctrl+X) in the active window
void SendKeyCombo(BYTE vkModifier, BYTE vk, BOOL extended)
{
	BOOL modPressed = GetKeyState(vkModifier) < 0;
	if(!modPressed)
		keybd_event(vkModifier, 0, 0, 0);
	keybd_event(vk, 0, extended ? KEYEVENTF_EXTENDEDKEY : 0, 0);
	if(!modPressed)
		keybd_event(vkModifier, 0, KEYEVENTF_KEYUP, 0);
	keybd_event(vk, 0, KEYEVENTF_KEYUP | (extended ? KEYEVENTF_EXTENDEDKEY : 0), 0);
}
