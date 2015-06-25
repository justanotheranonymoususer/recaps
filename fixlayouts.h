#pragma once

typedef struct
{
	UINT format;
	HANDLE dataHandle;
} ClipboardFormat;

typedef struct
{
	int count;
	ClipboardFormat* dataArray;
} ClipboardData;


// time in milliseconds to allow the target application
// to execute commands simulated by keystrokes
#define REMOTE_APP_WAIT 100

// The main function that converts the current selected text in the active 
// window from one layout to another.
void ConvertSelectedTextInActiveWindow(HKL hklSource, HKL hklTarget);

// Functions to convert UNICODE strings between keyboard layouts
WCHAR LayoutConvertChar(WCHAR ch, HKL hklSource, HKL hklTarget);
size_t LayoutConvertString(const WCHAR* str, WCHAR* buffer, size_t size, HKL hklSource, HKL hklTarget);
HKL DetectLayoutFromString(const WCHAR* str, BOOL* pmatches);

// Functions to store and restore all of the data in the clipboard
BOOL StoreClipboardData(ClipboardData* formats);
BOOL RestoreClipboardData(ClipboardData* formats);

// Convenience functions for the clipboard
WCHAR* GetClipboardText();
BOOL SetClipboardText(const WCHAR* text);

// Functions that simulate key presses in the current window
void SendKey(BYTE vk, BOOL extended);
void SendKeyCombo(BYTE vkModifier, BYTE vk, BOOL extended);
