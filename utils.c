#include "stdafx.h"

#define BUFSIZE 2048

///////////////////////////////////////////////////////////////////////////////
// Shows a system error message
void ShowError(const WCHAR* message)
{
	WCHAR buffer[BUFSIZE];

	LPVOID errMessage;
	FormatMessage(
		FORMAT_MESSAGE_ALLOCATE_BUFFER |
		FORMAT_MESSAGE_FROM_SYSTEM |
		FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		GetLastError(),
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPTSTR)&errMessage,
		0,
		NULL
		);

	buffer[0] = '\0';
	wcscat_s(buffer, BUFSIZE, message);
	wcscat_s(buffer, BUFSIZE, L"\n\nError : ");
	wcscat_s(buffer, BUFSIZE, (const WCHAR*)errMessage);

	MessageBox(NULL, (const WCHAR*)buffer, L"Recaps Error", MB_OK | MB_ICONINFORMATION);
	LocalFree(errMessage);
}

///////////////////////////////////////////////////////////////////////////////
// Prints an error message to the debugger
void PrintDebugString(const char* format, ...)
{
	char buffer[2048];

	va_list args;
	va_start(args, format);

	vsprintf_s(buffer, 2048, format, args);
	strcat_s(buffer, 2048, "\n");

	OutputDebugStringA(buffer);
}

///////////////////////////////////////////////////////////////////////////////
// Checks whether the requested command line switch exists
BOOL DoesCmdLineSwitchExists(const WCHAR* command)
{
	for(int i = 1; i < __argc; i++)
	{
		if(_wcsicmp(command, __wargv[i]) == 0)
			return TRUE;
	}

	return FALSE;
}
