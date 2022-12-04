#include "common.h"
#include <strsafe.h>
#include <Windows.h>

void debugOut(const wchar_t *format, ...) {
	const size_t bufsize = 1024;
	wchar_t buf[ bufsize ];
	va_list args;
	va_start( args, format );
	StringCbVPrintf( buf, sizeof(buf), format, args );
	va_end( args );
	OutputDebugString( buf );
}
