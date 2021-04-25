#pragma once

BOOL memmerge(void **buf1, size_t *len1, void *buf2, size_t len2);
BSTR charToBStr(const char *src);
char *WCharToChar(const wchar_t *wstr);
BOOL SetRegValue(HKEY hKey, LPCSTR lpSubKey, LPCSTR lpValue, DWORD dwType, LPBYTE lpData, DWORD dwDataSize);