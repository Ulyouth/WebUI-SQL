#include "NCBeispiel.h"

// Junta 2 buffers em 1
BOOL memmerge(void **buf1, size_t *len1, void *buf2, size_t len2)
{
	// Cria um novo buffer com espaço o suficiente para ambos os buffers
	void *bufex = malloc(*len1 + len2);
	
	if (!bufex){
		return FALSE;
	}

	// Copia o conteúdo do buffer 1 (caso exista) para o novo buffer
	if (*len1){
		memcpy(bufex, *buf1, *len1);
		free(*buf1);
	}

	// Copia o conteúdo do buffer 2 para o novo buffer
	memcpy(((unsigned char *)bufex + *len1), buf2, len2);

	// Substitui os ponteiros do buffer 1 pelos valores do novo buffer
	*buf1 = bufex;
	*len1 += len2;

	return TRUE;
}

// Encontra uma string dentro de outra string
void *memfind(void *buf, size_t buflen, void *target, size_t targetlen)
{
	for (size_t x = 0; x <= (buflen - targetlen); x++) {
		// Checa se a string foi encontrada
		if (memcmp(((unsigned char *)buf + x), target, targetlen) != 0)
			continue;

		return ((unsigned char *)buf + x);
	}

	return 0;
}

// Extrai uma string entre 2 strings
void *memgrab(void *buf, size_t buflen, void *begin, size_t beginlen, void *end, size_t endlen, size_t *grablen)
{
	for (size_t x = 0; x <= (buflen - beginlen - endlen); x++) {
		// Checa se a string inicial foi localizada
		if (memcmp(((unsigned char *)buf + x), begin, beginlen) != 0)
			continue;

		if (!endlen) {
			// Caso uma string final não tenha sido especificada, retorna um ponteiro após a string inicial e o 
			// tamanho restante da string
			*grablen = (buflen - x);

			return ((unsigned char *)buf + x);
		}

		for (x += beginlen, *grablen = 0; *grablen < (buflen - x); *grablen = (*grablen + 1)) {
			// Verifica se a string final foi localizada
			if (memcmp(((unsigned char *)buf + x + *grablen), end, endlen) == 0)
				// Retorna um ponteiro após a string inicial e o tamanho entre a string inicial e final
				return ((unsigned char *)buf + x);
		}

		break;
	}

	return 0;
}

BSTR charToBStr(const char *src)
{
	size_t srclen = strlen(src);
	size_t dstlen = (srclen + 1) * sizeof(wchar_t);
	wchar_t *dst = malloc(dstlen);

	if (!dst)
		return 0;

	size_t outlen;
	mbstowcs_s(&outlen, dst, dstlen / 2, src, srclen);

	BSTR bstr = SysAllocString(dst);
	free(dst);

	return bstr;
}

char *WCharToChar(const wchar_t *wstr)
{
	size_t srclen = wcslen(wstr);
	size_t dstlen = srclen + 1;
	char *dst = malloc(dstlen);

	if (!dst)
		return 0;

	size_t outlen;
	wcstombs_s(&outlen, dst, dstlen, wstr, srclen);

	return dst;
}

BOOL SetRegValue(HKEY hKey, LPCSTR lpSubKey, LPCSTR lpValue, DWORD dwType, LPBYTE lpData, DWORD dwDataSize)
{
	BOOL bRes = TRUE;
	HKEY hSubKey;

	if (RegOpenKeyEx(hKey, lpSubKey, 0, KEY_WOW64_64KEY | KEY_SET_VALUE, &hSubKey) != ERROR_SUCCESS)
		return FALSE;

	if (RegSetValueEx(hSubKey, lpValue, 0, dwType, lpData, dwDataSize) != ERROR_SUCCESS)
		bRes = FALSE;

	RegCloseKey(hSubKey);
	return bRes;
}