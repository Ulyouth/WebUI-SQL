#include "NCBeispiel.h"

BOOL DisableIERedirection(void)
{
	DWORD dwValue = 0x00000000;

	return SetRegValue(HKEY_CURRENT_USER, "Software\\Microsoft\\Edge\\IEToEdge", 
		"RedirectionMode", REG_DWORD, (LPBYTE)&dwValue, sizeof(DWORD));
}

READYSTATE WaitPageLoad(IWebBrowser2 **ppBrowser, DWORD dwTimeOut)
{
	DWORD dwEnd = GetTickCount() + dwTimeOut;
	READYSTATE State;
	HRESULT hr;

	do {
		hr = (*ppBrowser)->lpVtbl->get_ReadyState(*ppBrowser, &State);
	} while (SUCCEEDED(hr) && State != READYSTATE_COMPLETE && dwEnd >= GetTickCount());

	return State;
}

DWORD GetIEThreadProcessId(IWebBrowser2 **ppBrowser)
{
	IServiceProvider *pServProv = 0;
	HRESULT hr = (*ppBrowser)->lpVtbl->QueryInterface(*ppBrowser, &IID_IServiceProvider, (void **)&pServProv);
	
	if (FAILED(hr))
		return 0;

	IOleWindow *pWindow = 0;
	hr = pServProv->lpVtbl->QueryService(pServProv, &SID_SShellBrowser, &IID_IOleWindow, (void **)&pWindow);
	
	if (FAILED(hr))
		goto cleanup;

	HWND hWndBrowser = 0;
	hr = pWindow->lpVtbl->GetWindow(pWindow, &hWndBrowser);
	
	if (FAILED(hr))
		goto cleanup;

	DWORD dwPID = 0;
	GetWindowThreadProcessId(hWndBrowser, &dwPID);

cleanup:
	if (pWindow) pWindow->lpVtbl->Release(pWindow);
	pServProv->lpVtbl->Release(pServProv);

	return dwPID;
}


BOOL IWB2SendRequest(LPCSTR lpUrl, LPCSTR lpAgent, LPCSTR lpData, IWebBrowser2 **ppBrowser, DWORD dwTimeout, 
	BOOL bVisible)
{
	// Checa se já existe uma instância do browser
	if (!*ppBrowser) {
		HRESULT hr = CoCreateInstance(&CLSID_InternetExplorer, 0, CLSCTX_LOCAL_SERVER,
			&IID_IWebBrowser2, ppBrowser);

		if (!ppBrowser)
			return FALSE;
	}

	BOOL bRes = FALSE;
	VARIANT vtEmpty, vtPost, vtHeaders;

	VariantInit(&vtEmpty);
	VariantInit(&vtPost);
	VariantInit(&vtHeaders);

	// Muda o User-Agent do browser
	if (lpAgent) {
		DWORD dwHeadersSize = 0, dwAgentSize = (DWORD)strlen(lpAgent) + 1;
		LPSTR lpHeaders = 0;

		memmerge(&lpHeaders, (size_t *)&dwHeadersSize, "User-Agent: ", 12);
		memmerge(&lpHeaders, (size_t *)&dwHeadersSize, (void *)lpAgent, dwAgentSize);

		vtHeaders.vt = VT_BSTR;
		vtHeaders.bstrVal = charToBStr(lpHeaders);
	}

	if (lpData) {
		vtPost.vt = VT_BSTR;
		vtPost.bstrVal = charToBStr(lpData);
	}

	// Acessa o link desejado
	BSTR bstrUrl = charToBStr(lpUrl);
	HRESULT hr = (*ppBrowser)->lpVtbl->Navigate(*ppBrowser, bstrUrl, &vtEmpty, &vtEmpty, &vtPost, &vtHeaders);

	if (FAILED(hr))
		goto cleanup;

	// Espera até que a página tenha carregado por completo
	bRes = WaitPageLoad(ppBrowser, dwTimeout);

	// Torna o browser visível
	if (bVisible)
		(*ppBrowser)->lpVtbl->put_Visible(*ppBrowser, VARIANT_TRUE);

cleanup:
	if (bstrUrl) SysFreeString(bstrUrl);
	if (vtPost.bstrVal) SysFreeString(vtPost.bstrVal);
	if (vtHeaders.bstrVal) SysFreeString(vtHeaders.bstrVal);

	return bRes;
}

LPVOID IWB2GetFileData(LPCSTR lpUrl, LPCSTR lpAgent, LPCSTR lpData, IWebBrowser2 **ppBrowser, 
	DWORD dwTimeout, BOOL bVisible, LPDWORD lpdwFileSize)
{
	// Envia a requisição
	if (!IWB2SendRequest(lpUrl, lpAgent, lpData, ppBrowser, dwTimeout, bVisible))
		return 0;
	
	// Obtém a página resultante
	LPVOID lpPage = 0;
	IDispatch *pDispatch = 0;
	HRESULT hr = (*ppBrowser)->lpVtbl->get_Document(*ppBrowser, &pDispatch);
	
	if (FAILED(hr))
		return 0;
	
	// Cria um objeto stream para acessar o código fonte da página
	IPersistStreamInit *pStreamInit = 0;
	IStream *pStream = 0;

	hr = pDispatch->lpVtbl->QueryInterface(pDispatch, &IID_IPersistStreamInit, &pStreamInit);
	
	if (FAILED(hr))
		goto cleanup;

	hr = CreateStreamOnHGlobal(0, TRUE, &pStream);
	
	if (FAILED(hr))
		goto cleanup;
	
	// Salva o código fonte no stream
	hr = pStreamInit->lpVtbl->Save(pStreamInit, pStream, FALSE);

	if (FAILED(hr))
		goto cleanup;

	// Obtém o tamanho do código fonte
	STATSTG Stat;
	hr = pStream->lpVtbl->Stat(pStream, &Stat, STATFLAG_NONAME);

	if (FAILED(hr))
		goto cleanup;
	
	// Reserva espaço na memória para o código fonte
	lpPage = malloc((size_t)Stat.cbSize.QuadPart);

	if (!lpPage)
		goto cleanup;

	// Vai para o inicio do stream
	LARGE_INTEGER LIBegin = { 0 };
	hr = pStream->lpVtbl->Seek(pStream, LIBegin, STREAM_SEEK_SET, 0);

	if (FAILED(hr))
		goto cleanup;

	// Lê o conteúdo do stream
	DWORD dwRead = 0;
	hr = pStream->lpVtbl->Read(pStream, lpPage, (ULONG)Stat.cbSize.QuadPart, &dwRead);

	if (FAILED(hr)) {
		free(lpPage); lpPage = 0;
	}

	*lpdwFileSize = (DWORD)Stat.cbSize.QuadPart;

	/*HANDLE hFile = CreateFile("F:\\Programming\\Projects\\CataLista3\\Release\\Logs\\ae.htm", 
		GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
	WriteFile(hFile, lpPage, dwRead, &dwRead, 0);
	CloseHandle(hFile);*/

	/*if (((PBYTE)lpPage)[0] == 0xFF && ((PBYTE)lpPage)[1] == 0xFE) {
		MessageBox(0, "UTF16LE", 0, 0);
	}
	else if (((PBYTE)lpPage)[0] == 0xFE && ((PBYTE)lpPage)[1] == 0xFF) {
		MessageBox(0, "UTF16BE", 0, 0);
	}
	else if (((PBYTE)lpPage)[0] == 0xEF && ((PBYTE)lpPage)[1] == 0xBB && ((PBYTE)lpPage)[2] == 0xBF) {
		MessageBox(0, "UTF8", 0, 0);
	}*/

cleanup:
	if (pStream) pStream->lpVtbl->Release(pStream);
	if (pStreamInit) pStreamInit->lpVtbl->Release(pStreamInit);
	pDispatch->lpVtbl->Release(pDispatch);

	return lpPage;
}

BOOL IWB2SetFileData(IWebBrowser2 **ppBrowser, const wchar_t *lpHTML) 
{
	IDispatch *pDispatch = 0;
	HRESULT hr = (*ppBrowser)->lpVtbl->get_Document(*ppBrowser, &pDispatch);

	if (FAILED(hr))
		return FALSE;

	IHTMLDocument2 *pHTMLDoc = 0;
	IPersistStreamInit *pStreamInit = 0;
	IStream *pStream = 0;
	HGLOBAL hHTMLGlobal = 0;
	BOOL bRes = FALSE;

	hr = pDispatch->lpVtbl->QueryInterface(pDispatch, &IID_IHTMLDocument2, &pHTMLDoc);
	
	if (FAILED(hr))
		goto cleanup;

	hr = pHTMLDoc->lpVtbl->QueryInterface(pHTMLDoc, &IID_IPersistStreamInit, &pStreamInit);

	if (FAILED(hr))
		goto cleanup;

	DWORD dwHTMLSize = wcslen(lpHTML);
	hHTMLGlobal = GlobalAlloc(GMEM_MOVEABLE, (dwHTMLSize + 1) * 2);

	if (!hHTMLGlobal)
		goto cleanup;

	wchar_t *lpLock = GlobalLock(hHTMLGlobal);

	if (!lpLock)
		goto cleanup;

	memcpy(lpLock, lpHTML, (dwHTMLSize + 1) * 2);
	GlobalUnlock(hHTMLGlobal);

	hr = CreateStreamOnHGlobal(hHTMLGlobal, FALSE, &pStream);

	if (FAILED(hr))
		goto cleanup;

	hr = pStreamInit->lpVtbl->InitNew(pStreamInit);

	if (FAILED(hr))
		goto cleanup;

	hr = pStreamInit->lpVtbl->Load(pStreamInit, pStream);

	if (FAILED(hr))
		goto cleanup;

	bRes = TRUE;
	
cleanup:
	if (pStream) pStream->lpVtbl->Release(pStream);
	if (hHTMLGlobal) GlobalFree(hHTMLGlobal);
	if (pStreamInit) pStreamInit->lpVtbl->Release(pStreamInit);
	if (pHTMLDoc) pHTMLDoc->lpVtbl->Release(pHTMLDoc);
	pDispatch->lpVtbl->Release(pDispatch);

	return bRes;
}

BOOL IWB2ExecScript(IWebBrowser2 **ppBrowser, LPCSTR lpScript, DWORD dwTimeOut, VARIANT *pvtRes)
{
	VariantClear(pvtRes);

	IDispatch *pDispatch = 0;
	HRESULT hr = (*ppBrowser)->lpVtbl->get_Document(*ppBrowser, &pDispatch);

	if (FAILED(hr))
		return 0;

	IHTMLDocument2 *pHTMLDoc = 0;
	IHTMLWindow2 *pHTMLWnd = 0;
	IDispatchEx *pDispEx = 0;
	BSTR bstrEval = 0, bstrScript = 0;
	BOOL bRes = FALSE;

	hr = pDispatch->lpVtbl->QueryInterface(pDispatch, &IID_IHTMLDocument2, &pHTMLDoc);

	if (FAILED(hr))
		goto cleanup;

	hr = pHTMLDoc->lpVtbl->get_parentWindow(pHTMLDoc, &pHTMLWnd);

	if (FAILED(hr))
		goto cleanup;

	hr = pHTMLWnd->lpVtbl->QueryInterface(pHTMLWnd, &IID_IDispatchEx, &pDispEx);

	if (FAILED(hr))
		goto cleanup;

	bstrEval = charToBStr("eval");
	bstrScript = charToBStr(lpScript);

	if (!bstrEval || !bstrScript)
		goto cleanup;

	DISPID DispID = -1;
	hr = pDispEx->lpVtbl->GetDispID(pDispEx, bstrEval, fdexNameCaseSensitive, &DispID);

	if (FAILED(hr))
		goto cleanup;

	DISPPARAMS DispParams = { 0 };
	EXCEPINFO ExcepInfo = { 0 };
	UINT nArgErr = -1;
	VARIANT vtParam;

	VariantInit(&vtParam);
	vtParam.vt = VT_BSTR;
	vtParam.bstrVal = bstrScript;

	DispParams.cArgs = 1;
	DispParams.rgvarg = &vtParam;
	DispParams.cNamedArgs = 0;

	hr = pHTMLWnd->lpVtbl->Invoke(pHTMLWnd, DispID, &IID_NULL, 0, DISPATCH_METHOD,
		&DispParams, pvtRes, &ExcepInfo, &nArgErr);

	/*char teste[128];
	sprintf_s(teste, sizeof(teste), "hr=0x%.8X vtRes.vt=%d", hr, pvtRes->vt);
	MessageBox(0, teste, 0, 0);*/

	if (FAILED(hr))
		goto cleanup;

	bRes = TRUE;
	WaitPageLoad(ppBrowser, dwTimeOut);

cleanup:
	if (bstrScript) SysFreeString(bstrScript);
	if (bstrEval) SysFreeString(bstrEval);
	if (pDispEx) pDispEx->lpVtbl->Release(pDispEx);
	if (pHTMLWnd) pHTMLWnd->lpVtbl->Release(pHTMLWnd);
	if (pHTMLDoc) pHTMLDoc->lpVtbl->Release(pHTMLDoc);
	pDispatch->lpVtbl->Release(pDispatch);

	return bRes;
}

BOOL IWB2ExecScriptNWait(IWebBrowser2 **ppBrowser, LPCSTR lpScript, DWORD dwLoadTimeOut,
	DWORD dwResTimeOut, VARIANT *pvtRes)
{
	if (!dwResTimeOut)
		return IWB2ExecScript(ppBrowser, lpScript, dwLoadTimeOut, pvtRes);

	DWORD dwEnd = GetTickCount() + dwResTimeOut;

	do {
		IWB2ExecScript(ppBrowser, lpScript, dwLoadTimeOut, pvtRes);
	} while (!pvtRes->vt && dwEnd >= GetTickCount());

	return (pvtRes->vt) ? TRUE : FALSE;
}

BOOL IWB2ElementExists(IWebBrowser2 **ppBrowser, LPCSTR lpName, int nIndex, DWORD dwLoadTimeOut,
	DWORD dwResTimeOut, DWORD dwOpt)
{
	char lpScript[512];

	switch (dwOpt) {
	case IWB2_GetElementById:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementById('%s')", lpName);
		break;

	case IWB2_GetElementsByName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByName('%s')[%d]", lpName, nIndex);
		break;

	case IWB2_GetElementsByClassName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByClassName('%s')[%d]", lpName, nIndex);
		break;

	default:
		return FALSE;
	}

	VARIANT vtRes;
	return IWB2ExecScriptNWait(ppBrowser, lpScript, dwLoadTimeOut, dwResTimeOut, &vtRes);
}

BOOL IWB2Click(IWebBrowser2 **ppBrowser, LPCSTR lpName, int nIndex, DWORD dwTimeOut, DWORD dwOpt)
{
	char lpScript[512];

	switch (dwOpt) {
	case IWB2_GetElementById:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementById('%s').click()", lpName);
		break;

	case IWB2_GetElementsByName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByName('%s')[%d].click()", lpName, nIndex);
		break;

	case IWB2_GetElementsByClassName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByClassName('%s')[%d].click()", lpName, nIndex);
		break;

	default:
		return FALSE;
	}

	VARIANT vtRes;
	return IWB2ExecScript(ppBrowser, lpScript, dwTimeOut, &vtRes);
}

BOOL IWB2Click2(IWebBrowser2 **ppBrowser, LPCSTR lpName, int nIndex, DWORD dwTimeOut, DWORD dwOpt)
{
	char lpScript[512];

	switch (dwOpt) {
	case IWB2_GetElementById:
		sprintf_s(lpScript, sizeof(lpScript),
			"var e=document.getElementById('%s');var evt1=document.createEvent('MouseEvents');var evt2=document.createEvent('MouseEvents');evt1.initMouseEvent('mousedown',true,false,window,0,0,0,0,0,false,false,false,false,0,null);evt2.initMouseEvent('mouseup',true,false,window,0,0,0,0,0,false,false,false,false,0,null);e.dispatchEvent(evt1);e.dispatchEvent(evt2);", lpName);
		break;

	case IWB2_GetElementsByName:
		sprintf_s(lpScript, sizeof(lpScript),
			"var e=document.getElementsByName('%s')[%d];var evt1=document.createEvent('MouseEvents');var evt2=document.createEvent('MouseEvents');evt1.initMouseEvent('mousedown',true,false,window,0,0,0,0,0,false,false,false,false,0,null);evt2.initMouseEvent('mouseup',true,false,window,0,0,0,0,0,false,false,false,false,0,null);e.dispatchEvent(evt1);e.dispatchEvent(evt2);", lpName, nIndex);
		break;

	case IWB2_GetElementsByClassName:
		sprintf_s(lpScript, sizeof(lpScript),
			"var e=document.getElementsByClassName('%s')[%d];var evt1=document.createEvent('MouseEvents');var evt2=document.createEvent('MouseEvents');evt1.initMouseEvent('mousedown',true,false,window,0,0,0,0,0,false,false,false,false,0,null);evt2.initMouseEvent('mouseup',true,false,window,0,0,0,0,0,false,false,false,false,0,null);e.dispatchEvent(evt1);e.dispatchEvent(evt2);", lpName, nIndex);
		break;

	default:
		return FALSE;
	}

	VARIANT vtRes;
	return IWB2ExecScript(ppBrowser, lpScript, dwTimeOut, &vtRes);
}

BOOL IWB2GetElementProp(IWebBrowser2 **ppBrowser, LPCSTR lpName, LPCSTR lpProp,
	int nIndex, DWORD dwTimeOut, DWORD dwOpt, VARIANT *pvtRes)
{
	char lpScript[512];

	switch (dwOpt) {
	case IWB2_GetElementById:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementById('%s').%s", lpName, lpProp);
		break;

	case IWB2_GetElementsByName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByName('%s')[%d].%s", lpName, nIndex, lpProp);
		break;

	case IWB2_GetElementsByClassName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByClassName('%s')[%d].%s", lpName, nIndex, lpProp);
		break;

	default:
		return FALSE;
	}

	return IWB2ExecScript(ppBrowser, lpScript, dwTimeOut, pvtRes);
}

BOOL IWB2SetElementProp(IWebBrowser2 **ppBrowser, LPCSTR lpName, LPCSTR lpProp, LPCSTR lpText, 
	int nIndex, DWORD dwTimeOut, DWORD dwOpt)
{
	char lpScript[4096];

	switch (dwOpt) {
	case IWB2_GetElementById:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementById('%s').%s='%s'", lpName, lpProp, lpText);
		break;

	case IWB2_GetElementsByName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByName('%s')[%d].%s='%s'", lpName, nIndex, lpProp, lpText);
		break;

	case IWB2_GetElementsByClassName:
		sprintf_s(lpScript, sizeof(lpScript),
			"document.getElementsByClassName('%s')[%d].%s='%s'", lpName, nIndex, lpProp, lpText);
		break;

	default:
		return FALSE;
	}

	VARIANT vtRes;
	return IWB2ExecScript(ppBrowser, lpScript, dwTimeOut, &vtRes);
}

IHTMLElement *IWB2GetElement(IWebBrowser2 **ppBrowser, LPCSTR lpName, DWORD dwMethod)
{
	// Obtém a página atual no navegador
	IDispatch *pDispatch = 0;
	HRESULT hr = (*ppBrowser)->lpVtbl->get_Document(*ppBrowser, &pDispatch);

	if (FAILED(hr))
		return 0;

	IHTMLDocument2 *pHTMLDoc = 0;
	IHTMLElementCollection *pElements = 0;
	IHTMLElement *pElem = 0, *pElemTmp = 0;
	BSTR bstrTarget = 0;

	hr = pDispatch->lpVtbl->QueryInterface(pDispatch, &IID_IHTMLDocument2, (void **)&pHTMLDoc);

	if (FAILED(hr))
		goto cleanup;

	// Obtém a lista completa de elementos contidos na página
	hr = pHTMLDoc->lpVtbl->get_all(pHTMLDoc, &pElements);

	if (FAILED(hr))
		goto cleanup;

	// Obtém a quantidade de elementos
	int nCount = 0;
	hr = pElements->lpVtbl->get_length(pElements, &nCount);

	if (FAILED(hr))
		goto cleanup;

	bstrTarget = charToBStr(lpName);

	if (!bstrTarget)
		goto cleanup;

	for (int x = 0; x < nCount; x++) {
		if (pDispatch) {
			pDispatch->lpVtbl->Release(pDispatch); pDispatch = 0;
		}

		// Define o index do elemento
		VARIANT vtName, vtIndex;

		VariantInit(&vtIndex);
		vtIndex.vt = VT_I4;
		vtIndex.iVal = (dwMethod == IWB2_GetElementsByName) ? 0 : x;

		if (dwMethod == IWB2_GetElementsByName) {
			VariantInit(&vtName);
			vtName.vt = VT_BSTR;
			vtName.bstrVal = bstrTarget;
		}
		else
			VariantCopy(&vtName, &vtIndex);

		// Obtém o elemento atual
		hr = pElements->lpVtbl->item(pElements, vtName, vtIndex, &pDispatch);

		if (FAILED(hr))
			continue;

		if (pElemTmp) {
			pElemTmp->lpVtbl->Release(pElemTmp); pElemTmp = 0;
		}

		hr = pDispatch->lpVtbl->QueryInterface(pDispatch, &IID_IHTMLElement, &pElemTmp);

		if (FAILED(hr))
			continue;

		BSTR bstrId = 0;

		switch (dwMethod) {
		// Identifica o elemento pelo ID
		case IWB2_GetElementById:
			pElemTmp->lpVtbl->get_id(pElemTmp, &bstrId);
			break;

		case IWB2_GetElementsByName:
			bstrId = bstrTarget;
			break;

		// Identifica o elemento pela classe
		case IWB2_GetElementsByClassName:
			pElemTmp->lpVtbl->get_className(pElemTmp, &bstrId);
			break;

		default:
			break;
		}

		if (bstrId) {
			// Checa se o elemento foi encontrado
			if (wcscmp(bstrId, bstrTarget) == 0) {
				pElem = pElemTmp;

				if (dwMethod != IWB2_GetElementsByName)
					SysFreeString(bstrId);

				break;
			}
			else
				SysFreeString(bstrId);
		}
	}

cleanup:
	if (bstrTarget) SysFreeString(bstrTarget);
	if (pElements) pElements->lpVtbl->Release(pElements);
	if (pHTMLDoc) pHTMLDoc->lpVtbl->Release(pHTMLDoc);
	if (pDispatch) pDispatch->lpVtbl->Release(pDispatch);

	return pElem;
}

IHTMLElement *IWB2GetElementTimeout(IWebBrowser2 **ppBrowser, LPCSTR lpName, DWORD dwMethod, DWORD dwTimeout)
{
	DWORD dwEnd = GetTickCount() + dwTimeout;
	IHTMLElement *pElem = 0;

	do {
		pElem = IWB2GetElement(ppBrowser, lpName, dwMethod);
	} while (!pElem && dwEnd >= GetTickCount());

	return pElem;
}

BOOL IWB2FindExecNWait(IWebBrowser2 **ppBrowser, LPCSTR lpName, DWORD dwMethod, DWORD dwAction, DWORD dwTimeout)
{
	IHTMLElement *pElem = 0;

	if (dwTimeout > 0)
		pElem = IWB2GetElementTimeout(ppBrowser, lpName, dwMethod, dwTimeout);
	else
		pElem = IWB2GetElement(ppBrowser, lpName, dwMethod);

	if (!pElem)
		return FALSE;

	BOOL bRes = FALSE;

	switch (dwAction) {
	case IWB2_ClickElement:
		pElem->lpVtbl->click(pElem);
		bRes = TRUE;
		break;

	default:
		break;
	}

	pElem->lpVtbl->Release(pElem);
	return bRes;
}
