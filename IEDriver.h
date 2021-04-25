#pragma once

#define IWB2_GetElementById         0x00000001
#define IWB2_GetElementsByName      0x00000002
#define IWB2_GetElementsByClassName 0x00000003

#define IWB2_ClickElement      0x00000001

BOOL DisableIERedirection(void);
READYSTATE WaitPageLoad(IWebBrowser2 **ppBrowser, DWORD dwTimeOut);
DWORD GetIEThreadProcessId(IWebBrowser2 **ppBrowser);
BOOL IWB2SendRequest(LPCSTR lpUrl, LPCSTR lpAgent, LPCSTR lpData, IWebBrowser2 **ppBrowser, 
	DWORD dwTimeout, BOOL bVisible);
LPVOID IWB2GetFileData(LPCSTR lpUrl, LPCSTR lpAgent, LPCSTR lpData, IWebBrowser2 **ppBrowser,
	DWORD dwTimeout, BOOL bVisible, LPDWORD lpdwFileSize);
BOOL IWB2SetFileData(IWebBrowser2 **ppBrowser, const wchar_t *lpHTML);
BOOL IWB2ExecScript(IWebBrowser2 **ppBrowser, LPCSTR lpScript, DWORD dwTimeOut, VARIANT *pvtRes);
BOOL IWB2ExecScriptNWait(IWebBrowser2 **ppBrowser, LPCSTR lpScript, DWORD dwLoadTimeOut,
	DWORD dwResTimeOut, VARIANT *pvtRes);
BOOL IWB2ElementExists(IWebBrowser2 **ppBrowser, LPCSTR lpName, int nIndex, DWORD dwLoadTimeOut,
	DWORD dwResTimeOut, DWORD dwOpt);
BOOL IWB2Click(IWebBrowser2 **ppBrowser, LPCSTR lpName, int nIndex, DWORD dwTimeOut, DWORD dwOpt);
BOOL IWB2Click2(IWebBrowser2 **ppBrowser, LPCSTR lpName, int nIndex, DWORD dwTimeOut, DWORD dwOpt);
BOOL IWB2GetElementProp(IWebBrowser2 **ppBrowser, LPCSTR lpName, LPCSTR lpProp,
	int nIndex, DWORD dwTimeOut, DWORD dwOpt, VARIANT *pvtRes);
BOOL IWB2SetElementProp(IWebBrowser2 **ppBrowser, LPCSTR lpName, LPCSTR lpProp, LPCSTR lpText,
	int nIndex, DWORD dwTimeOut, DWORD dwOpt);

IHTMLElement *IWB2GetElement(IWebBrowser2 **ppBrowser, LPCSTR lpName, DWORD dwMethod);
IHTMLElement *IWB2GetElementTimeout(IWebBrowser2 **ppBrowser, LPCSTR lpName, DWORD dwMethod, DWORD dwTimeout);
BOOL IWB2FindExecNWait(IWebBrowser2 **ppBrowser, LPCSTR lpName, DWORD dwMethod, DWORD dwAction, DWORD dwTimeout);