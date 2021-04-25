#pragma comment(lib, "Iepmapi.lib")

#ifdef __cplusplus
#include <iostream>
#include <sstream>
#include <string>
#include <locale>
#include <codecvt>
#include <chrono>
#include <thread>
#include <Windows.h>
#include <MsHTML.h>
#include <Exdisp.h>
#include <ExDispid.h>
#include <Iepmapi.h>
#include <shlguid.h>

extern "C" {
#include "sqlite3.h"
#include "Utils.h"
#include "IEDriver.h"
}
#else
#include <stdio.h>
#include <Windows.h>
#include <MsHTML.h>
#include <Exdisp.h>
#include <ExDispid.h>
#include <Iepmapi.h>
#include <shlguid.h>
#include "sqlite3.h"
#include "Utils.h"
#include "IEDriver.h"
#endif
