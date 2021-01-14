#include <Windows.h>
#include <objbase.h> 
#include <system_error>
#include <shellapi.h>


wchar_t* utf8_to_wchar(const char* in)
{
    wchar_t* out;
    int len;
    if (in == NULL) {
        return NULL;
    }
    len = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, in, -1, NULL, 0);
    if (len <= 0) {
        return NULL;
    }
    out = (wchar_t*)calloc(len, sizeof(wchar_t));
    if (out == NULL) {
        return NULL;
    }
    if (MultiByteToWideChar(CP_UTF8, 0, in, -1, out, len) == 0) {
        free(out);
        out = NULL;
    }
    return out;
}

char* wchar_to_utf8(const wchar_t* in)

{
    char* out;
    int len;
    if (in == NULL) {
        return NULL;
    }
    len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, NULL, NULL);
    if (len <= 0) {
        return NULL;
    }
    out = (char*)calloc(len, sizeof(char));
    if (out == NULL) {
        return NULL;
    }
    if (WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, NULL, FALSE) == 0) {
        free(out);
        out = NULL;
    }
    return out;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID lpReserved)
{
    //if (reason != DLL_PROCESS_ATTACH)
    //    return true;
    //Do stuff when it gets loaded, messagebox for test
    //MessageBoxA(NULL, "DLLMain!", "We've started.", 0);
    return TRUE;
}

//extern "C" to prevent C++ name mangling

extern "C" __declspec(dllexport) LPSTR StartInstance(LPSTR input) {
    LPSTR retValue;
    const wchar_t* wHelp = L"Starts an instance of a COM object. That's it, literally just starts it, then exits. If you want more functionality, you probably want to use powershell. Other functions will be added later."
        L"Slingshot usage: callmodule -e StartInstance IP {CLSID}"
        L"Rundll32 useage: rundll32.exe,FromRundll IP {CLSID}"
        L"Example: callmodule -e StartInstance 192.168.224.137 {73FDDC80-AEA9-101A-98A7-00AA00374959}"
        L"This creates am instance of the com object in the wordpad COM server on 192.168.224.137.";
    const wchar_t* wSuccess = L"Success";
    //Parse the input
    LPWSTR wInput = utf8_to_wchar(input);

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(wInput, &argc);

    if (argc != 2) {
        //const wchar_t* wError = L"Not enough arguements\n";
        retValue = wchar_to_utf8(wHelp);
        return retValue;
    }
    //Now we got our parameters
    LPWSTR wHostname = argv[0];
    LPWSTR wClsid = argv[1];

    //https://social.msdn.microsoft.com/Forums/vstudio/en-US/181ef8cb-a3e0-47cf-a0d4-77496617226d/how-to-write-a-dcom-project-using-vc?forum=vcgeneral

    HRESULT hr;
    COSERVERINFO csi = { 0 };
    MULTI_QI qi[1] = { 0 };
    CLSID CLSID_Wordpad;


    csi.pwszName = wHostname;
    csi.pAuthInfo = NULL;
    qi[0].pIID = &IID_IUnknown;
    const wchar_t* clsid_str = wClsid;

    CLSIDFromString(clsid_str, &CLSID_Wordpad);
    CoInitializeEx(NULL, NULL);
    hr = CoCreateInstanceEx(CLSID_Wordpad, NULL, CLSCTX_REMOTE_SERVER, &csi, 1, qi);

    std::string message = std::system_category().message(hr);
    LPSTR error = const_cast<char*>(message.c_str());

    if (hr == S_OK) {
        retValue = wchar_to_utf8(wSuccess);
    }
    else {
        retValue = error;
    }

    free(wInput);
    return retValue;
}

extern "C" __declspec(dllexport) VOID FromRundll(HWND hwnd, HINSTANCE hinst, LPSTR input, int nCmdShow) {
    printf(StartInstance(input));
    return;
}