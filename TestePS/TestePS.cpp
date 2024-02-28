// TestePS.cpp : Este arquivo contém a função 'main'. A execução do programa começa e termina ali.
//

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

int main() {
    HRESULT hres;

    // Initialize COM
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        std::cerr << "Failed to initialize COM library. Error code: " << hres << std::endl;
        return 1;
    }

    // Initialize security (optional, but recommended)
    hres = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE,
        NULL
    );
    if (FAILED(hres)) {
        std::cerr << "Failed to initialize security. Error code: " << hres << std::endl;
        CoUninitialize();
        return 1;
    }

    // Obtain the initial locator to WMI
    IWbemLocator* pLoc = nullptr;
    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
        reinterpret_cast<LPVOID*>(&pLoc)
    );
    if (FAILED(hres)) {
        std::cerr << "Failed to create IWbemLocator object. Error code: " << hres << std::endl;
        CoUninitialize();
        return 1;
    }

    // Connect to the desired namespace
    IWbemServices* pSvc = nullptr;
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),  // Namespace
        NULL,                      // User name
        NULL,                      // User password
        0,                         // Locale
        NULL,                      // Security flags
        0,                         // Authority
        0,                         // Context object
        &pSvc                      // IWbemServices proxy
    );
    if (FAILED(hres)) {
        std::cerr << "Could not connect to WMI namespace. Error code: " << hres << std::endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Set security levels on the proxy
    hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE
    );
    if (FAILED(hres)) {
        std::cerr << "Failed to set proxy blanket. Error code: " << hres << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Use the IWbemServices pointer to make requests of WMI
    IEnumWbemClassObject* pEnumerator = nullptr;
    hres = pSvc->ExecQuery(
        _bstr_t(L"WQL"),
        _bstr_t(L"SELECT UUID FROM Win32_ComputerSystemProduct"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator
    );
    if (FAILED(hres)) {
        std::cerr << "Query for computer system product data failed. Error code: " << hres << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Iterate over the retrieved objects
    IWbemClassObject* pclsObj = nullptr;
    ULONG uReturn = 0;
    while (pEnumerator) {
        hres = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn)
            break;

        VARIANT vtProp = {0};
        // Get the value of the "UUID" property
        hres = pclsObj->Get(L"UUID", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hres))
            std::wcout << L"UUID: " << vtProp.bstrVal << std::endl;

        VariantClear(&vtProp);
        pclsObj->Release();
    }

    // Cleanup
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();
    CoUninitialize();

    return 0;
}
