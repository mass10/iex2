// iex2.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "framework.h"
#include "iex2.h"

CLSID GetClsID(LPCTSTR name)
{
    CLSID clsid;
    if (S_OK != CLSIDFromProgID(name, &clsid))
        return clsid;
    return clsid;
}

CComPtr<IDispatch> CreateObjectPtr(LPCTSTR name)
{
    CLSID clsid = GetClsID(name);
    CComPtr<IDispatch> pDisp = NULL;
    HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pDisp);
    if (SUCCEEDED(hr)) {
        return pDisp;
    }
    return NULL;
}

class CoScope
{
public:
    CoScope();
    ~CoScope();

};

CoScope::CoScope()
{
    if (S_OK != CoInitialize(NULL))
    {
    }
}

CoScope::~CoScope()
{
    CoUninitialize();
}

void LaunchInternetExplorer(LPCTSTR url)
{
    CoScope __scope__;

    // Internet Explorer の COM 参照
    CComPtr<IDispatch> pDisp = CreateObjectPtr(_T("InternetExplorer.Application"));
    CComQIPtr<IWebBrowser2> pBrowser = (CComQIPtr<IWebBrowser2>)pDisp;
    if (pBrowser == NULL)
    {
        return;
    }

    pBrowser->put_Visible(VARIANT_TRUE);
    if (url && url[0])
    {
        pBrowser->Navigate(CComBSTR(url), NULL, NULL, NULL, NULL);
    }
    else
    {
        // pBrowser->Navigate(CComBSTR(_T("about:blank")), NULL, NULL, NULL, NULL);
    }

    // Wait for the browser to finish loading
    // Sleep(5000);

    // Get the HWND of the Internet Explorer window
    HWND hwnd;
    pBrowser->get_HWND((SHANDLE_PTR*)&hwnd);

    // Bring the window to the front
    SetForegroundWindow(hwnd);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    LaunchInternetExplorer(_T("about:blank"));

    return 0;
}
