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
    CComPtr<IDispatch> p = NULL;
    HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&p);
    if (SUCCEEDED(hr)) {
        return p;
    }
    return NULL;
}

/// <summary>
/// CoInitialize ～ CoUninitialize の管理
/// </summary>
class co_scope
{
public:
    co_scope();
    ~co_scope();
};

co_scope::co_scope()
{
    if (S_OK != CoInitialize(NULL))
    {
    }
}

co_scope::~co_scope()
{
    CoUninitialize();
}

/// <summary>
/// Internet Explorer を開きます。
/// </summary>
/// <param name="url"></param>
void LaunchInternetExplorer(LPCTSTR url)
{
    co_scope co;

    // Internet Explorer の COM 参照
    CComPtr<IDispatch> p = CreateObjectPtr(_T("InternetExplorer.Application"));
    CComQIPtr<IWebBrowser2> browser = (CComQIPtr<IWebBrowser2>)p;
    if (browser == NULL)
    {
        return;
    }

    if (url && url[0])
    {
        browser->Navigate(CComBSTR(url), NULL, NULL, NULL, NULL);
    }
    else
    {
        // pBrowser->Navigate(CComBSTR(_T("about:blank")), NULL, NULL, NULL, NULL);
    }

    // Wait for the browser to finish loading
    // Sleep(5000);

    // Get the HWND of the Internet Explorer window
    HWND hwnd;
    browser->get_HWND((SHANDLE_PTR*)&hwnd);

    // Bring the window to the front
    SetForegroundWindow(hwnd);

    browser->put_Visible(VARIANT_TRUE);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Internet Explorer を開きます。
    LaunchInternetExplorer(lpCmdLine);

    return 0;
}
