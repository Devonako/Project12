#include <windows.h>
#include <tchar.h>
#include <cstdio>
const CLSID CLSID_XLApplication = { 0x00024500,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46} }; // CLSID of Excel
const IID   IID_Application = { 0x000208D5,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46} }; // IID of _Application

int main()
{
    DISPPARAMS NoArgs = { NULL,NULL,0,0 };
    IDispatch* pXLApp = NULL;
    DISPPARAMS DispParams;
    VARIANT CallArgs[1];
    void* pMsgBuf = NULL;
    VARIANT vResult;
    DWORD dwFlags;
    DISPID dispid;
    HRESULT hr;
    LCID lcid;

    CoInitialize(NULL);
    dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM;
    hr = CoCreateInstance(CLSID_XLApplication, NULL, CLSCTX_LOCAL_SERVER, IID_Application, (void**)&pXLApp);
    if (SUCCEEDED(hr))
    {
        OLECHAR* szVisible = (OLECHAR*)L"Visible";
        lcid = GetUserDefaultLCID();
        hr = pXLApp->GetIDsOfNames(IID_NULL, &szVisible, 1, lcid, &dispid);
        if (SUCCEEDED(hr))
        {
            VariantInit(&CallArgs[0]);
            CallArgs[0].vt = VT_BOOL;
            CallArgs[0].boolVal = TRUE;
            DISPID dispidNamed = DISPID_PROPERTYPUT;
            DispParams.rgvarg = CallArgs;
            DispParams.rgdispidNamedArgs = &dispidNamed;
            DispParams.cArgs = 1;
            DispParams.cNamedArgs = 1;
            VariantInit(&vResult);                   //  Call or Invoke _Application::Visible(true);
            hr = pXLApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &DispParams, &vResult, NULL, NULL);
            if (SUCCEEDED(hr))
            {
                OLECHAR* szWorkbooks = (OLECHAR*)L"Workbooks";
                hr = pXLApp->GetIDsOfNames(IID_NULL, &szWorkbooks, 1, GetUserDefaultLCID(), &dispid);
                if (SUCCEEDED(hr))
                {
                    IDispatch* pXLBooks = NULL;          //  Get Workbooks Collection
                    VariantInit(&vResult);             //  Invoke _Application::Workbooks(&pXLBooks) << returns IDispatch** of Workbooks Collection
                    hr = pXLApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &NoArgs, &vResult, NULL, NULL);
                    if (SUCCEEDED(hr))
                    {
                        pXLBooks = vResult.pdispVal;
                        IDispatch* pXLBook = NULL;        //  Try to Open Workbook
                        OLECHAR* szOpen = (OLECHAR*)L"Open";
                        hr = pXLBooks->GetIDsOfNames(IID_NULL, &szOpen, 1, GetUserDefaultLCID(), &dispid);
                        if (SUCCEEDED(hr))
                        {
                            VariantInit(&vResult);       //  Invoke Workbooks::Open(&Workbook)  << returns IDispatch** of Workbook Object
                            CallArgs[0].vt = VT_BSTR;
                            CallArgs[0].bstrVal = SysAllocString(L"C:\\Users\\dakoz\\Downloads\\xls\\Book1.xls");
                            DispParams.rgvarg = CallArgs;
                            DispParams.rgdispidNamedArgs = &dispidNamed;
                            DispParams.cArgs = 1;
                            DispParams.cNamedArgs = 0;  // Invoke _Worksheet::Range("A1")  << returns IDispatch** to dispinterface Range
                            hr = pXLBooks->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD | DISPATCH_METHOD, &DispParams, &vResult, NULL, NULL);
                            // [id(0x000000b5), helpcontext(0x000100b5)] HRESULT Add([in, lcid] long lcid, [out, retval] Workbook** RHS);
                            // [id(0x000002aa), helpcontext(0x000102aa)] HRESULT Open([in] BSTR Filename, [in, lcid] long lcid, [out, retval] Workbook** RHS);
                            if (SUCCEEDED(hr))
                            {
                                printf("We Succeeded!\n");
                                pXLBook = vResult.pdispVal;
                                OLECHAR* szActiveSheet = (OLECHAR*)L"ActiveSheet";
                                hr = pXLApp->GetIDsOfNames(IID_NULL, &szActiveSheet, 1, GetUserDefaultLCID(), &dispid);
                                if (SUCCEEDED(hr))
                                {
                                    IDispatch* pXLSheet = NULL;  // Try To Get ActiveSheet
                                    VariantInit(&vResult);     // Invoke _Application::ActiveSheet(&pXLSheet);  << ret IDispatch** to Worksheet (Worksheet)
                                    hr = pXLApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &NoArgs, &vResult, NULL, NULL);
                                    if (SUCCEEDED(hr))
                                    {
                                        pXLSheet = vResult.pdispVal;
                                        OLECHAR* szRange = (OLECHAR*)L"Range";
                                        hr = pXLSheet->GetIDsOfNames(IID_NULL, &szRange, 1, GetUserDefaultLCID(), &dispid);
                                        if (SUCCEEDED(hr))
                                        {
                                            IDispatch* pXLRange = NULL;
                                            VariantInit(&vResult);
                                            CallArgs[0].vt = VT_BSTR,
                                                CallArgs[0].bstrVal = SysAllocString(L"A1");
                                            DispParams.rgvarg = CallArgs;
                                            DispParams.rgdispidNamedArgs = 0;
                                            DispParams.cArgs = 1;  // Try to get Range
                                            DispParams.cNamedArgs = 0;  // Invoke _Worksheet::Range("A1")  << returns IDispatch** to dispinterface Range
                                            hr = pXLSheet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &DispParams, &vResult, NULL, NULL);
                                            if (SUCCEEDED(hr))
                                            {
                                                pXLRange = vResult.pdispVal;
                                                OLECHAR* szValue = (OLECHAR*)L"Value";
                                                hr = pXLRange->GetIDsOfNames(IID_NULL, &szValue, 1, GetUserDefaultLCID(), &dispid);
                                                if (SUCCEEDED(hr))
                                                {
                                                    printf("dispid (Value) = %d\n", (int)dispid);
                                                    VariantClear(&CallArgs[0]);
                                                    hr = pXLRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &NoArgs, &vResult, NULL, NULL);
                                                    if (SUCCEEDED(hr))
                                                    {
                                                        wprintf(L"vResult.bstrVal = %s\n", vResult.bstrVal);
                                                        VariantClear(&vResult);
                                                    }
                                                    else
                                                    {
                                                        FormatMessage(dwFlags, NULL, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&pMsgBuf, 0, NULL);
                                                        printf("%s\n", (char*)pMsgBuf);
                                                        LocalFree(pMsgBuf);
                                                    }
                                                    pXLRange->Release();
                                                }
                                            }
                                        }
                                        pXLSheet->Release();
                                    }
                                }
                                pXLBook->Release();
                            }
                            else
                            {
                                FormatMessage(dwFlags, NULL, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&pMsgBuf, 0, NULL);
                                printf("%s\n", (char*)pMsgBuf);
                                LocalFree(pMsgBuf);
                            }
                        }
                        pXLBooks->Release();
                    }
                }
            }
            getchar();
        }
        VariantInit(&vResult);  // Try to do _Application::Close()
        hr = pXLApp->Invoke(0x0000012e, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &NoArgs, &vResult, NULL, NULL);
        pXLApp->Release();
    }
    CoUninitialize();

    return 0;
}