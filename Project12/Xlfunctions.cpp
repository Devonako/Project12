#define _CRT_SECURE_NO_WARNINGS
#ifndef   UNICODE
#define   UNICODE
#endif
#ifndef   _UNICODE
#define   _UNICODE
#endif
#include <windows.h>
#include <stdio.h>
#include "XLFunctions.h"


HRESULT SetVisible(IDispatch* pObject)
{
    VARIANT         vArgArray[1];
    DISPPARAMS      DispParams;
    DISPID          dispidNamed;
    VARIANT         vResult;
    HRESULT         hr;
    LCID            lcid;

    VariantInit(&vArgArray[0]);
    vArgArray[0].vt = VT_BOOL;
    vArgArray[0].boolVal = TRUE;
    dispidNamed = DISPID_PROPERTYPUT;
    DispParams.rgvarg = vArgArray;
    DispParams.rgdispidNamedArgs = &dispidNamed;
    DispParams.cArgs = 1;
    DispParams.cNamedArgs = 1;
    lcid = GetUserDefaultLCID();
    VariantInit(&vResult);
    hr = pObject->Invoke(0x0000022e, IID_NULL, lcid, DISPATCH_PROPERTYPUT, &DispParams, &vResult, NULL, NULL);

    return hr;
}


HRESULT GetXLCell(IDispatch* pXLWorksheet, wchar_t* pszRange, wchar_t* pszCell, size_t iBufferLength)
{
    DISPPARAMS      NoArgs = { NULL,NULL,0,0 };
    IDispatch* pXLRange = NULL;
    VARIANT         vArgArray[1];
    VARIANT         vResult;
    DISPPARAMS      DispParams;
    HRESULT         hr;
    LCID            lcid;

    VariantInit(&vResult);
    lcid = GetUserDefaultLCID();
    vArgArray[0].vt = VT_BSTR,
        vArgArray[0].bstrVal = SysAllocString(pszRange);
    DispParams.rgvarg = vArgArray;
    DispParams.rgdispidNamedArgs = 0;
    DispParams.cArgs = 1;
    DispParams.cNamedArgs = 0;
    hr = pXLWorksheet->Invoke
    (
        0xC5,
        IID_NULL,
        lcid,
        DISPATCH_PROPERTYGET,
        &DispParams,
        &vResult,
        NULL,
        NULL
    );
    if (FAILED(hr))
        return E_FAIL;
    pXLRange = vResult.pdispVal;

    //Member Get Value <6> () As Variant         
    VariantClear(&vArgArray[0]);
    hr = pXLRange->Invoke
    (
        6,
        IID_NULL,
        lcid,
        DISPATCH_PROPERTYGET,
        &NoArgs,
        &vResult,
        NULL,
        NULL
    );
    if (SUCCEEDED(hr))
    {
        if (vResult.vt == VT_BSTR)
        {
            if (SysStringLen(vResult.bstrVal) < iBufferLength)
            {
                wcscpy(pszCell, vResult.bstrVal);
                VariantClear(&vResult);
                return S_OK;
            }
            else
            {
                VariantClear(&vResult);
                return E_FAIL;
            }
        }
        else
        {
            pszCell[0] = 0;
            VariantClear(&vResult);
        }
    }
    pXLRange->Release();

    return E_FAIL;
}


HRESULT GetCell(IDispatch* pXLSheet, wchar_t* pszRange, VARIANT& pVt)
{
    DISPPARAMS      NoArgs = { NULL,NULL,0,0 };
    IDispatch* pXLRange = NULL;
    VARIANT         vArgArray[1];
    VARIANT         vResult;
    DISPPARAMS      DispParams;
    HRESULT         hr;
    LCID            lcid;

    VariantInit(&vResult);
    vArgArray[0].vt = VT_BSTR,
        vArgArray[0].bstrVal = SysAllocString(pszRange);
    DispParams.rgvarg = vArgArray;
    DispParams.rgdispidNamedArgs = 0;
    DispParams.cArgs = 1;
    DispParams.cNamedArgs = 0;
    lcid = GetUserDefaultLCID();
    hr = pXLSheet->Invoke(0xC5, IID_NULL, lcid, DISPATCH_PROPERTYGET, &DispParams, &vResult, NULL, NULL);
    if (FAILED(hr))VariantClear(&vResult);
    return hr;
    pXLRange = vResult.pdispVal;

    //Member Get Value <6> () As Variant
    VariantClear(&vArgArray[0]);
    VariantClear(&pVt);
    hr = pXLRange->Invoke(6, IID_NULL, lcid, DISPATCH_PROPERTYGET, &NoArgs, &pVt, NULL, NULL);
    pXLRange->Release();

    return hr;
}


IDispatch* SelectWorkSheet(IDispatch* pXLWorksheets, LCID& lcid, wchar_t* pszSheet)
{
    VARIANT         vResult;
    HRESULT         hr;
    VARIANT         vArgArray[1];
    DISPPARAMS      DispParams;
    DISPID          dispidNamed;
    IDispatch* pXLWorksheet = NULL;

    // Member Get Item <170> (In Index As Variant<0>) As IDispatch  >> Gets pXLWorksheet
    // [id(0x000000aa), propget, helpcontext(0x000100aa)] IDispatch* Item([in] VARIANT Index);
    VariantInit(&vResult);
    vArgArray[0].vt = VT_BSTR;
    vArgArray[0].bstrVal = SysAllocString(pszSheet);
    DispParams.rgvarg = vArgArray;
    DispParams.rgdispidNamedArgs = &dispidNamed;
    DispParams.cArgs = 1;
    DispParams.cNamedArgs = 0;
    hr = pXLWorksheets->Invoke(0xAA, IID_NULL, lcid, DISPATCH_PROPERTYGET, &DispParams, &vResult, NULL, NULL);
    if (FAILED(hr))
        return NULL;
    pXLWorksheet = vResult.pdispVal;
    SysFreeString(vArgArray[0].bstrVal);

    // Worksheet::Select()
    VariantInit(&vResult);
    VARIANT varReplace;
    varReplace.vt = VT_BOOL;
    varReplace.boolVal = VARIANT_TRUE;
    dispidNamed = 0;
    DispParams.rgvarg = &varReplace;
    DispParams.rgdispidNamedArgs = &dispidNamed;
    DispParams.cArgs = 1;
    DispParams.cNamedArgs = 1;
    hr = pXLWorksheet->Invoke(0xEB, IID_NULL, lcid, DISPATCH_METHOD, &DispParams, &vResult, NULL, NULL);

    return pXLWorksheet;
}


IDispatch* OpenXLWorkBook(IDispatch* pXLWorkbooks, wchar_t* pszWorkBookPath)
{
    VARIANT         vResult;
    VARIANT         vArgArray[1];
    DISPPARAMS      DispParams;
    DISPID          dispidNamed;
    LCID            lcid;
    HRESULT         hr;

    VariantInit(&vResult);         // Call Workbooks::Open() - 682  >> Gets pXLWorkbook
    vArgArray[0].vt = VT_BSTR;
    vArgArray[0].bstrVal = SysAllocString(pszWorkBookPath);
    DispParams.rgvarg = vArgArray;
    DispParams.rgdispidNamedArgs = &dispidNamed;
    DispParams.cArgs = 1;
    DispParams.cNamedArgs = 0;
    lcid = GetUserDefaultLCID();
    hr = pXLWorkbooks->Invoke(682, IID_NULL, lcid, DISPATCH_METHOD, &DispParams, &vResult, NULL, NULL);
    SysFreeString(vArgArray[0].bstrVal);
    if (FAILED(hr))
        return NULL;
    else
        return vResult.pdispVal;
}
IDispatch* GetDispatchObject(IDispatch* pCallerObject, DISPID dispid, WORD wFlags, LCID lcid)
{
    DISPPARAMS   NoArgs = { NULL,NULL,0,0 };
    VARIANT      vResult;
    HRESULT      hr;

    VariantInit(&vResult);
    hr = pCallerObject->Invoke(dispid, IID_NULL, lcid, wFlags, &NoArgs, &vResult, NULL, NULL);
    if (FAILED(hr))
        return NULL;
    else
        return vResult.pdispVal;
}


IDispatch* XLStart(bool blnVisible, IDispatch** pXLWorkBooks)
{
    const CLSID  CLSID_XLApplication = { 0x00024500,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46} };
    const IID    IID_Application = { 0x000208D5,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46} };
    IDispatch* pXLApp = NULL;
    LCID         lcid;
    HRESULT      hr;

    hr = CoCreateInstance(CLSID_XLApplication, NULL, CLSCTX_LOCAL_SERVER, IID_Application, (void**)&pXLApp); // Returns in last [out] parameter pointer to app object
    if (SUCCEEDED(hr))  // macro that tests HRESULT, which is a bit field entity, for S_OK, i.e., success
    {
        lcid = GetUserDefaultLCID();
        if (blnVisible)
            SetVisible(pXLApp);
        *pXLWorkBooks = GetDispatchObject(pXLApp, 572, DISPATCH_PROPERTYGET, lcid);  // Wrapper function in XLFunctions.cpp will return IDispatch pointer tp WorkBooks Collection
        if (pXLWorkBooks)
            return pXLApp;
        else
        {
            pXLApp->Release();
            return NULL;
        }
    }

    return NULL;
}


IDispatch* XLOpenWorkSheet(IDispatch* pXLWorkBook, wchar_t* pSheet)
{
    IDispatch* pXLWorkSheets = NULL;
    IDispatch* pXLWorkSheet = NULL;
    LCID         lcid;

    lcid = GetUserDefaultLCID();
    pXLWorkSheets = GetDispatchObject(pXLWorkBook, 494, DISPATCH_PROPERTYGET, lcid);
    pXLWorkSheet = SelectWorkSheet(pXLWorkSheets, lcid, pSheet);
    pXLWorkSheets->Release();

    return pXLWorkSheet;
}


HRESULT XLQuit(IDispatch* pXLApp)
{
    DISPPARAMS   NoArgs = { NULL,NULL,0,0 };
    return pXLApp->Invoke(0x0000012e, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &NoArgs, NULL, NULL, NULL); // pXLApp->Quit() 0x12E
}