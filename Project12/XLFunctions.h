// XLFunctions.h
#ifndef XLFunctions_h
#define XLFunctions_h

HRESULT SetVisible(IDispatch* pObject);
HRESULT GetXLCell(IDispatch* pXLWorksheet, wchar_t* pszRange, wchar_t* pszCell, size_t iBufferLength);
HRESULT GetCell(IDispatch* pXLSheet, wchar_t* pszRange, VARIANT& pVt);
IDispatch* SelectWorkSheet(IDispatch* pXLWorksheets, LCID& lcid, wchar_t* pszSheet);
IDispatch* OpenXLWorkBook(IDispatch* pXLWorkbooks, wchar_t* pszWorkBookPath);
IDispatch* GetDispatchObject(IDispatch* pCallerObject, DISPID dispid, WORD wFlags, LCID lcid);
IDispatch* XLStart(bool blnVisible, IDispatch** pXLWorkBooks);
IDispatch* XLOpenWorkSheet(IDispatch* pXLWorkBook, wchar_t* pSheet);
HRESULT XLQuit(IDispatch* pXLApp);

#endif