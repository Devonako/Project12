#ifndef   UNICODE         
#define   UNICODE     
#endif                    
#ifndef   _UNICODE        
#define   _UNICODE    
#endif                    
#include <windows.h>      
#ifdef TCLib              
#include "stdio.h"     
#else                     
#include <stdio.h>     
#endif
#include "XLFunctions.h"

int main()
{
    IDispatch* pXLApp = NULL;
    IDispatch* pXLWorkBooks = NULL;
    IDispatch* pXLWorkBook = NULL;
    IDispatch* pXLWorkSheet = NULL;
    wchar_t      szWorkBookPath[] = L"C:\\Users\\dakoz\\Downloads\\xls\\Book1.xls";
    wchar_t      szSheet[] = L"Sheet2";
    wchar_t      szRange[] = L"A1";
    wchar_t      szCell[64];

    CoInitialize(NULL);                                           // Start COM Subsystem
    pXLApp = XLStart(true, &pXLWorkBooks);                   // Start Excel
    pXLWorkBook = OpenXLWorkBook(pXLWorkBooks, szWorkBookPath);   // Open *.xls WorkBook    
    pXLWorkSheet = XLOpenWorkSheet(pXLWorkBook, szSheet);          // Select/Open Specific WorkSheet
    GetXLCell(pXLWorkSheet, szRange, szCell, 64);                    // Retrieve Cell Data
    wprintf(L"szCell = %s\n", szCell);          
    // Output Cell Contents To Console
    lstrcpyW(szRange, L"B1");
    GetXLCell(pXLWorkSheet, szRange, szCell, 64);
    wprintf(L"szCell = %s\n", szCell);
    pXLWorkSheet->Release();                                      // Release Pointer To Work Sheet
    pXLWorkBook->Release();                                       // Release Pointer To Work Book
    pXLWorkBooks->Release();                                      // Release Pointer To Work Books Collection
    getchar();                                                    // Hold Console & Excel Open
    XLQuit(pXLApp);                                               // Close Excel Application
    pXLApp->Release();                                            // Release Pointer To Excel App
    CoUninitialize();                                             // Shut Down COM Subsystem

    return 0;
}