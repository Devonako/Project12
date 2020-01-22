//Sql.cpp
/*
#include  <windows.h>
#include  <string>
#include  <cstdio>
#include  <tchar.h>
#include  <odbcinst.h>
#include  <sqlext.h>
#include  "CSql.h"

CSql::CSql()        // CSql Constructor
{
    szCnOut[0] = 0;
    szErrMsg[0] = 0;
    szErrCode[0] = 0;
    this->hConn = NULL;
    this->hEnvr = NULL;
    this->blnConnected = 0;
}


CSql::~CSql()       // CSql Destructor
{
    if (this->hConn)
    {
        if (this->blnConnected)
            SQLDisconnect(this->hConn);
        SQLFreeHandle(SQL_HANDLE_DBC, this->hConn);
    }
    if (this->hEnvr)
        SQLFreeHandle(SQL_HANDLE_ENV, this->hEnvr);
}


void CSql::MakeConnectionString(void)
{
    if (this->strConnectionString == _T(""))
    {
        if (strDriver == _T("SQL Server"))
        {
            if (strDBQ == (TCHAR*)_T(""))
            {
                strConnectionString = (TCHAR*)_T("DRIVER=");
                strConnectionString = strConnectionString + strDriver + (TCHAR*)_T(";") + (TCHAR*)_T("SERVER=") + strServer + (TCHAR*)_T(";");
                if (strDatabase != _T(""))
                    strConnectionString = strConnectionString + (TCHAR*)_T("DATABASE=") + strDatabase + (TCHAR*)_T(";");
            }
            else
            {
                strConnectionString = (TCHAR*)_T("DRIVER=");
                strConnectionString = strConnectionString + strDriver + (TCHAR*)_T(";") + (TCHAR*)_T("SERVER=") + strServer + (TCHAR*)_T(";") + \
                    (TCHAR*)_T("DATABASE=") + strDatabase + (TCHAR*)_T(";") + (TCHAR*)_T("DBQ=") + strDBQ + (TCHAR*)_T(";");
            }
        }
        else if (strDriver == (TCHAR*)_T("Microsoft Access Driver (*.mdb)"))
        {
            strConnectionString = (TCHAR*)_T("DRIVER=");
            strConnectionString = strConnectionString + strDriver + (TCHAR*)_T(";") + (TCHAR*)_T("DBQ=") + strDBQ + (TCHAR*)_T(";");
        }
        else if (strDriver == (TCHAR*)_T("Microsoft Access Driver (*.mdb, *.accdb)"))
        {
            strConnectionString = _T("DRIVER=");
            strConnectionString = strConnectionString + strDriver + _T(";") + _T("DBQ=") + strDBQ + _T(";");
        }
        else if (strDriver == (TCHAR*)_T("Microsoft Excel Driver (*.xls)"))
        {
            strConnectionString = (TCHAR*)_T("DRIVER=");
            strConnectionString = strConnectionString + strDriver + (TCHAR*)_T(";") + (TCHAR*)_T("DBQ=") + strDBQ + (TCHAR*)_T(";");
        }
    }
}


void CSql::ODBCConnect(void)
{
    TCHAR szCnIn[512];
    UINT iResult;

    MakeConnectionString();
    SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &hEnvr);
    SQLSetEnvAttr(hEnvr, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, SQL_IS_INTEGER);
    SQLAllocHandle(SQL_HANDLE_DBC, hEnvr, &hConn);
    _tcscpy(szCnIn, strConnectionString.c_str());
    iResult = SQLDriverConnect(hConn, NULL, (SQLTCHAR*)szCnIn, (SQLSMALLINT)_tcslen(szCnIn), (SQLTCHAR*)szCnOut, 512, &swStrLen, SQL_DRIVER_NOPROMPT);
    if (iResult == SQL_SUCCESS || iResult == SQL_SUCCESS_WITH_INFO)
    {
        blnConnected = TRUE;
        this->strConnectionString = szCnOut;
    }
    else
    {
        SQLGetDiagRec(SQL_HANDLE_DBC, hConn, 1, szErrCode, &iNativeErrPtr, szErrMsg, 512, &iTextLenPtr);
        blnConnected = FALSE;
        SQLFreeHandle(SQL_HANDLE_DBC, this->hConn), this->hConn = NULL;
        SQLFreeHandle(SQL_HANDLE_ENV, this->hEnvr), this->hEnvr = NULL;
    }
}


void CSql::ODBCDisconnect(void)
{
    if (blnConnected == TRUE)
    {
        SQLDisconnect(hConn);
        SQLFreeHandle(SQL_HANDLE_DBC, hConn), hConn = NULL;
        SQLFreeHandle(SQL_HANDLE_ENV, hEnvr), hEnvr = NULL;
        this->blnConnected = FALSE;
    }
    this->strConnectionString = _T("");
}*/