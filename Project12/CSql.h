//MySql.h
#if !defined(SQL_H)
#define SQL_H
#include <string>
#include <sqltypes.h>

class CSql
{
public:
    CSql();
    ~CSql();
    void MakeConnectionString(void);
    void ODBCConnect(void);
    void ODBCDisconnect(void);

public:
    std::string       strConnectionString;
    std::string       strDatabase;
    std::string       strDriver;
    std::string       strServer;
    std::string       strDBQ;
    char              szCnOut[512];
    short int         iBytes;
    SWORD             swStrLen;
    SQLHENV           hEnvr;
    SQLHDBC           hConn;
    SQLINTEGER        iNativeErrPtr;
    SQLSMALLINT       iTextLenPtr;
    SQLTCHAR          szErrMsg[512];
    SQLTCHAR          szErrCode[8];
    unsigned int      blnConnected;
};

#endif 