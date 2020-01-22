/*#include <windows.h>
#include <sqlext.h>
#include "CSql.h"

int main()
{
    std::string strQuery;  // One uses SQL (Structured Query Language) to describe to the database driver the data one wants
    unsigned int iCol[5];  // The addresses of these variables are bound to the resultant database curson generated
    SQLINTEGER iRead[5];   // This parameter in SQLBindCol() is an OUT parameter which will contain the # bytes read into bound variable
    HSTMT hStmt = NULL;    // ODBC Statement HANDLE
    CSql Sql;              // Connection Object/Class

    Sql.strDriver = "Microsoft Excel Driver (*.xls)";  // Identify Database Driver To ODBC
    Sql.strDBQ = "Book1.xls";                        // Identify 'Database', so to speak, such as it is
    Sql.ODBCConnect();                               // Connect using generated connection string
    if (Sql.blnConnected == TRUE)                       //  Everything's gotta be tested for errors!!!
    {
        printf("Sql.blnConnected = TRUE\n");
        if (SQLAllocHandle(SQL_HANDLE_STMT, Sql.hConn, &hStmt) == SQL_SUCCESS)  // Allocate Statement Handle
        {
            printf("SQLAllocHandle() Succeeded!\n");
            strQuery = "SELECT Col1, Col2, Col3, Col4, Col5 FROM [Sheet1$];";
            SQLBindCol(hStmt, 1, SQL_C_ULONG, &iCol[0], 0, &iRead[0]);           // Bind app variables to
            SQLBindCol(hStmt, 2, SQL_C_ULONG, &iCol[1], 0, &iRead[1]);           // retrieved database cursor
            SQLBindCol(hStmt, 3, SQL_C_ULONG, &iCol[2], 0, &iRead[2]);
            SQLBindCol(hStmt, 4, SQL_C_ULONG, &iCol[3], 0, &iRead[3]);
            SQLBindCol(hStmt, 5, SQL_C_ULONG, &iCol[4], 0, &iRead[4]);
            if (SQLExecDirect(hStmt, (SQLTCHAR*)strQuery.c_str(), SQL_NTS) == SQL_SUCCESS)  // execute SQL Statement
            {
                SQLFetch(hStmt);  // Retrieve database cursor into app variables
                printf("%u\t%u\t%u\t%u\t%u\n", iCol[0], iCol[1], iCol[2], iCol[3], iCol[4]);
                printf("%d\t%d\t%d\t%d\t%d\n", iRead[0], iRead[1], iRead[2], iRead[3], iRead[4]);
                SQLCloseCursor(hStmt);   // Close Database Cursor
            }
            SQLFreeHandle(SQL_HANDLE_STMT, hStmt);  // Free Statement Handle
        }
        Sql.ODBCDisconnect();  // Disconnect from database
    }
    else
        printf("Sql.blnConnected == FALSE!\n");
    getchar();

    return 0;
}*/