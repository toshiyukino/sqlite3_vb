  #include <windows.h>
  #include "sqlite3.h"

int __stdcall vb_sqlite3_open(const char *zFilename, sqlite3 **ppDb)
{
  return sqlite3_open(zFilename, ppDb);
}

int __stdcall vb_sqlite3_close(sqlite3 *db)
{
  return sqlite3_close(db);
}

int __stdcall vb_sqlite3_exec(sqlite3 *db, const char *zSql, const char **ErrMsg)
{
	int ret;
 	char *ErrMessage=0;
	ret = sqlite3_exec(db, zSql, NULL, NULL, ErrMsg);
	return ret;
}

int __stdcall vb_sqlite3_prepare(
  sqlite3 *db,            /* Database handle */
  const char *zSql,       /* SQL statement, UTF-8 encoded */
  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
  const char **pzTail     /* OUT: Pointer to unused portion of zSql */
)
{
  //return sqlite3_prepare(db, zSql,-1, ppStmt, pzTail);
  return sqlite3_prepare_v2(db, zSql,-1, ppStmt, pzTail);
}

int __stdcall vb_sqlite3_step(sqlite3_stmt *pStmt)
{
  return sqlite3_step(pStmt);
}

int __stdcall vb_sqlite3_finalize(sqlite3_stmt *pStmt)
{
  return sqlite3_finalize(pStmt);
}

int __stdcall vb_sqlite3_column_count(sqlite3_stmt *pStmt)
{
  return sqlite3_column_count(pStmt);
}

const unsigned char __stdcall *vb_sqlite3_column_text (sqlite3_stmt *pStmt, int iCol)
{
  return sqlite3_column_text(pStmt, iCol);
}

int __stdcall vb_sqlite3_column_type(sqlite3_stmt *pStmt, int iCol)
{
  return sqlite3_column_type(pStmt, iCol);
}


const char __stdcall *vb_sqlite3_column_name(sqlite3_stmt *pStmt ,int iCol)
{
  return sqlite3_column_name(pStmt, iCol);
}


int __stdcall vb_sqlite3_bind_double(sqlite3_stmt *pStmt, int iCol, double dValue )
{
  return sqlite3_bind_double(pStmt, iCol, dValue);
}

int __stdcall vb_sqlite3_bind_int(sqlite3_stmt *pStmt, int iCol, int iValue)
{
  return sqlite3_bind_int(pStmt, iCol, iValue);
}

//int sqlite3_bind_null(sqlite3_stmt*, int);

int __stdcall vb_sqlite3_bind_text(sqlite3_stmt *pStmt, int iCol, const char *sValue )
{
  return sqlite3_bind_text(pStmt, iCol, sValue, -1, SQLITE_TRANSIENT);
}

int __stdcall vb_sqlite3_reset(sqlite3_stmt *pStmt)
{
  return sqlite3_reset(pStmt);
}

int __stdcall vb_sqlite3_clear_bindings(sqlite3_stmt *pStmt)
{
  return sqlite3_clear_bindings(pStmt);
}

//
///* user function add */
//static void halfFunc(
//  sqlite3_context *context,
//  int argc,
//  sqlite3_value **argv
//){
//  sqlite3_result_double(context, 0.5*sqlite3_value_double(argv[0]));
//}

//static void baiFunc(
//  sqlite3_context *context,
//  int argc,
//  sqlite3_value **argv
//){
//  sqlite3_result_double(context, 2*sqlite3_value_double(argv[0]));
//}

//void sqlite3RegisterUserFunctions(sqlite3 *db){
//  sqlite3_create_function(db, "half", 1, SQLITE_UTF8, 0, halfFunc, 0, 0);
//  sqlite3_create_function(db, "bai",  1, SQLITE_UTF8, 0, baiFunc,  0, 0);
//}

