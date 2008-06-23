#include <windows.h>
#include "sqlite3.h"
//#include <wtypes.h>
#include <oleauto.h> 


#ifdef EXPORT
#undef EXPORT
#endif
#define EXPORT __declspec(dllexport)

/*
**----------------------------------------------------- 
**関数のプロトタイプ宣言
**-----------------------------------------------------
*/

void sqlite3RegisterUserFunctions(sqlite3 *db);
const char* AnsiToUTF8(const char* str_sourse);
const char* UTF8ToAnsi(const char* str_sourse);

/*----------------------------------------------------*/

/*DLLのエントリーポイント*/
BOOL WINAPI DllMain(
  HINSTANCE hinstDLL,  // DLL モジュールのハンドル
  DWORD fdwReason,     // 関数を呼び出す理由
  LPVOID lpvReserved   // 予約済み
)
{
    return TRUE;
}

/*動くかテスト用*/
EXPORT sqlite3* WINAPI sqlite_test(void)
{
	int ret;
	sqlite3* db;
	ret = sqlite3_open("c:\\temp\\abc.db",&db);
	if (ret != SQLITE_OK)
	{
		sqlite3_close(db);
	}
	ret=sqlite3_exec(db,"create table test (id,name);",NULL,NULL,NULL);
	ret = sqlite3_close(db);
	return db;
}

/*sqlite3_open()が良く分からない*/
EXPORT sqlite3* _stdcall vb_sqlite3_open2(const char *zFilename)
{
  sqlite3* db;
  char* sFilename;
  int ret;
  sFilename=AnsiToUTF8(zFilename);
  ret = sqlite3_open(sFilename, &db);
  free(sFilename);
  if (ret != SQLITE_OK)
    {
      sqlite3_close(db);
      return 0;
    }
  return db;
}

/*sqlite3のバージョンを返す*/
EXPORT const char* __stdcall vb_sqlite3_libversion(void)
{
    return sqlite3_libversion();
}

/*DBオープン*/
int __stdcall vb_sqlite3_open(const char *zFilename, sqlite3 **ppDb, int flags)
{
  int ret;
  char* sFilename;
  
  /*一度UTF8にエンコードする*/
  sFilename = AnsiToUTF8(zFilename);
 
  /*DBを開く*/
  if (flags == 0){
      /*従来の開き方*/
      ret = sqlite3_open(sFilename, ppDb);	
  }else{
      /*新しいパターン。開き方を設定できる（読み取り専用など）　*/
      ret = sqlite3_open_v2(sFilename,ppDb,flags, 0);
  }

  /*メモリ開放*/
  free(sFilename);

  /*ユーザ定義関数*/
  sqlite3RegisterUserFunctions(*ppDb);
 
  return ret;
}


/*DBを閉じる*/
int __stdcall vb_sqlite3_close(sqlite3 *db)
{
  int ret;
  ret = sqlite3_close(db);
  //ZeroMemory(db,sizeof(db));
  return ret;
}

/*DMLステートメント実行用*/
int __stdcall vb_sqlite3_exec(sqlite3 *db, const char *zSql, BSTR* ErrMsg)
{
  int ret;
  char* szsql;
  char* ErrMessage = 0;
  char* szErrMsg;
  
  szsql = AnsiToUTF8(zSql);
  ret = sqlite3_exec(db, szsql, NULL, NULL, &ErrMessage);
  free(szsql);
  
  /*エラーが発生するとErrMessageにエラーが書き込まれる*/
  /*sqlite3_execを呼び出した側が開放しないといけない*/
  if (ret != SQLITE_OK) {
   
    /*UTF8から変換する*/ 
    szErrMsg = UTF8ToAnsi(ErrMessage);
    
    /*メモリ開放*/
    sqlite3_free(ErrMessage);
    
    /*vbで確保された変数へコピー*/
    *ErrMsg = SysAllocStringByteLen(szErrMsg,strlen(szErrMsg));
    
    /*メモリ開放*/
    free(szErrMsg);

    return ret;
    
  };
  
  return ret;
}

/* sqlite3_mprintf() のメモリ開放用 */
void __stdcall vb_sqlite3_free(void* z)
{
 sqlite3_free(z);
} 

/*SQL文をコンパイル*/
int __stdcall vb_sqlite3_prepare(sqlite3 *db, const char *zSql, sqlite3_stmt **ppStmt, const char **pzTail)
{
  int ret;
  char* szSql;
  szSql = AnsiToUTF8(zSql);
  
  //return sqlite3_prepare(db, zSql,-1, ppStmt, pzTail);
  ret = sqlite3_prepare_v2(db, szSql,-1, ppStmt, pzTail);

  free(szSql);

  return ret;
}

/*SQL文をコンパイルして構築したものを実行する*/
int __stdcall vb_sqlite3_step(sqlite3_stmt *pStmt)
{
  return sqlite3_step(pStmt);
}

/*コンパイルしたSQLを削除する*/
int __stdcall vb_sqlite3_finalize(sqlite3_stmt *pStmt)
{
  return sqlite3_finalize(pStmt);
}

/*フィールド数を返す*/
int __stdcall vb_sqlite3_column_count(sqlite3_stmt *pStmt)
{
  return sqlite3_column_count(pStmt);
}

/*フィールドの値を取得*/
const unsigned char* __stdcall vb_sqlite3_column_text (sqlite3_stmt *pStmt, int iCol)
{
  char* szretvalue;
  char* szcoltxt;
  BSTR  bstrcoltxt;
  
  szretvalue = sqlite3_column_text(pStmt, iCol);
  
  szcoltxt = UTF8ToAnsi(szretvalue);

  bstrcoltxt = SysAllocStringByteLen(szcoltxt, strlen(szcoltxt));

  free(szcoltxt);
  return bstrcoltxt;
  
}

/*フィールドのタイプを取得*/
int __stdcall vb_sqlite3_column_type(sqlite3_stmt *pStmt, int iCol)
{
  return sqlite3_column_type(pStmt, iCol);
}


/*フィールド名を取得*/
BSTR __stdcall vb_sqlite3_column_name(sqlite3_stmt *pStmt ,int iCol)
{
  
  char* pszretvalue;
  char* pszcolname;
  BSTR  bstrcolname;

  pszretvalue = sqlite3_column_name(pStmt, iCol);

  pszcolname = UTF8ToAnsi(pszretvalue);
    
  bstrcolname = SysAllocStringByteLen(pszcolname,strlen(pszcolname));

  free(pszcolname);

  return bstrcolname;/*VBA側で開放される？*/

}


/*コンパイルしたSQL文にバインドするためのもの*/
int __stdcall vb_sqlite3_bind_double(sqlite3_stmt *pStmt, int iCol, double dValue )
{
  return sqlite3_bind_double(pStmt, iCol, dValue);
}

int __stdcall vb_sqlite3_bind_int(sqlite3_stmt *pStmt, int iCol, int iValue)
{
  return sqlite3_bind_int(pStmt, iCol, iValue);
}

//int sqlite3_bind_null(sqlite3_stmt*, int);

/*コンパイルした文にバインドするためのもの*/
int __stdcall vb_sqlite3_bind_text(sqlite3_stmt *pStmt, int iCol, const char *sValue )
{
  int ret;
  char* szValue;
  szValue = AnsiToUTF8(sValue);
  ret = sqlite3_bind_text(pStmt, iCol, szValue, -1, SQLITE_TRANSIENT);
  free(szValue);
  return ret;
}

/*コンパイルした文にバインドするためのもの*/
int __stdcall vb_sqlite3_reset(sqlite3_stmt *pStmt)
{
  return sqlite3_reset(pStmt);
}

/*SQLの全てのパラメータをNULLに設定*/
int __stdcall vb_sqlite3_clear_bindings(sqlite3_stmt *pStmt)
{
  return sqlite3_clear_bindings(pStmt);
}

/*エラー情報取得*/
BSTR __stdcall vb_sqlite3_errmsg(sqlite3* db)
{
  char* retValue;
  char* szValue;
  BSTR bstrValue;
  
  retValue = sqlite3_errmsg(db);
  szValue = UTF8ToAnsi(retValue); 
  bstrValue = SysAllocStringByteLen(szValue,strlen(szValue));
  free(szValue);
  return bstrValue;
}
/*
-------------------------------------------------------------------------
**プライベート関数
-------------------------------------------------------------------------
*/

/*ANSIからUTF8に変換する関数*/
const char* AnsiToUTF8(const char* str_sourse)
{
  size_t   wcsize;
  size_t   mbsize;
  wchar_t* pszWCstr;
  char*    pszMBstr;

  wcsize   = MultiByteToWideChar(CP_ACP, 0, str_sourse, -1, NULL, 0);
  pszWCstr = malloc(sizeof(wchar_t)*(wcsize + 1));
  wcsize   = MultiByteToWideChar(CP_ACP, 0, str_sourse, -1, pszWCstr, wcsize + 1);
  pszWCstr[wcsize] = '\0'; /*最後に終端文字を追加*/

  mbsize   = WideCharToMultiByte(CP_UTF8, 0, pszWCstr, -1, NULL, 0, NULL, NULL);
  pszMBstr = malloc(mbsize + 1);
  mbsize   = WideCharToMultiByte(CP_UTF8, 0, pszWCstr, -1, pszMBstr, mbsize, NULL, NULL);
  pszMBstr[mbsize] = '\0'; /*最後に終端文字を追加*/

  free(pszWCstr);
  return pszMBstr;

}

/*UTF8からANSIへ変換する関数*/
const char* UTF8ToAnsi(const char* str_sourse) 
{
    size_t   wcsize;
    size_t   mbsize;
    wchar_t* pszWCstr;
    char*    pszMBstr;
    
    wcsize   = MultiByteToWideChar(CP_UTF8, 0, str_sourse, -1, NULL, 0);
    pszWCstr = malloc(sizeof(wchar_t)*(wcsize + 1));
    wcsize   = MultiByteToWideChar(CP_UTF8, 0, str_sourse, -1, pszWCstr, wcsize + 1);
    pszWCstr[wcsize] = '\0'; /*最後に終端文字を追加*/

    mbsize   = WideCharToMultiByte(CP_ACP, 0, pszWCstr, -1, NULL, 0, NULL, NULL);
    pszMBstr = malloc(mbsize + 1);
    mbsize   = WideCharToMultiByte(CP_ACP, 0, pszWCstr, -1, pszMBstr, mbsize, NULL, NULL);
    pszMBstr[mbsize] = '\0'; /*最後に終端文字を追加*/

    free(pszWCstr);
    return pszMBstr;
}

/* user function add */
static void halfFunc(
  sqlite3_context *context,
  int argc,
  sqlite3_value **argv
){
  sqlite3_result_double(context, 0.5*sqlite3_value_double(argv[0]));
}

static void baiFunc(
  sqlite3_context *context,
  int argc,
  sqlite3_value **argv
){
  sqlite3_result_double(context, 2*sqlite3_value_double(argv[0]));
}

void sqlite3RegisterUserFunctions(sqlite3 *db){
  sqlite3_create_function(db, "half", 1, SQLITE_UTF8, 0, halfFunc, 0, 0);
  sqlite3_create_function(db, "bai",  1, SQLITE_UTF8, 0, baiFunc,  0, 0);
}

