#include <windows.h>
#include "sqlite3.h"
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
char* AnsiToUTF8(const char* str_sourse);
char* UTF8ToAnsi(const char* str_sourse);

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



/*sqlite3のバージョンを返す*/
EXPORT BSTR __stdcall vb_sqlite3_libversion(void)
{
  const char* ret;
  ret = sqlite3_libversion();
  return SysAllocStringByteLen(ret, strlen(ret));
}

/*DBオープン*/
//EXPORT int __stdcall vb_sqlite3_open(const char *zFilename, int* handle , int flags)
EXPORT int __stdcall vb_sqlite3_open(const char *zFilename, sqlite3** db , int flags)
{
  int ret;
  char* zMBFilename;
  //sqlite3 *db;
  
  /*一度UTF8にエンコードする*/
  zMBFilename = AnsiToUTF8(zFilename);

  /*DBを開く*/
  if (flags == 0){
      /*従来の開き方*/
      ret = sqlite3_open(zMBFilename, db);
      if (ret != SQLITE_OK){
        sqlite3_close(*db);
        ret = -1;
      }
  }else{
      /*新しいパターン。開き方を設定できる（読み取り専用など）　*/
      ret = sqlite3_open_v2(zMBFilename, db, flags, 0);
      if (ret != SQLITE_OK){
        sqlite3_close(*db);
        ret = -1;
      }
  }

  /*メモリ開放*/
  free(zMBFilename);

  /*ユーザ定義関数*/
  sqlite3RegisterUserFunctions(*db);
 
  //*handle = (int)db;
  return ret;
}


/*DBを閉じる*/
EXPORT int __stdcall vb_sqlite3_close(sqlite3* db)
{
  int ret;
  ret = sqlite3_close(db);
  return ret;
}

/*DMLステートメント実行用*/
EXPORT int __stdcall vb_sqlite3_exec(sqlite3 *db, const char *zSql, BSTR* ErrMsg)
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
EXPORT void __stdcall vb_sqlite3_free(void* z)
{
 sqlite3_free(z);
} 

/*SQL文をコンパイル*/
//EXPORT int __stdcall vb_sqlite3_prepare(sqlite3 *db, const char *zSql, sqlite3_stmt **ppStmt, const char **pzTail)
EXPORT int __stdcall vb_sqlite3_prepare(sqlite3 *db, const char *zSql, sqlite3_stmt **ppStmt, BSTR *pzTail)
{
  int ret;
  char* szSql;
  const char* pzTailUTF8;
  char* szTail;

  szSql = AnsiToUTF8(zSql);
  
  //return sqlite3_prepare(db, zSql,-1, ppStmt, pzTail);
  //ret = sqlite3_prepare_v2(db, szSql,-1, ppStmt, pzTail);
  ret = sqlite3_prepare_v2(db, szSql,-1, ppStmt, &pzTailUTF8);

  szTail = UTF8ToAnsi(pzTailUTF8);
  *pzTail = SysAllocStringByteLen(szTail,strlen(szTail));

  free(szSql);
  free(szTail);
  
  return ret;
}

/*SQL文をコンパイルして構築したものを実行する*/
EXPORT int __stdcall vb_sqlite3_step(sqlite3_stmt *pStmt)
{
  return sqlite3_step(pStmt);
}

/*コンパイルしたSQLを削除する*/
EXPORT int __stdcall vb_sqlite3_finalize(sqlite3_stmt *pStmt)
{
  return sqlite3_finalize(pStmt);
}

/*フィールド数を返す*/
EXPORT int __stdcall vb_sqlite3_column_count(sqlite3_stmt *pStmt)
{
  return sqlite3_column_count(pStmt);
}

/*フィールド数を返す２*/
EXPORT int __stdcall vb_sqlite3_data_count(sqlite3_stmt *pStmt)
{
  return sqlite3_data_count(pStmt);
}

/*フィールドの値を取得*/
EXPORT BSTR __stdcall vb_sqlite3_column_text (sqlite3_stmt *pStmt, int iCol)
{
  const unsigned char* szretvalue;
  char* szcoltxt;
  BSTR  bstrcoltxt;
  
  szretvalue = sqlite3_column_text(pStmt, iCol);
  
  szcoltxt = UTF8ToAnsi(szretvalue);

  bstrcoltxt = SysAllocStringByteLen(szcoltxt, strlen(szcoltxt));

  free(szcoltxt);
  return bstrcoltxt;
  
}

/*フィールドのタイプを取得*/
EXPORT int __stdcall vb_sqlite3_column_type(sqlite3_stmt *pStmt, int iCol)
{
  return sqlite3_column_type(pStmt, iCol);
}

/*フィールド名を取得*/
EXPORT BSTR __stdcall vb_sqlite3_column_name(sqlite3_stmt *pStmt ,int iCol)
{
  
  const char* pszretvalue;
  char* pszcolname;
  BSTR  bstrcolname;

  pszretvalue = sqlite3_column_name(pStmt, iCol);

  pszcolname = UTF8ToAnsi(pszretvalue);
    
  bstrcolname = SysAllocStringByteLen(pszcolname,strlen(pszcolname));

  free(pszcolname);

  return bstrcolname;/*VBA側で開放される？*/

}

/*コンパイルしたSQL文にバインドするためのもの*/
EXPORT int __stdcall vb_sqlite3_bind_double(sqlite3_stmt *pStmt, int iCol, double dValue )
{
  return sqlite3_bind_double(pStmt, iCol, dValue);
}

EXPORT int __stdcall vb_sqlite3_bind_int(sqlite3_stmt *pStmt, int iCol, int iValue)
{
  return sqlite3_bind_int(pStmt, iCol, iValue);
}

//int sqlite3_bind_null(sqlite3_stmt*, int);

/*コンパイルした文にバインドするためのもの*/
EXPORT int __stdcall vb_sqlite3_bind_text(sqlite3_stmt *pStmt, int iCol, const char *sValue )
{
  int ret;
  char* szValue;
  szValue = AnsiToUTF8(sValue);
  ret = sqlite3_bind_text(pStmt, iCol, szValue, -1, SQLITE_TRANSIENT);
  //ret = sqlite3_bind_text(pStmt, iCol, szValue, -1, SQLITE_STATIC);
  
  free(szValue);
  return ret;
}

/*コンパイルした文にバインドするためのもの*/
EXPORT int __stdcall vb_sqlite3_reset(sqlite3_stmt *pStmt)
{
  return sqlite3_reset(pStmt);
}

/*SQLの全てのパラメータをNULLに設定*/
EXPORT int __stdcall vb_sqlite3_clear_bindings(sqlite3_stmt *pStmt)
{
  return sqlite3_clear_bindings(pStmt);
}

/*エラー情報取得*/
EXPORT BSTR __stdcall vb_sqlite3_errmsg(sqlite3* db)
{
  const char* retValue;
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
char* AnsiToUTF8(const char* str_sourse)
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
char* UTF8ToAnsi(const char* str_sourse) 
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

/*
**********************************************************************
** SQLite3のソースファイル「os_win.c」から引用する。
** [sqlite3_win32_mbcs_to_utf8]はエクスポートされているのでそれを使う
**********************************************************************
*/
/*
** UTF-8 ⇒ MS-UNICODE
*/
static WCHAR *utf8ToUnicode(const char *zFilename){
  int nChar;
  WCHAR *zWideFilename;

  nChar = MultiByteToWideChar(CP_UTF8, 0, zFilename, -1, NULL, 0);
  zWideFilename = malloc( nChar*sizeof(zWideFilename[0]) );
  if( zWideFilename==0 ){
    return 0;
  }
  nChar = MultiByteToWideChar(CP_UTF8, 0, zFilename, -1, zWideFilename, nChar);
  if( nChar==0 ){
    free(zWideFilename);
    zWideFilename = 0;
  }
  return zWideFilename;
}


/*
** MS-UNICODE ⇒ ANSI
*/
static char *unicodeToMbcs(const WCHAR *zWideFilename){
  int nByte;
  char *zFilename;
  int codepage = AreFileApisANSI() ? CP_ACP : CP_OEMCP;

  nByte = WideCharToMultiByte(codepage, 0, zWideFilename, -1, 0, 0, 0, 0);
  zFilename = malloc( nByte );
  if( zFilename==0 ){
    return 0;
  }
  nByte = WideCharToMultiByte(codepage, 0, zWideFilename, -1, zFilename, nByte,
                              0, 0);
  if( nByte == 0 ){
    free(zFilename);
    zFilename = 0;
  }
  return zFilename;
}

/*
** UTF-8 ⇒ ANSI
*/
static char *utf8ToMbcs(const char *zFilename){
  char *zFilenameMbcs;
  WCHAR *zTmpWide;

  zTmpWide = utf8ToUnicode(zFilename);
  if( zTmpWide==0 ){
    return 0;
  }
  zFilenameMbcs = unicodeToMbcs(zTmpWide);
  free(zTmpWide);
  return zFilenameMbcs;
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

