/***************************************************************************************************
*										通知の挙動設定
***************************************************************************************************
*              "wpndatabase.db" の操作に特化した SQLite 操作コードを提供します。
* 
***************************************************************************************************/



//設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"
#include "sqlite3.h"   // SQLiteのヘッダー



//***************************************************************************************************
//                              ■■■ VBA側から呼び出す関数 ■■■
//***************************************************************************************************
//* 機能　　：SQLite方式のデータベースの操作を行います
//---------------------------------------------------------------------------------------------------
//* 返り値　：SELECT文：先頭行、先頭列の1列。それ以外は、実行件数
//* 引数　　：dbPath        dbファイルの絶対パス
//            sql           SQL文
//---------------------------------------------------------------------------------------------------
//* 注意事項：・SELECT文を行う際は、列は1つ、WHERE句も1件しか返さなそうな条件にすること。
//              これは、返り値をシンプルに抑えるため、SELECT文で複数件ヒットしても、先頭行、先頭列しか返せません。
//***************************************************************************************************
BSTR __stdcall ExecuteSQLite(const wchar_t* dbPath, const wchar_t* sql)
{
    // COMの機能(BSTR)を使うので初期化
    CoInitialize(NULL);

    sqlite3* db = nullptr;
    BSTR bstrResult = nullptr; // 結果を格納するBSTRポインタを直接宣言

    // データベースを開く (UTF-16対応のopen16を使用)
    if (sqlite3_open16(dbPath, &db) != SQLITE_OK) {
        MessageBoxW(nullptr, L"データベースを開けませんでした。", L"SQLiteエラー", MB_ICONERROR);
        CoUninitialize();
        return nullptr;
    }

    // --- SQL文がSELECTか、それ以外かを判定 ---
    std::wstring sql_str = sql;
    std::wstring sql_upper = sql_str;
    // SQL文を大文字に変換して、先頭が "SELECT" かどうかをチェック
    std::transform(sql_upper.begin(), sql_upper.end(), sql_upper.begin(), ::toupper);

    // 先頭の空白をトリム
    sql_upper.erase(0, sql_upper.find_first_not_of(L" \t\n\r"));

    if (sql_upper.rfind(L"SELECT", 0) == 0) {
        // --- SELECT文の処理 ---
        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare16_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
            // 最初の行を取得
            if (sqlite3_step(stmt) == SQLITE_ROW) {
                // 最初の列(0)をテキストとして取得
                const wchar_t* text = (const wchar_t*)sqlite3_column_text(stmt, 0);
                if (text) {
                    // 取得したテキストの長さを取得し、その長さでBSTRを割り当てる
                    bstrResult = SysAllocStringByteLen((const CHAR*)text , wcslen(text));
                }
            }
            // else: 行が見つからなかった場合は、bstrResultは空のまま
        }
        else {
            MessageBoxW(nullptr, (const wchar_t*)sqlite3_errmsg16(db), L"SQLite SELECTエラー", MB_ICONERROR);
        }
        sqlite3_finalize(stmt);
    }
    else {
        // --- UPDATE, DELETE, INSERT などの処理 ---
        char* errMsg = nullptr;
        // sqlite3_execはUTF-8で動作するため、ワイド文字列を変換する必要がある
        // ワイド文字列をUTF-8に変換
        int utf8_len = WideCharToMultiByte(CP_UTF8, 0, sql, -1, NULL, 0, NULL, NULL);
        char* utf8_sql = new char[utf8_len];
        WideCharToMultiByte(CP_UTF8, 0, sql, -1, utf8_sql, utf8_len, NULL, NULL);

        if (sqlite3_exec(db, utf8_sql, 0, 0, &errMsg) == SQLITE_OK) {
            // 成功した場合、影響を受けた行数を取得
            int changes = sqlite3_changes(db);
            std::wstring changesStr = std::to_wstring(changes);

            // 変換した文字列とその長さでBSTRを割り当てる
            bstrResult = SysAllocStringByteLen((const CHAR*)changesStr.c_str(), changesStr.length());
        }
        else {
            // エラーメッセージをワイド文字列に変換して表示
            int wide_len = MultiByteToWideChar(CP_UTF8, 0, errMsg, -1, NULL, 0);
            wchar_t* wide_errMsg = new wchar_t[wide_len];
            MultiByteToWideChar(CP_UTF8, 0, errMsg, -1, wide_errMsg, wide_len);
            MessageBoxW(nullptr, wide_errMsg, L"SQLite 操作エラー", MB_ICONERROR);
            delete[] wide_errMsg;
            sqlite3_free(errMsg);
        }
        delete[] utf8_sql;
    }

    sqlite3_close(db);
    CoUninitialize();

    // 作成したBSTRポインタをそのまま返す
    return bstrResult;
}
