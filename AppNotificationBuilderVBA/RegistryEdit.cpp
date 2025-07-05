/***************************************************************************************************
*									    通知の挙動設定 その2
***************************************************************************************************
* レジストリエディターのうち、下記のトースト関連の隠し設定を編集します。下記の箇所以外のレジストリは編集しません。
* ・コンピューター\HKEY_CLASSES_ROOT\AppUserModelId\[AppUserModelID名称]
* 
* 管理者権限が必要な領域のため、ユーザーアカウント制御機能をかまして、実現します。
***************************************************************************************************/



//設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"



//***************************************************************************************************
//                                  ■■■ 静的定数 ■■■
//***************************************************************************************************
constexpr wchar_t ParameterDelimiter = L'|';                       //レジストリ書き込み処理を行う際の区切り文字(1文字分)
constexpr const wchar_t* TargetRegistryPath = L"AppUserModelId\\";  //レジストリ操作を行う相対パス



//***************************************************************************************************
//                                  ■■■ ヘルパー用関数 ■■■
//***************************************************************************************************
//* 機能　　：引数に従い、文字列を分割します。
//---------------------------------------------------------------------------------------------------
//* 返り値  ：※割愛
//* 引数　　：s             文字列本体
//            delimiter     分割の目印とする1文字
//***************************************************************************************************
static std::vector<std::wstring> split(const std::wstring& s, wchar_t delimiter) {
    std::vector<std::wstring> tokens;
    std::wstring token;
    std::wstringstream tokenStream(s);
    while (std::getline(tokenStream, token, delimiter)) {
        tokens.push_back(token);
    }
    return tokens;
}

//***************************************************************************************************
//* 機能　　：レジストリ書き込み処理を行います
//---------------------------------------------------------------------------------------------------
//* 返り値  ：実行結果コード
//* 引数　　：combinedArgs   法則に則った所定の区切りパラメーター
//            →AppUserModelID|値の名前|型ID(0~2)|値そのもの
//---------------------------------------------------------------------------------------------------
//* 詳細説明：コンピューター\HKEY_CLASSES_ROOT\AppUserModelId\[AppUserModelID名称]　に対して、隠し設定を行います。
//* 注意事項：管理者権限付与状態の rundll32 を実行する関係上、所定の区切りパラメーターを採用しています。
//            　そのため、引数自体に所定の区切り文字があると、正常動作しません。
//***************************************************************************************************
static HRESULT WriteKeyFromArgs(const std::wstring& combinedArgs) {
    //所定の区切り文字を基に、パラメーターを抽出
    auto args = split(combinedArgs, ParameterDelimiter);
    if (args.size() < 4) {
        return E_INVALIDARG; // 引数が足りない
    }

    //抽出した4つのパラメータを、格納
    const std::wstring& aumid = args[0];
    const std::wstring& valueName = args[1];
    long valueType = _wtol(args[2].c_str());
    const std::wstring& valueData = args[3];

    //HKEY_CLASSES_ROOT 始まりとして、編集要求
    HKEY hKey;
    std::wstring regPath = TargetRegistryPath + aumid;
    LSTATUS status = RegCreateKeyExW(
        HKEY_CLASSES_ROOT,
        regPath.c_str(),
        0, NULL, REG_OPTION_NON_VOLATILE,
        KEY_SET_VALUE, NULL, &hKey, NULL
    );
    if (status != ERROR_SUCCESS) return HRESULT_FROM_WIN32(status);

    //値の型に応じた処理。ひとまず3種類用意します。
    switch (valueType) {
    case 0: // REG_DWORD
    {
        DWORD dwordValue = _wtol(valueData.c_str());
        status = RegSetValueExW(hKey, valueName.c_str(), 0, REG_DWORD, (const BYTE*)&dwordValue, sizeof(dwordValue));
        break;
    }
    case 1: // REG_SZ
    {
        status = RegSetValueExW(hKey, valueName.c_str(), 0, REG_SZ, (const BYTE*)valueData.c_str(), (valueData.length() + 1) * sizeof(wchar_t));
        break;
    }
    case 2: // REG_EXPAND_SZ
    {
        status = RegSetValueExW(hKey, valueName.c_str(), 0, REG_EXPAND_SZ, (const BYTE*)valueData.c_str(), (valueData.length() + 1) * sizeof(wchar_t));
        break;
    }
    default:
        status = ERROR_INVALID_PARAMETER;
        break;
    }

    //後始末して、結果コードを返す
    RegCloseKey(hKey);
    return HRESULT_FROM_WIN32(status);
}



//***************************************************************************************************
//                             ■■■ rundll32側から呼び出す関数 ■■■
//***************************************************************************************************
//* 機能　　：引数を受け取って、レジストリ登録へ進みます。
//---------------------------------------------------------------------------------------------------
//* 引数　　：lpszCmdLine    法則に則った所定の区切りパラメーター
//            →AppUserModelID|値の名前|型ID(0~2)|値そのもの    
//---------------------------------------------------------------------------------------------------
//* 詳細説明：管理者権限でこれを呼び出す仕組みを使えば、レジストリ登録が可能です。
//***************************************************************************************************
void __stdcall WriteRegistryAsAdmin(HWND hwnd, HINSTANCE hinst, LPSTR lpszCmdLine, int nCmdShow)
{
    //文字サイズを動的確保
    int required = MultiByteToWideChar(CP_ACP, 0, lpszCmdLine, -1, NULL, 0);
    WCHAR* wideCmdLine = new WCHAR[required];

    // Shift_JIS → UTF-16 に変換（ANSI想定）
    MultiByteToWideChar(CP_ACP, 0, lpszCmdLine, -1, wideCmdLine, 512);

    // wideCmdLine を使って処理
    //MessageBoxW(NULL, wideCmdLine, L"受け取った引数", MB_OK);

    //レジストリ登録へ移る
    CoInitialize(NULL);
    HRESULT hr = WriteKeyFromArgs(wideCmdLine);

    //結果を表示する
    //if (SUCCEEDED(hr)) {
    //    MessageBoxW(nullptr, L"レジストリの設定に成功しました。", L"成功", MB_OK);
    //}
    //else {
    //    _com_error err(hr);
    //    wchar_t buf[512];
    //    swprintf_s(buf, L"レジストリ書き込みに失敗しました。\nHRESULT=0x%08X\n%s", hr, err.ErrorMessage());
    //    MessageBoxW(nullptr, buf, L"エラー", MB_OK);
    //}

    //解放
    CoUninitialize();
}



//***************************************************************************************************
//                             ■■■ VBA側から呼び出す関数 ■■■
//***************************************************************************************************
//* 機能　　：引数を受け取って、レジストリ登録を試みます。
//---------------------------------------------------------------------------------------------------
//* 返り値  ：結果コード
//* 引数　　：lpszCmdLine    法則に則った所定の区切りパラメーター
//            →AppUserModelID|値の名前|型ID(0~2)|値そのもの    
//---------------------------------------------------------------------------------------------------
//* 詳細説明：Excelを管理者権限で実行していない場合、ユーザーアカウント制御機能を利用して昇格を促します。
//***************************************************************************************************
long __stdcall AttemptToWriteRegistry(HWND hwnd, HINSTANCE hinst, LPWSTR lpszCmdLine, int nCmdSho)
{
    //MessageBoxW(nullptr, lpszCmdLine, L"引数の中身", MB_OK);
    CoInitialize(NULL);

    // このDLL のハンドル取得を試みる
    HMODULE hModule = NULL;
    HRESULT hr = GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        (LPCWSTR)AttemptToWriteRegistry,
        &hModule
    );

    if (hModule == NULL) {
        MessageBoxW(nullptr, L"自分自身のDLLハンドルを取得できませんでした。", L"致命的エラー", MB_OK);
        CoUninitialize();
        return hr;
    }

    //引数を組み立てる
    long Result;
    wchar_t dllPath[MAX_PATH];
    GetModuleFileNameW(hModule, dllPath, MAX_PATH); // ★取得した自分自身のハンドルを使う！
    std::wstring params = L"\"";
    params += dllPath;
    params += L"\",WriteRegistryAsAdmin ";  // rundll32 に実行させたい関数名
    params += lpszCmdLine;                 // AUMID|Name|Type|Data の文字列
    
    //ユーザーアカウント制御 に渡すパラメーター一式
    SHELLEXECUTEINFOW sei = { sizeof(sei) };
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb = L"runas";
    sei.lpFile = L"rundll32.exe";
    sei.lpParameters = params.c_str();
    sei.nShow = SW_HIDE;
    
    //ユーザーアカウント制御 を表示させる
    //※Excel自体が、管理者権限を持っている場合、即 True を返します。
    if (ShellExecuteExW(&sei)) {
        //許可されたら、rundll32 側の処理が終わるまで待つ
        WaitForSingleObject(sei.hProcess, INFINITE);
        CloseHandle(sei.hProcess);

        Result = -1;
        //MessageBoxW(nullptr, L"権限昇格を伴う処理が完了しました。", L"情報", MB_OK);
    }
    else {
        //拒否された場合
        Result = -2;
        //MessageBoxW(nullptr, L"操作はユーザーによってキャンセルされました。", L"キャンセル", MB_OK);
    }
    //解放
    CoUninitialize();
    return Result;
}
