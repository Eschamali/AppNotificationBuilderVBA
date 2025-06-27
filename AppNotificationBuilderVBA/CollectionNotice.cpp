/***************************************************************************************************
 *										コレクション通知機能
 ***************************************************************************************************
 * 以下の機能を記述してます
 * ・トースト通知のグループ化作成/削除
 *
 * 
 * URL
 * https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-collections
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"



//***************************************************************************************************
//                              ■■■ VBA側から呼び出す関数 ■■■
//***************************************************************************************************
//* 機能　　：引数に渡された値から、コレクションを使用したトースト通知のグループ化を作成します。エラーコード返却に対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData   ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//            displayName       コレクション名
//            launchArgs        起動引数
//            iconUri           アイコンパス
//***************************************************************************************************
long __stdcall CreateToastCollection(ToastNotificationParams* ToastConfigData, const wchar_t* displayName, const wchar_t* launchArgs, const wchar_t* iconUri) {
    // COMの初期化
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // 既に異なるアパートメント モードで初期化されている場合は、そのまま続行
    }
    else if (FAILED(hr)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM初期化に失敗しました。HRESULT: 0x%08X", hr);
        MessageBoxW(nullptr, errorMsg, L"エラー", MB_OK);
        return -1;
    }

    //値Check用
    //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
    //MessageBoxW(nullptr, ToastConfigData->CollectionID, L"collectionId", MB_OK);
    //MessageBoxW(nullptr, displayName, L"displayName", MB_OK);
    //MessageBoxW(nullptr, launchArgs, L"launchArgs", MB_OK);
    //MessageBoxW(nullptr, iconUri, L"iconUri", MB_OK);

    try {
        // トースト通知のマネージャーを取得
        ToastNotificationManagerForUser userManager = ToastNotificationManager::GetDefault();
        ToastCollectionManager collectionManager = userManager.GetToastCollectionManager(ToastConfigData->AppUserModelID);

        //iconUriから、Uri クラスのインスタンスを作成します
        Uri siteUri = Uri(iconUri);

        // コレクションを作成
        ToastCollection collection = ToastCollection(ToastConfigData->CollectionID, displayName, launchArgs, siteUri);
        collectionManager.SaveToastCollectionAsync(collection);

        // 成功したら0を返す
        return 0;
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング (エラーコードを返す)
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        return e.code();
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }

}

//***************************************************************************************************
//* 機能　　：コレクションを使用したトースト通知のグループ化を削除します。エラーコード返却に対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//***************************************************************************************************
long __stdcall DeleteToastCollection(ToastNotificationParams* ToastConfigData) {
    // COMの初期化
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // 既に異なるアパートメント モードで初期化されている場合は、そのまま続行
    }
    else if (FAILED(hr)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM初期化に失敗しました。HRESULT: 0x%08X", hr);
        MessageBoxW(nullptr, errorMsg, L"エラー", MB_OK);
        return -1;
    }

    try {
        // トースト通知のマネージャーを取得
        ToastNotificationManagerForUser userManager = ToastNotificationManager::GetDefault();
        ToastCollectionManager collectionManager = userManager.GetToastCollectionManager(ToastConfigData->AppUserModelID);

        //CollectionIDの定義Check
        if (ToastConfigData->CollectionID) {
            //何かのCollectionIDが指定してあったら、それのみ削除
            collectionManager.RemoveToastCollectionAsync(ToastConfigData->CollectionID);
        }
        else {
            //未定義の場合、全てのToastCollectionを削除
            collectionManager.RemoveAllToastCollectionsAsync();
        }

        // 成功したら0を返す
        return 0;
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング (エラーコードを返す)
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        return e.code();
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }

}
