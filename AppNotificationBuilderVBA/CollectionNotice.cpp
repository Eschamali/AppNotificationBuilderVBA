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
}

//***************************************************************************************************
//* 機能　　：コレクションを使用したトースト通知のグループ化を削除します。エラーコード返却に対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//***************************************************************************************************
long __stdcall DeleteToastCollection(ToastNotificationParams* ToastConfigData) {
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
}
