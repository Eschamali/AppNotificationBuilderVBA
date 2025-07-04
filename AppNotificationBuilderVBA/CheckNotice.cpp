/***************************************************************************************************
 *										    設定確認
 ***************************************************************************************************
 * 以下の機能を記述します
 * ・トースト通知の表示に関する制限を確認します。
 *
 *
 * URL
 * https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.notificationsetting
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"



//***************************************************************************************************
//                              ■■■ VBA側から呼び出す関数 ■■■
//***************************************************************************************************
//* 機能　　：トースト通知の表示に制限があるかの確認を行います
//---------------------------------------------------------------------------------------------------
//* 返り値　：NotificationSetting 列挙型。詳細は、タイトルにあるURLを
//* 引数　　：ToastConfigData            ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//***************************************************************************************************
long __stdcall CheckNotificationSetting(ToastNotificationParams* ToastConfigData) {
    try {

        //コレクションIDが指定されてあったら、そっちのオブジェクトを使用する
        NotificationSetting currentSetting =
            ToastConfigData->CollectionID
            ? ToastNotificationManager::GetDefault().GetToastNotifierForToastCollectionIdAsync(ToastConfigData->CollectionID).get().Setting()
            : ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID).Setting();

        // WinRTのenumをlongに静的キャストして返す
        long returnValue = static_cast<long>(currentSetting);

        //  結果を返す
        return returnValue;
    }
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        return e.code();
    }
}
