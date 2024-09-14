#include "AppNotificationBuilder.h"

#include <combaseapi.h>  // CoInitializeExのために必要

using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;


// SYSTEMTIMEをDateTimeに変換する関数
Windows::Foundation::DateTime SystemTimeToDateTime(const SYSTEMTIME& st) {
    FILETIME fileTime;
    SystemTimeToFileTime(&st, &fileTime);

    // FILETIMEをLARGE_INTEGERに変換
    ULARGE_INTEGER largeInt;
    largeInt.LowPart = fileTime.dwLowDateTime;
    largeInt.HighPart = fileTime.dwHighDateTime;

    // FILETIMEの値を100ナノ秒単位で格納し、DateTimeに変換
    Windows::Foundation::DateTime dateTime;
    dateTime = winrt::clock::from_FILETIME(fileTime);
    return dateTime;
}

void ShowToastNotification(ToastNotificationParams* ToastConfigData){
    // COMの初期化
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // 既に異なるアパートメント モードで初期化されている場合は、そのまま続行
    }
    else if (FAILED(hr)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM初期化に失敗しました。HRESULT: 0x%08X", hr);
        MessageBoxW(nullptr, errorMsg, L"エラー", MB_OK);
        return;
    }

    try {
        // トースト通知のXMLを構築
        XmlDocument toastXml;
        toastXml.LoadXml(ToastConfigData->XmlTemplate);


        // 通常の通知か、スケジュール通知かを分岐
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID);
        if (ToastConfigData->Schedule_DeliveryTime > 0) {
            // スケジュール通知の場合
            SYSTEMTIME st;
            VariantTimeToSystemTime(ToastConfigData->Schedule_DeliveryTime, &st);

            // SYSTEMTIMEをDateTimeに変換
            Windows::Foundation::DateTime scheduleDateTime = SystemTimeToDateTime(st);

            // スケジュールされたトースト通知を作成
            ScheduledToastNotification scheduledToast(toastXml, scheduleDateTime);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            scheduledToast.Id(ToastConfigData->Schedule_ID);
            scheduledToast.Group(ToastConfigData->Group);
            scheduledToast.Tag(ToastConfigData->Tag);
            scheduledToast.SuppressPopup(ToastConfigData->SuppressPopup);

            // スケジュールトーストを追加
            toastNotifier.AddToSchedule(scheduledToast);
        }
        else {
            // 通常のトースト通知を作成
            ToastNotification toast(toastXml);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            toast.ExpiresOnReboot(ToastConfigData->ExpiresOnReboot);
            toast.Group(ToastConfigData->Group);
            toast.Tag(ToastConfigData->Tag);
            toast.SuppressPopup(ToastConfigData->SuppressPopup);

            // 通常の即時通知を作動
            toastNotifier.Show(toast);
        }
    }
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }
}
