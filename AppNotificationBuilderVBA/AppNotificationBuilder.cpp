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

void ShowToastNotification(ToastNotificationParams* params){
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

    //必要な設定値などを引数から取得
    const wchar_t* appUserModelID = params->appUserModelID;
    const wchar_t* xmlTemplate = params->xmlTemplate;
    const wchar_t* group = params->group;
    const wchar_t* tag = params->tag;
    double scheduleTime = params->scheduleTime;

    try {
        // トースト通知のXMLを構築
        XmlDocument toastXml;
        toastXml.LoadXml(xmlTemplate);


        // 通常の通知か、スケジュール通知かを分岐
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(appUserModelID);
        if (scheduleTime > 0) {
            // スケジュール通知の場合
            SYSTEMTIME st;
            VariantTimeToSystemTime(scheduleTime, &st);

            // SYSTEMTIMEをDateTimeに変換
            Windows::Foundation::DateTime scheduleDateTime = SystemTimeToDateTime(st);

            // time_since_epoch() から100ナノ秒単位の値を取得
            auto duration = scheduleDateTime.time_since_epoch();
            int64_t count = std::chrono::duration_cast<std::chrono::duration<int64_t, std::ratio<1, 10000000>>>(duration).count();

            // スケジュールされたトースト通知を作成
            ScheduledToastNotification scheduledToast(toastXml, scheduleDateTime);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            scheduledToast.Group(group);
            scheduledToast.Tag(tag);

            // スケジュールトーストを追加
            toastNotifier.AddToSchedule(scheduledToast);
        }
        else {
            // 通常のトースト通知を作成
            ToastNotification toast(toastXml);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            toast.Group(group);
            toast.Tag(tag);

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
