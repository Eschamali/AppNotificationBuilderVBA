#include "AppNotificationBuilder.h"

#include <combaseapi.h>  // CoInitializeExのために必要

using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;


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
        // トースト通知の作成
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(appUserModelID);
        XmlDocument toastXml;
        toastXml.LoadXml(xmlTemplate);  // XMLテンプレートをロード

        // トースト通知オブジェクトを作成
        ToastNotification toast{ toastXml };

        // グループとタグを設定
        toast.Group(group);
        toast.Tag(tag);

        //MessageBoxW(nullptr, appUserModelID, L"AppUserModelID", MB_OK);
        //MessageBoxW(nullptr, xmlTemplate, L"XML Template", MB_OK);
        //MessageBoxW(nullptr, group, L"Group", MB_OK);
        //MessageBoxW(nullptr, tag, L"Tag", MB_OK);
        wchar_t buffer[256];
        swprintf(buffer, 256, L"ScheduleTime: %f", scheduleTime);
        MessageBoxW(nullptr, buffer, L"Schedule Time", MB_OK);

        if (scheduleTime > 0) {
            // スケジュール日時が指定されている場合
            // VBAの日付型は1970年からの秒数で与えられるので、それを日時に変換する
            FILETIME fileTime;
            SYSTEMTIME systemTime;
            LARGE_INTEGER largeInt;

            // VBAから渡されるシリアル値（days since 1899-12-30）をWindows FILETIME形式に変換
            double daysSinceEpoch = scheduleTime - 25569.0;  // 25569は1970-01-01からのシリアル値
            long long total100Nanoseconds = static_cast<long long>(daysSinceEpoch * 864000000000.0);  // 1日=86400秒, 100ナノ秒=1秒/10^7

            swprintf(buffer, 256, L"daysSinceEpoch: %f", daysSinceEpoch);
            MessageBoxW(nullptr, buffer, L"Schedule Time", MB_OK);

            largeInt.QuadPart = total100Nanoseconds;
            fileTime.dwLowDateTime = largeInt.LowPart;
            fileTime.dwHighDateTime = largeInt.HighPart;

            // FILETIMEをSYSTEMTIMEに変換
            FileTimeToSystemTime(&fileTime, &systemTime);

            // SYSTEMTIMEをDateTimeに変換
            Windows::Foundation::DateTime scheduleDateTime;
            scheduleDateTime = winrt::clock::from_FILETIME(fileTime);  // FILETIMEを使ってDateTimeに変換

            // ScheduledToastNotification を作成してスケジュール設定
            ScheduledToastNotification scheduledToast{ toastXml, scheduleDateTime };
            scheduledToast.Group(group);
            scheduledToast.Tag(tag);

            // 通知をスケジュール
            toastNotifier.AddToSchedule(scheduledToast);
        }
        else {
            // 即時通知を表示
            toastNotifier.Show(toast);
        }
    }
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }
}
