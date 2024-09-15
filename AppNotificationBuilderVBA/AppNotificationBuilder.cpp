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

void __stdcall ShowToastNotification(ToastNotificationParams_String* ToastConfigData_String, ToastNotificationParams_Boolean* ToastConfigData_Boolean, ToastNotificationParams_Date* ToastConfigData_Date){
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
        //値Check用
        MessageBoxW(nullptr, ToastConfigData_String->AppUserModelID, L"AppUserModelID", MB_OK);
        MessageBoxW(nullptr, ToastConfigData_String->XmlTemplate, L"XmlTemplate", MB_OK);
        MessageBoxW(nullptr, ToastConfigData_String->Tag, L"Tag", MB_OK);
        MessageBoxW(nullptr, ToastConfigData_String->Group, L"Group", MB_OK);
        MessageBoxW(nullptr, ToastConfigData_String->Schedule_ID, L"Schedule_ID", MB_OK);

        if (ToastConfigData_Boolean->ExpiresOnReboot) {
            MessageBoxW(nullptr, L"ExpiresOnReboot is TRUE", L"ExpiresOnReboot", MB_OK);
        }
        else {
            MessageBoxW(nullptr, L"ExpiresOnReboot is FALSE", L"ExpiresOnReboot", MB_OK);
        }

        if (ToastConfigData_Boolean->SuppressPopup) {
            MessageBoxW(nullptr, L"SuppressPopup is TRUE", L"SuppressPopup", MB_OK);
        }
        else {
            MessageBoxW(nullptr, L"SuppressPopup is FALSE", L"SuppressPopup", MB_OK);
        }

        wchar_t buffer[256];
        swprintf(buffer, 256, L"ScheduleTime: %f", ToastConfigData_Date->Schedule_DeliveryTime);
        MessageBoxW(nullptr, buffer, L"Schedule Time", MB_OK);

        swprintf(buffer, 256, L"ExpirationTime: %f", ToastConfigData_Date->ExpirationTime);
        MessageBoxW(nullptr, buffer, L"ExpirationTime", MB_OK);

        // トースト通知のXMLを構築
        XmlDocument toastXml;
        toastXml.LoadXml(ToastConfigData_String->XmlTemplate);

        //通知の有効期限が設定されてあったら、設定値を準備する
        SYSTEMTIME ex;
        Windows::Foundation::DateTime ExpirationTimeValue;
        if (ToastConfigData_Date->ExpirationTime > 0) {
            //変換処理
            VariantTimeToSystemTime(ToastConfigData_Date->ExpirationTime, &ex);
            ExpirationTimeValue = SystemTimeToDateTime(ex);
        }

        // 通常の通知か、スケジュール通知かを分岐
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(ToastConfigData_String->AppUserModelID);
        if (ToastConfigData_Date->Schedule_DeliveryTime > 0) {
            // スケジュール通知の場合
            SYSTEMTIME sc;
            VariantTimeToSystemTime(ToastConfigData_Date->Schedule_DeliveryTime, &sc);

            // SYSTEMTIMEをDateTimeに変換
            Windows::Foundation::DateTime scheduleDateTime = SystemTimeToDateTime(sc);

            // スケジュールされたトースト通知を作成
            ScheduledToastNotification scheduledToast(toastXml, scheduleDateTime);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            scheduledToast.Id(ToastConfigData_String->Schedule_ID);
            scheduledToast.Group(ToastConfigData_String->Group);
            scheduledToast.Tag(ToastConfigData_String->Tag);
            scheduledToast.SuppressPopup(ToastConfigData_Boolean->SuppressPopup);
            //if (ToastConfigData_Date->ExpirationTime > 0) scheduledToast.ExpirationTime(ExpirationTimeValue);

            // スケジュールトーストを追加
            toastNotifier.AddToSchedule(scheduledToast);
        }
        else {
            // 通常のトースト通知を作成
            ToastNotification toast(toastXml);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            toast.ExpiresOnReboot(ToastConfigData_Boolean->ExpiresOnReboot);
            toast.Group(ToastConfigData_String->Group);
            toast.Tag(ToastConfigData_String->Tag);
            toast.SuppressPopup(ToastConfigData_Boolean->SuppressPopup);
            //if (ToastConfigData_Date->ExpirationTime > 0) toast.ExpirationTime(ExpirationTimeValue);

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
