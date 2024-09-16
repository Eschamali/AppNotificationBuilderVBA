#include "AppNotificationBuilder.h"

#include <combaseapi.h>  // CoInitializeExのために必要

using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;
using namespace winrt::Windows::Foundation;


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


//引数に渡された値で、単純なトースト通知を表示します。指定日時に通知するスケジュール機能も対応します
void __stdcall ShowToastNotification(ToastNotificationParams* ToastConfigData){
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
        //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->XmlTemplate, L"XmlTemplate", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Tag, L"Tag", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Group, L"Group", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Schedule_ID, L"Schedule_ID", MB_OK);

        //if (ToastConfigData->ExpiresOnReboot) {
        //    MessageBoxW(nullptr, L"ExpiresOnReboot is TRUE", L"ExpiresOnReboot", MB_OK);
        //}
        //else {
        //    MessageBoxW(nullptr, L"ExpiresOnReboot is FALSE", L"ExpiresOnReboot", MB_OK);
        //}

        //if (ToastConfigData->SuppressPopup) {
        //    MessageBoxW(nullptr, L"SuppressPopup is TRUE", L"SuppressPopup", MB_OK);
        //}
        //else {
        //    MessageBoxW(nullptr, L"SuppressPopup is FALSE", L"SuppressPopup", MB_OK);
        //}

        //wchar_t buffer[256];
        //swprintf(buffer, 256, L"ScheduleTime: %f", ToastConfigData->Schedule_DeliveryTime);
        //MessageBoxW(nullptr, buffer, L"Schedule Time", MB_OK);

        //swprintf(buffer, 256, L"ExpirationTime: %f", ToastConfigData->ExpirationTime);
        //MessageBoxW(nullptr, buffer, L"ExpirationTime", MB_OK);

        // トースト通知のXMLを構築
        XmlDocument toastXml;
        toastXml.LoadXml(ToastConfigData->XmlTemplate);

        //通知の有効期限が設定されてあったら、設定値を準備する
        Windows::Foundation::DateTime ExpirationTimeValue;
        if (ToastConfigData->ExpirationTime > 0) {
            //変換処理
            SYSTEMTIME ex;
            VariantTimeToSystemTime(ToastConfigData->ExpirationTime, &ex);
            ExpirationTimeValue = SystemTimeToDateTime(ex);
        }

        //指定のAppUserModelIDで、トーストオブジェクトを生成
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID);
        
        // 日付が指定されているかで、通常の通知か、スケジュール通知かを分岐
        if (ToastConfigData->Schedule_DeliveryTime > 0) {
            // スケジュール通知の場合
            SYSTEMTIME sc;
            VariantTimeToSystemTime(ToastConfigData->Schedule_DeliveryTime, &sc);

            // SYSTEMTIMEをDateTimeに変換
            Windows::Foundation::DateTime scheduleDateTime = SystemTimeToDateTime(sc);

            // スケジュールされたトースト通知を作成
            ScheduledToastNotification scheduledToast(toastXml, scheduleDateTime);

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            scheduledToast.Id(ToastConfigData->Schedule_ID);
            scheduledToast.Group(ToastConfigData->Group);
            scheduledToast.Tag(ToastConfigData->Tag);
            scheduledToast.SuppressPopup(ToastConfigData->SuppressPopup);
            if (ToastConfigData->ExpirationTime > 0) scheduledToast.ExpirationTime(ExpirationTimeValue);

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
            if (ToastConfigData->ExpirationTime > 0) toast.ExpirationTime(ExpirationTimeValue);

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

//引数に渡された値で、最初のトーストの進行状況バーを表示します
void __stdcall ShowToastNotificationWithProgressBar(ToastNotificationParams* ToastConfigData, const wchar_t* ProgressStatus, double ProgressValue, const wchar_t* ProgressTitle, const wchar_t* ProgressValueStringOverride) {
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
        //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->XmlTemplate, L"XmlTemplate", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Tag, L"Tag", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Group, L"Group", MB_OK);

        //if (ToastConfigData->SuppressPopup) {
        //    MessageBoxW(nullptr, L"SuppressPopup is TRUE", L"SuppressPopup", MB_OK);
        //}
        //else {
        //    MessageBoxW(nullptr, L"SuppressPopup is FALSE", L"SuppressPopup", MB_OK);
        //}

        //if (ToastConfigData->ExpiresOnReboot) {
        //    MessageBoxW(nullptr, L"ExpiresOnReboot is TRUE", L"ExpiresOnReboot", MB_OK);
        //}
        //else {
        //    MessageBoxW(nullptr, L"ExpiresOnReboot is FALSE", L"ExpiresOnReboot", MB_OK);
        //}

        //MessageBoxW(nullptr, ProgressStatus, L"ProgressStatus", MB_OK);

        //wchar_t buffer[256];
        //swprintf(buffer, 256, L"ProgressValue: %f", ProgressValue);
        //MessageBoxW(nullptr, buffer, L"ProgressValue", MB_OK);

        //MessageBoxW(nullptr, ProgressTitle, L"ProgressTitle", MB_OK);
        //MessageBoxW(nullptr, ProgressValueStringOverride, L"ProgressValueStringOverride", MB_OK);

        //指定のAppUserModelIDで、トーストオブジェクトを生成
        ToastNotifier toastNotifierWithProgressBar = ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID);

        // プログレスバー付きトースト通知のXMLを構築
        XmlDocument toastXmlWithProgress;
        toastXmlWithProgress.LoadXml(ToastConfigData->XmlTemplate);

        // プログレスバー付きのトースト通知を作成
        ToastNotification toastWithProgress(toastXmlWithProgress);

        //初期のNotificationData値を割り当てる
        //※予め、xmlに仕込む変数名({XXX})をここと統一する必要があります
        NotificationData ProgressParams;
        auto ProgressParamsValues = ProgressParams.Values();         // 戻り値の型を明示的に指定

        ProgressParamsValues.Insert(L"progressTitle", ProgressTitle);                               //タイトル
        ProgressParamsValues.Insert(L"progressStatus", ProgressStatus);                             //左下の進行状況バーの下に表示される状態文字列
        
        //進捗値の場合、負になってたら、ドットアニメーションの不確定式にします。
        if (ProgressValue < 0) {
            ProgressParamsValues.Insert(L"progressValue", L"Indeterminate");                        //進行状況バーの状態を「不確定」として、設定
        }
        else {
            ProgressParamsValues.Insert(L"progressValue", std::to_wstring(ProgressValue).c_str());  //進行状況バーの状態を設定
        }

        //文字列がない場合は、バインディング処理しません
        if (ProgressValueStringOverride) ProgressParamsValues.Insert(L"progressValueString", ProgressValueStringOverride);           //既定のパーセンテージ文字列の代わりに表示される省略可能な文字列を取得または設定します。 これが指定されていない場合は、"70%" のようなものが表示されます。

        //上記のパラメーターをトーストに設定
        toastWithProgress.Data(ProgressParams);

        // 上記で作成されたトーストオブジェクトに各種設定(GroupとTag等)を施す
        toastWithProgress.ExpiresOnReboot(ToastConfigData->ExpiresOnReboot);
        toastWithProgress.Group(ToastConfigData->Group);
        toastWithProgress.Tag(ToastConfigData->Tag);
        toastWithProgress.SuppressPopup(ToastConfigData->SuppressPopup);

        //順序外の更新を防ぐため、シーケンス番号を指定します。初回なので1にします。
        ProgressParams.SequenceNumber(1);

        // プログレスバー通知を作動
        toastNotifierWithProgressBar.Show(toastWithProgress);
    }
    
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }
}


//引数に渡された値で、トーストの進行状況バーを更新します。
long __stdcall UpdateToastNotificationWithProgressBar(ToastNotificationParams* ToastConfigData, const wchar_t* ProgressStatus, double ProgressValue, const wchar_t* ProgressTitle, const wchar_t* ProgressValueStringOverride, long SequenceNumber) {
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

    //更新結果用の変数を定義
    NotificationUpdateResult result;
    try {
        //値Check用
        //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Tag, L"Tag", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Group, L"Group", MB_OK);


        //MessageBoxW(nullptr, ProgressStatus, L"ProgressStatus", MB_OK);

        //wchar_t buffer[256];
        //swprintf(buffer, 256, L"ProgressValue: %f", ProgressValue);
        //MessageBoxW(nullptr, buffer, L"ProgressValue", MB_OK);

        //MessageBoxW(nullptr, ProgressTitle, L"ProgressTitle", MB_OK);
        //MessageBoxW(nullptr, ProgressValueStringOverride, L"ProgressValueStringOverride", MB_OK);

        //指定のAppUserModelIDで、トーストオブジェクトを生成
        ToastNotifier toastNotifierWithProgressBar = ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID);

        //更新のNotificationData値を割り当てる
        //※予め、xmlに仕込む変数名({XXX})をここと統一する必要があります
        NotificationData ProgressParams;
        auto ProgressParamsValues = ProgressParams.Values();         // 戻り値の型を明示的に指定

        ProgressParamsValues.Insert(L"progressTitle", ProgressTitle);                               //タイトル
        ProgressParamsValues.Insert(L"progressStatus", ProgressStatus);                             //左下の進行状況バーの下に表示される状態文字列

        //進捗値の場合、負になってたら、ドットアニメーションの不確定式にします。
        if (ProgressValue < 0) {
            ProgressParamsValues.Insert(L"progressValue", L"Indeterminate");                        //進行状況バーの状態を「不確定」として、設定
        }
        else {
            ProgressParamsValues.Insert(L"progressValue", std::to_wstring(ProgressValue).c_str());  //進行状況バーの状態を設定
        }

        //文字列がない場合は、バインディング処理しません
        if (ProgressValueStringOverride) ProgressParamsValues.Insert(L"progressValueString", ProgressValueStringOverride);           //既定のパーセンテージ文字列の代わりに表示される省略可能な文字列を取得または設定します。 これが指定されていない場合は、"70%" のようなものが表示されます。

        //順序外の更新を防ぐため、シーケンス番号を指定します。
        ProgressParams.SequenceNumber(SequenceNumber);

        // トースト通知を更新
        result = toastNotifierWithProgressBar.Update(ProgressParams, ToastConfigData->Tag, ToastConfigData->Group);
    }

    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }

    //結果値を返す
    return static_cast<long>(result);
}