#include "AppNotificationBuilder.h"

#include <combaseapi.h>  // CoInitializeExのために必要

using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;


void __stdcall ShowToastNotification(
    LPCWSTR appUserModelID,  // アプリケーションID
    LPCWSTR xmlTemplate,     // XMLテンプレート
    LPCWSTR group,           // グループ
    LPCWSTR tag              // タグ
) {
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
        // トースト通知の作成
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(appUserModelID);
        XmlDocument toastXml;
        toastXml.LoadXml(xmlTemplate);  // XMLテンプレートをロード

        // トースト通知オブジェクトを作成
        ToastNotification toast{ toastXml };

        // グループとタグを設定
        toast.Group(group);
        toast.Tag(tag);

        // トーストを表示
        toastNotifier.Show(toast);
    }
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }
}