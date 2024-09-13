#include <windows.h>  // Windows APIの基本的な型や関数を含む

#include <winrt/Windows.UI.Notifications.h>
#include <winrt/Windows.Data.Xml.Dom.h>
#include <winrt/base.h>

#ifdef AppNotificationBuilderVBA_EXPORTS
#define AppNotificationBuilderVBA_API __declspec(dllexport)
#else
#define AppNotificationBuilderVBA_API __declspec(dllimport)
#endif

//関数宣言
extern "C" AppNotificationBuilderVBA_API void __stdcall ShowToastNotification(
    LPCWSTR appUserModelID,  // アプリケーションID
    LPCWSTR xmlTemplate,     // XMLテンプレート
    LPCWSTR group,           // グループ
    LPCWSTR tag              // タグ
);