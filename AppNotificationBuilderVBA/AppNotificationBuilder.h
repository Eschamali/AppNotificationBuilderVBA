#include <windows.h>  // Windows APIの基本的な型や関数を含む

#include <winrt/Windows.UI.Notifications.h> //トースト関連モジュール全般
#include <winrt/Windows.Data.Xml.Dom.h>     //XMLコンテンツ解析モジュール
#include <winrt/Windows.Foundation.h>       //Nullable 系を扱うためのモジュール
#include <winrt/base.h>                     //WinRT-APIの基本モジュール
#include <string>
#include <chrono>

#ifdef AppNotificationBuilderVBA_EXPORTS
#define AppNotificationBuilderVBA_API __declspec(dllexport)
#else
#define AppNotificationBuilderVBA_API __declspec(dllimport)
#endif

// 1つ目の構造体を定義。スタック領域の制限の都合上、複数の構造体を定義します。ここでは、文字列に関するパラメーターです
// ※VBA側で、シグネチャ（型や順序）が合うようにすること。
#pragma pack(4)
struct ToastNotificationParams_String {
    const wchar_t* AppUserModelID;
    const wchar_t* XmlTemplate;
    const wchar_t* Tag;
    const wchar_t* Group;
    const wchar_t* Schedule_ID;
};

// 2つ目の構造体を定義。ここでは、スイッチングに関するパラメーターです
struct ToastNotificationParams_Boolean {
    BOOLEAN ExpiresOnReboot;
    BOOLEAN SuppressPopup;
};

// 3つ目の構造体を定義。ここでは、日付に関するパラメーターです
struct ToastNotificationParams_Date {
    double Schedule_DeliveryTime;
    double ExpirationTime;
};
#pragma pack()

//関数宣言
extern "C" AppNotificationBuilderVBA_API void __stdcall ShowToastNotification(ToastNotificationParams_String* ToastConfigData_String, ToastNotificationParams_Boolean* ToastConfigData_Boolean, ToastNotificationParams_Date* ToastConfigData_Date);
