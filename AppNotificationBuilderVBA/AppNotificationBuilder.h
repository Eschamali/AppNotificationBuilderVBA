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

// 構造体で、定義します。ここでは、文字列に関するパラメーターです
// ※VBA側で、シグネチャ（型や順序）が合うようにすること。例外として、BOOLはlongで渡さないと上手くいきません
#pragma pack(4)
struct ToastNotificationParams {
    const wchar_t* AppUserModelID;
    const wchar_t* XmlTemplate;
    const wchar_t* Tag;
    const wchar_t* Group;
    const wchar_t* Schedule_ID;
    BOOL ExpiresOnReboot;
    BOOL SuppressPopup;
    double Schedule_DeliveryTime;
    double ExpirationTime;

};
#pragma pack()

//関数宣言
extern "C" AppNotificationBuilderVBA_API void __stdcall ShowToastNotification(ToastNotificationParams* ToastConfigData);    //一般的な通知
extern "C" AppNotificationBuilderVBA_API void __stdcall ShowToastNotificationWithProgressBar(ToastNotificationParams* ToastConfigData, const wchar_t* ProgressStatus, double ProgressValue, const wchar_t* ProgressTitle, const wchar_t* ProgressValueStringOverride);    //プログレスバー付き通知
