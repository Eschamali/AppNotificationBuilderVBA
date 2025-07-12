#pragma once                                            //おまじない

//必要なライブラリ等を読み込む
#include <windows.h>                                    //Windows APIの基本的な型や関数を含む(今回は、Windows::Foundation::DateTime で使用します)
#include <winrt/base.h>                                 //WinRT-APIの基本モジュール
#include <winrt/Windows.UI.Notifications.h>             //トースト関連モジュール全般
#include <winrt/Windows.UI.Notifications.Management.h>  //トースト関連モジュール制御
#include <winrt/Windows.Data.Xml.Dom.h>                 //XMLコンテンツ解析モジュール
#include <winrt/Windows.Foundation.h>                   //Nullable 系を扱うためのモジュール
#include <winrt/Windows.Foundation.Collections.h>       //NotificationDataを扱うためのモジュール
#include <atlbase.h>                                    //Excelインスタンス制御関連
#include <comdef.h>                                     //デバッグによるエラーチェック用
#include <oleacc.h>                                     //AccessibleObjectFromWindow の使用に必要
#include <tlhelp32.h>                                   //PID→HWND の特定に CreateToolhelp32Snapshot を使うために必要

//名前定義を用意
using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;
using namespace winrt::Windows::Foundation;


//外部参照設定つまりはVBAからでもアクセスできるようにする設定。おまじないと思ってください。
//詳細→https://liclog.net/vba-dll-create-1/
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
    const wchar_t* CollectionID;
    BOOL ExpiresOnReboot;
    BOOL SuppressPopup;
    double Schedule_DeliveryTime;
    double ExpirationTime;

};

struct ToastNotificationVariable {
    const wchar_t* TitleText;
    const wchar_t* ContentsText;
    const wchar_t* AttributionText;
    const wchar_t* ProgressTitle;
    const wchar_t* ProgressValueStringOverride;
    const wchar_t* ProgressStatus;
    double ProgressValue;
};
#pragma pack()


//このプロジェクト用パブリック関数を宣言(主にイベント系)
void OnActivated(ToastNotification const& sender, IInspectable const& args);                //ToastNotification.Activated イベント
void OnDismissed(ToastNotification const& sender, ToastDismissedEventArgs const& args);     //ToastNotification.Dismissed イベント
void OnFailed(ToastNotification const& sender, ToastFailedEventArgs const& args);           //ToastNotification.Failed イベント


//----------関数宣言----------
//通知の基本機能
extern "C" AppNotificationBuilderVBA_API void __stdcall ShowToastNotification(ToastNotificationParams* ToastConfigData, ToastNotificationVariable* ToastUpdata);    //通知を表示
extern "C" AppNotificationBuilderVBA_API long __stdcall UpdateToastNotification(ToastNotificationParams* ToastConfigData, ToastNotificationVariable* ToastUpdata);  //通知更新
extern "C" AppNotificationBuilderVBA_API void __stdcall RemoveToastNotification(ToastNotificationParams* ToastConfigData);                                          //通知削除
extern "C" AppNotificationBuilderVBA_API long __stdcall CheckNotificationSetting(ToastNotificationParams* ToastConfigData);                                         //設定確認
//Collection通知によるグループ化機能
extern "C" AppNotificationBuilderVBA_API long __stdcall CreateToastCollection(ToastNotificationParams* ToastConfigData, const wchar_t* displayName, const wchar_t* launchArgs, const wchar_t* iconUri);  //トーストCollectionの作成
extern "C" AppNotificationBuilderVBA_API long __stdcall DeleteToastCollection(ToastNotificationParams* ToastConfigData);  //トーストCollectionを削除
//wpndatabase.db を SQLite で操作する関数
extern "C" AppNotificationBuilderVBA_API BSTR __stdcall ExecuteSQLite(const wchar_t* dbPath, const wchar_t* sql);
