﻿#include <windows.h>  // Windows APIの基本的な型や関数を含む

#include <winrt/Windows.UI.Notifications.h>
#include <winrt/Windows.Data.Xml.Dom.h>
#include <winrt/base.h>
#include <string>
#include <chrono>

#ifdef AppNotificationBuilderVBA_EXPORTS
#define AppNotificationBuilderVBA_API __declspec(dllexport)
#else
#define AppNotificationBuilderVBA_API __declspec(dllimport)
#endif

// 構造体を定義
struct ToastNotificationParams {
    const wchar_t* AppUserModelID;
    const wchar_t* XmlTemplate;
    const wchar_t* Group;
    const wchar_t* Tag;
    double Schedule_DeliveryTime;
    const wchar_t* Schedule_ID;
};


//関数宣言
extern "C" AppNotificationBuilderVBA_API void ShowToastNotification(ToastNotificationParams* params);
