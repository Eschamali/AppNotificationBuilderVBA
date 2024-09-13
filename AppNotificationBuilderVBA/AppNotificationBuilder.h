#include <windows.h>  // Windows API�̊�{�I�Ȍ^��֐����܂�

#include <winrt/Windows.UI.Notifications.h>
#include <winrt/Windows.Data.Xml.Dom.h>
#include <winrt/base.h>

#ifdef AppNotificationBuilderVBA_EXPORTS
#define AppNotificationBuilderVBA_API __declspec(dllexport)
#else
#define AppNotificationBuilderVBA_API __declspec(dllimport)
#endif

//�֐��錾
extern "C" AppNotificationBuilderVBA_API void __stdcall ShowToastNotification(
    LPCWSTR appUserModelID,  // �A�v���P�[�V����ID
    LPCWSTR xmlTemplate,     // XML�e���v���[�g
    LPCWSTR group,           // �O���[�v
    LPCWSTR tag              // �^�O
);