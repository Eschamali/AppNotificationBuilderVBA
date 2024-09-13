#include "AppNotificationBuilder.h"

#include <combaseapi.h>  // CoInitializeEx�̂��߂ɕK�v

using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;


void __stdcall ShowToastNotification(
    LPCWSTR appUserModelID,  // �A�v���P�[�V����ID
    LPCWSTR xmlTemplate,     // XML�e���v���[�g
    LPCWSTR group,           // �O���[�v
    LPCWSTR tag              // �^�O
) {
    // COM�̏�����
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (hr == RPC_E_CHANGED_MODE) {
        // ���ɈقȂ�A�p�[�g�����g ���[�h�ŏ���������Ă���ꍇ�́A���̂܂ܑ��s
    }
    else if (FAILED(hr)) {
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, 256, L"COM�������Ɏ��s���܂����BHRESULT: 0x%08X", hr);
        MessageBoxW(nullptr, errorMsg, L"�G���[", MB_OK);
        return;
    }

    try {
        // �g�[�X�g�ʒm�̍쐬
        ToastNotifier toastNotifier = ToastNotificationManager::CreateToastNotifier(appUserModelID);
        XmlDocument toastXml;
        toastXml.LoadXml(xmlTemplate);  // XML�e���v���[�g�����[�h

        // �g�[�X�g�ʒm�I�u�W�F�N�g���쐬
        ToastNotification toast{ toastXml };

        // �O���[�v�ƃ^�O��ݒ�
        toast.Group(group);
        toast.Tag(tag);

        // �g�[�X�g��\��
        toastNotifier.Show(toast);
    }
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"�G���[", MB_OK);
    }

    // CoUninitialize()�́ACoInitializeEx�����������ꍇ�̂݌Ăяo��
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }
}