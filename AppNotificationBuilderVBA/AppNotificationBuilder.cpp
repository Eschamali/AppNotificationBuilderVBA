﻿//設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"

//よく使う名前定義を用意する
using namespace winrt;
using namespace Windows::UI::Notifications;
using namespace Windows::Data::Xml::Dom;
using namespace winrt::Windows::Foundation;



//***************************************************************************************************
//                                 ■■■ 内部のヘルパー関数 ■■■
//***************************************************************************************************
//* 機能　　 ：SYSTEMTIMEをDateTimeに変換します
//---------------------------------------------------------------------------------------------------
//* 引数　　 ：SYSTEMTIME
//* 返り値　 ：dateTime
//***************************************************************************************************
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

//***************************************************************************************************
//* 機能　　：Excel マクロを実行する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：ExcelMacroPass     Action要素のarguments。マクロ名を想定してます。
//            UserInputs         Input要素で入力した内容、あるいはSelect要素のID名称とそれに紐づくInput要素のIDとのセットとなる2次元配列                             
//***************************************************************************************************
void ExecuteExcelMacro(const wchar_t* ExcelMacroPass, SAFEARRAY* UserInputs) {
    //詳細メッセージ、取得用
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(EXCEPINFO));  // 初期化

    // 1. ExcelのCLSIDを取得
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

    // 恐らく、Excelがインストールされてない場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get CLSID for Excel", L"Error", MB_OK);
        return;
    }

    // 2. 既存のExcelインスタンスを取得
    CComPtr<::IUnknown> pUnk;
    hr = GetActiveObject(clsid, nullptr, reinterpret_cast<::IUnknown**>(&pUnk));  // まずIUnknownを取得

    // 起動中のExcelがない場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get active Excel instance", L"Error", MB_OK);
        return ;
    }

    // 3. IUnknownからIDispatchへのキャスト
    CComPtr<IDispatch> pExcelApp;
    hr = pUnk->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&pExcelApp));

    //キャストに失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get IDispatch from Excel instance", L"Error", MB_OK);
        return ;
    }

    // 4. DISPIDの取得
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Run");  // 実行するメソッド名(VBAのApplication.Run 相当)
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);

    //Runメソッドの取得に失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get DISPID for Run method", L"Error", MB_OK);
        return ;
    }

    // 5. Application.Run メソッドの引数を設定
    CComVariant macroName(ExcelMacroPass);  // 1. 実行したいマクロのフルパス(action要素のarguments属性)

    // 2次元配列とマクロ名を引数として渡す(input要素一式)
    CComVariant saVariant;
    saVariant.vt = VT_ARRAY | VT_BSTR;
    saVariant.parray = UserInputs;

    CComVariant macroArg1(saVariant);      // 2. input要素一式

    // 6. 引数を配列として渡す(※これらの引数は逆の順序で表示されるため、それを考慮した代入を行うこと)
    CComVariant argsArray[2] = { macroArg1,macroName };
    DISPPARAMS params = { argsArray, nullptr, 2, 0 };

    // 7. マクロの呼び出し
    CComVariant result;
    hr = pExcelApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &excepInfo, nullptr);

    //// 現在のExcelインスタンス内に、指定マクロがないと想定
    //if (FAILED(hr)) {
    //    MessageBoxW(nullptr, L"Failed to get Excel macro", L"Error", MB_OK);
    //}

    ////-------------以降は、デバッグ用-------------

    ////MessageBoxでDISPPARAMSの内容を確認
    //std::wstring debugMessage;

    //// cArgsの確認
    //debugMessage += L"Number of arguments: " + std::to_wstring(params.cArgs) + L"\n";

    //// rgvarg の中身を文字列化
    //for (UINT i = 0; i < params.cArgs; ++i) {
    //    VARIANT& arg = params.rgvarg[i];

    //    if (arg.vt == VT_BSTR) {
    //        debugMessage += L"Argument " + std::to_wstring(i) + L": " + arg.bstrVal + L"\n";
    //    }
    //    else {
    //        debugMessage += L"Argument " + std::to_wstring(i) + L": [not a BSTR]\n";
    //    }
    //}

    //// rgvarg の中身を確認
    //MessageBoxW(nullptr, debugMessage.c_str(), L"DISPPARAMS Debug", MB_OK);

    ////エラーが起こったら、エラーコードと詳細メッセージ(ある場合)を表示。
    //if (FAILED(hr)) {
    //    std::wstring errorMessage = L"Invoke failed. HRESULT: " + std::to_wstring(hr);

    //    if (excepInfo.bstrDescription) {
    //        errorMessage += L"\nException: " + std::wstring(excepInfo.bstrDescription);
    //        SysFreeString(excepInfo.bstrDescription);  // リソース解放
    //    }

    //    MessageBoxW(nullptr, errorMessage.c_str(), L"Error1", MB_OK);
    //}
    //else {
    //    _com_error err(hr);
    //    MessageBoxW(nullptr, err.ErrorMessage(), L"Info", MB_OK);
    //}
}

//***************************************************************************************************
//* 機能　　：トースト通知のアクティベーションを処理する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：※割愛します 
//---------------------------------------------------------------------------------------------------
//* URL     ：・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastactivatedeventargs.arguments
//            ・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.activated
//***************************************************************************************************
void OnActivated(ToastNotification const& sender, IInspectable const& args) {
    // IInspectable から ToastActivatedEventArgs にキャストして引数情報を取得
    auto activatedArgs = args.try_as<ToastActivatedEventArgs>();

    // UserInput()からすべてのキーと値のペアを取得
    auto userInputs = activatedArgs.UserInput();

    //トーストからの args.try_as<ToastActivatedEventArgs> があれば、Excelマクロを動かす準備に入る
    if (activatedArgs) {
        try {
            //ボタン押下したAction要素のarguments属性の内容を取得(マクロ名を想定)
            winrt::hstring argument = activatedArgs.Arguments();

            // Input要素のIDと値を格納するための2次元配列を作成する準備(SAFEARRAY)
            long inputCount = userInputs.Size();  // 入力フィールドの数を取得
            SAFEARRAYBOUND bounds[2];                            // 2次元配列として設定
            bounds[0].lLbound = 0;                               // 行数-最小要素番号
            bounds[0].cElements = inputCount;                    // 行数-最大要素番号 (入力フィールドの数)
            bounds[1].lLbound = 0;                               // 列数-最小要素番号
            bounds[1].cElements = 2;                             // 列数-最大要素番号 (キーと値のペア)

            // 上記の設定を基に、2次元配列を作成
            SAFEARRAY* InputElementsArray = SafeArrayCreate(VT_BSTR, 2, bounds);

            long indices[2];
            long rowIndex = 0;
            for (auto const& input : userInputs) {
                auto key = input.Key();                // 入力フィールドのID (キー)
                auto value = input.Value();            // 入力された値 (IInspectable型)
                auto inputValue = value.as<winrt::hstring>();

                // 配列にキーを追加する準備
                indices[0] = rowIndex;  //現時点のInput要素位置
                indices[1] = 0;  // キーは0列目に
                CComBSTR bstrKey(key.c_str()); //Input要素のID属性を取得
                
                //上記の設定で配列にキーを追加
                SafeArrayPutElement(InputElementsArray, indices, bstrKey);

                // 配列に値を追加する準備
                indices[1] = 1;  // 値は1列目に
                CComBSTR bstrValue(inputValue.c_str());//Input要素の値を取得

                //上記の設定で配列にキーを追加
                SafeArrayPutElement(InputElementsArray, indices, bstrValue);

                rowIndex++;
            }

            //トーストのaction要素にあるarguments属性の値(マクロ名)と、Input要素一式をExcelマクロ処理用に渡す
            ExecuteExcelMacro(argument.c_str(), InputElementsArray);
        }
        catch (const hresult_error& e)
        {
            // エラーハンドリング
            MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        }
    }
}

//***************************************************************************************************
//* 機能　　：コレクション(CollectionToast)を使用したトースト通知の表示を行います
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//---------------------------------------------------------------------------------------------------
//* 機能説明：非同期処理をラップするヘルパー関数。
//            処理の流れは、ShowToastNotificationとほとんど一緒ですが、非同期処理が必要のため、このようなラッパー用関数を作成しています。
//***************************************************************************************************
winrt::fire_and_forget SendToastWithCollectionAsyncHelper(ToastNotificationParams* ToastConfigData)
{
    try
    {
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

        // ToastNotifierForToastCollectionIdAsyncを使って特定のコレクションのNotifierを非同期で取得
        ToastNotifier notifier = co_await ToastNotificationManager::GetDefault().GetToastNotifierForToastCollectionIdAsync(ToastConfigData->CollectionID);

        //スケジュールが指定されてあったらその処理を行う
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
            notifier.AddToSchedule(scheduledToast);
        }
        else {

            // 通常のトースト通知を作成
            ToastNotification toast(toastXml);

            // イベントハンドラーを設定
            toast.Activated(TypedEventHandler<ToastNotification, IInspectable>(OnActivated)); // OnActivated関数をハンドラーとして設定

            // 上記で作成されたオブジェクトに各種設定(GroupとTag等)を施す
            toast.ExpiresOnReboot(ToastConfigData->ExpiresOnReboot);
            toast.Group(ToastConfigData->Group);
            toast.Tag(ToastConfigData->Tag);
            toast.SuppressPopup(ToastConfigData->SuppressPopup);
            if (ToastConfigData->ExpirationTime > 0) toast.ExpirationTime(ExpirationTimeValue);

            // Collection経由によるトーストを表示
            notifier.Show(toast);
        }
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }
}



//***************************************************************************************************
//                              ■■■ VBA側から呼び出せる関数 ■■■
//***************************************************************************************************
//* 機能　　：単純なトースト通知を表示します。指定日時に通知するスケジュール機能も対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//***************************************************************************************************
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
        //MessageBoxW(nullptr, ToastConfigData->CollectionID, L"CollectionID", MB_OK);

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
       
        //1. CollectionIDが指定されてあったら、専用の処理に移る
        if (ToastConfigData->CollectionID) {
            // 非同期処理の呼び出し
            SendToastWithCollectionAsyncHelper(ToastConfigData);
        }
        
        //2. スケジュール通知、指定時
        else if (ToastConfigData->Schedule_DeliveryTime > 0) {
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

            // イベントハンドラーを設定
            toast.Activated(TypedEventHandler<ToastNotification, IInspectable>(OnActivated)); // OnActivated関数をハンドラーとして設定

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

//***************************************************************************************************
//* 機能　　：引数に渡された値で、最初のトーストの進行状況バーを表示します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData                ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//            ProgressStatus                 ステータス表記
//            ProgressValue                  0.0-1.0の値。プログレスバーの表示割合を決めます 
//          　ProgressTitle                  タイトル
//            ProgressValueStringOverride    %表記から変えたいとき用(0/88等)                       
//***************************************************************************************************
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
        //if (ToastConfigData->ExpirationTime > 0) toastWithProgress.ExpirationTime(ExpirationTimeValue);

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

//***************************************************************************************************
//* 機能　　：引数に渡された値で、トーストの進行状況バーを更新します。
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData                ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//            ProgressStatus                 ステータス表記
//            ProgressValue                  0.0-1.0の値。プログレスバーの表示割合を決めます 
//          　ProgressTitle                  タイトル
//            ProgressValueStringOverride    %表記から変えたいとき用(0/88等)                       
//          　SequenceNumber                 整合性用       
//***************************************************************************************************
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

//***************************************************************************************************
//* 機能　　：トースト通知を削除します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//---------------------------------------------------------------------------------------------------
//* 機能説明：最も細かく設定できる引数、tag,group,appid　の3つで削除を指定します。
//***************************************************************************************************
void __stdcall RemoveToastNotification(ToastNotificationParams* ToastConfigData) {
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

    try{

        //値Check用
        //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Tag, L"Tag", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Group, L"Group", MB_OK);

        // ToastNotificationManagerからToastNotificationHistory.Remove メソッドを使用し、該当の通知を削除する
        ToastNotificationManager::History().Remove(ToastConfigData->Tag, ToastConfigData->Group, ToastConfigData->AppUserModelID);
    }
    catch (const hresult_error& ex){
        // エラー処理: 必要に応じてエラーメッセージを表示
        MessageBox(nullptr, ex.message().c_str(), L"Error", MB_OK);
    }
}

//***************************************************************************************************
//* 機能　　：引数に渡された値から、コレクションを使用したトースト通知のグループ化を作成します。エラーコード返却に対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData   ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//            displayName       コレクション名
//            launchArgs        起動引数
//            iconUri           アイコンパス
//***************************************************************************************************
long __stdcall CreateToastCollection(ToastNotificationParams* ToastConfigData, const wchar_t* displayName, const wchar_t* launchArgs, const wchar_t* iconUri) {
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

    //値Check用
    //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
    //MessageBoxW(nullptr, ToastConfigData->CollectionID, L"collectionId", MB_OK);
    //MessageBoxW(nullptr, displayName, L"displayName", MB_OK);
    //MessageBoxW(nullptr, launchArgs, L"launchArgs", MB_OK);
    //MessageBoxW(nullptr, iconUri, L"iconUri", MB_OK);

    try {
        // トースト通知のマネージャーを取得
        ToastNotificationManagerForUser userManager = ToastNotificationManager::GetDefault();
        ToastCollectionManager collectionManager = userManager.GetToastCollectionManager(ToastConfigData->AppUserModelID);

        //iconUriから、Uri クラスのインスタンスを作成します
        Uri siteUri = Uri(iconUri);

        // コレクションを作成
        ToastCollection collection = ToastCollection(ToastConfigData->CollectionID, displayName, launchArgs, siteUri);
        collectionManager.SaveToastCollectionAsync(collection);

        // 成功したら0を返す
        return 0;
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング (エラーコードを返す)
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        return e.code();
    }         

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }

}

//***************************************************************************************************
//* 機能　　：コレクションを使用したトースト通知のグループ化を削除します。エラーコード返却に対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//***************************************************************************************************
long __stdcall DeleteToastCollection(ToastNotificationParams* ToastConfigData) {
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

    try {
        // トースト通知のマネージャーを取得
        ToastNotificationManagerForUser userManager = ToastNotificationManager::GetDefault();
        ToastCollectionManager collectionManager = userManager.GetToastCollectionManager(ToastConfigData->AppUserModelID);

        //CollectionIDの定義Check
        if (ToastConfigData->CollectionID) {
            //何かのCollectionIDが指定してあったら、それのみ削除
            collectionManager.RemoveToastCollectionAsync(ToastConfigData->CollectionID);
        }
        else {
            //未定義の場合、全てのToastCollectionを削除
            collectionManager.RemoveAllToastCollectionsAsync();
        }

        // 成功したら0を返す
        return 0;
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング (エラーコードを返す)
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        return e.code();
    }

    // CoUninitialize()は、CoInitializeExが成功した場合のみ呼び出す
    if (SUCCEEDED(hr)) {
        CoUninitialize();
    }

}
