/***************************************************************************************************
 *										    イベント機能
 ***************************************************************************************************
 * 以下のイベントを記述します
 * ・Activated イベント
 *      通知へのアクション(ボタン押下など)で、指定VBAマクロの実行
 * 　   →入力値、ドロップダウンリストで選択したIDも取得可能
 *
 *
 * URL
 * https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification#events
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"



//***************************************************************************************************
//                                  ■■■ 静的定数 ■■■
//***************************************************************************************************
constexpr const wchar_t* EXCEL_DESK_CLASS_NAME = L"XLDESK";                 //"XLMAIN"ウィンドウの子名称
constexpr const wchar_t* EXCEL_SHEET_CLASS_NAME = L"EXCEL7";                //"XLDESK"の子名称
constexpr const wchar_t* EXCEL_APPLICATION_CLASS_NAME = L"Application";     //"Application"のオブジェクト名称
constexpr const wchar_t* EXCEL_APPLICATION_RUN_MethodName = L"Run";         //"Application.Run"のメソッド名称

constexpr const wchar_t* EventTriggerMacroName_ToastDismissed = L"ExcelToast_Dismissed";     //トースト通知のDismissed イベント時に使うプロシージャ名



//***************************************************************************************************
//                                  ■■■ ヘルパー用関数 ■■■
//***************************************************************************************************
//* 機能　　：引数に従った Application オブジェクトを取得します
//---------------------------------------------------------------------------------------------------
//* 引数　　：※割愛します
//---------------------------------------------------------------------------------------------------
//* 詳細説明：WorkbookからApplicationを取得するために使います
//***************************************************************************************************
static HRESULT GetProperty(IDispatch* pDisp, const wchar_t* propName, CComVariant& result) {
    if (!pDisp) return E_POINTER;
    OLECHAR* name = (OLECHAR*)propName;
    DISPID dispID;
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) return hr;
    DISPPARAMS params = { NULL, NULL, 0, 0 };
    return pDisp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &result, NULL, NULL);
}



//***************************************************************************************************
//                                  ■■■ メイン関数 ■■■
//***************************************************************************************************
//* 機能　　：Excel マクロを実行する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：ExcelMacroPass    Action要素のarguments。{マクロ名}-{ExcelVBA からの Application.hwnd} を想定してます。
//            UserInputs        Input要素で入力した内容、あるいはSelect要素のID名称とそれに紐づくInput要素のIDとのセットとなる2次元配列                             
//            targetHWND        プロシージャ起動先のExcel ハンドル値。ToastNotification.Group プロパティ から得る設計にしてます
//---------------------------------------------------------------------------------------------------
//* 詳細説明：ExcelのHWNDも渡すことで、複数プロセスで起動してるExcel環境でも対応できます
//***************************************************************************************************
static void ExecuteExcelMacro(const wchar_t* ExcelMacroPass, SAFEARRAY* UserInputs, HWND targetHWND) {
    //---------- 1. 孫ウィンドウ経由で、Excel Applicationオブジェクトを取得 ----------
    CComPtr<IDispatch> pExcelDispatch;
    HRESULT hr = E_FAIL; // 見つからなかった場合のデフォルト

    // 1-1. XLMAINウィンドウの子である「XLDESK」ウィンドウを探す
    HWND hXlDesk = FindWindowExW(targetHWND, NULL, EXCEL_DESK_CLASS_NAME, NULL);
    if (hXlDesk) {
        // 1-2. XLDESKの子である「EXCEL7」ウィンドウを探す
        HWND hExcel7 = FindWindowExW(hXlDesk, NULL, EXCEL_SHEET_CLASS_NAME, NULL);
        if (hExcel7) {
            // 1-3. EXCEL7ウィンドウから直接Workbookオブジェクトを取得
            CComPtr<IDispatch> pWorkbookDisp;
            hr = AccessibleObjectFromWindow(hExcel7, OBJID_NATIVEOM, IID_IDispatch, (void**)&pWorkbookDisp);

            if (SUCCEEDED(hr) && pWorkbookDisp) {
                // 1-4. WorkbookオブジェクトからApplicationオブジェクトを取得
                CComVariant varApp;
                hr = GetProperty(pWorkbookDisp, EXCEL_APPLICATION_CLASS_NAME, varApp);
                if (SUCCEEDED(hr) && varApp.vt == VT_DISPATCH) {
                    pExcelDispatch = varApp.pdispVal; // 成功！
                }
            }
        }
        else {
            hr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
    }
    else {
        hr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    // --- ここまで ---

    if (SUCCEEDED(hr) && pExcelDispatch) {
        //成功！
    }
    else {
        _com_error err(hr);
        wchar_t buf[512];
        const wchar_t* reason = L"不明なエラー";
        if (!hXlDesk) reason = L"子ウィンドウ 'XLDESK' が見つかりません";
        else if (!FindWindowExW(hXlDesk, NULL, L"EXCEL7", NULL)) reason = L"孫ウィンドウ 'EXCEL7' が見つかりません";
        else reason = L"EXCEL7からオブジェクト取得に失敗しました";

        swprintf_s(buf, L"エラー理由: %s\nHRESULT=0x%08X\n%s", reason, hr, err.ErrorMessage());
        MessageBoxW(nullptr, buf, L"エラー", MB_OK);

        return;
    }

    // 2. DISPIDの取得
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(EXCEL_APPLICATION_RUN_MethodName);  // 実行するメソッド名(VBAのApplication.Run 相当)
    hr = pExcelDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);

    //　Runメソッドの取得に失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get DISPID for Run method", L"Error", MB_OK);
        return;
    }

    //---------- 3. Application.Run メソッドの引数を設定 ---------- 
    // 3-1. 実行したいマクロのフルパス(action要素のarguments属性値 ortoast要素のlaunch属性値 を想定)
    CComVariant macroName(ExcelMacroPass);

    // 3-2. 2次元配列(input要素一式)とマクロ名を引数として渡す設定をする
    CComVariant saVariant;
    saVariant.vt = VT_ARRAY | VT_BSTR;
    saVariant.parray = UserInputs;      //input要素一式 2次元配列
    CComVariant macroArg1(saVariant);   //適用      

    // 4. 引数を配列として渡す(※これらの引数は逆の順序で表示されるため、それを考慮した代入を行うこと)
    CComVariant argsArray[2] = { macroArg1,macroName };
    DISPPARAMS params = { argsArray, nullptr, 2, 0 };

    // 5. マクロの呼び出し
    //　デバッグ詳細メッセージ、取得用
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(EXCEPINFO));  // 初期化

    CComVariant result;
    hr = pExcelDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &excepInfo, nullptr);


    ////-------------以降は、デバッグ用-------------
    //// 現在のExcelインスタンス内に、指定マクロがないと想定
    //if (FAILED(hr)) {
    //    MessageBoxW(nullptr, L"Failed to get Excel macro", L"Error", MB_OK);
    //}

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
//* 引数　　：sender     通知オブジェクト
//            args       Input要素にて入力したパラメーター一式と、発動元のAction要素のarguments属性値あるいは、toast要素のlaunch属性値
//---------------------------------------------------------------------------------------------------
//* URL     ：・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastactivatedeventargs.arguments
//            ・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.activated
//            ・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationactiontriggerdetail
//---------------------------------------------------------------------------------------------------
//* 注意事項：残念ながら、Headerクリックによるトリガーはできません
//***************************************************************************************************
void OnActivated(ToastNotification const& sender, IInspectable const& args) {
    // IInspectable から ToastActivatedEventArgs にキャストして引数情報を取得
    auto activatedArgs = args.try_as<ToastActivatedEventArgs>();

    //トーストからの args.try_as<ToastActivatedEventArgs> があれば、Excelマクロを動かす準備に入る
    if (activatedArgs) {
        try {
            // GroupプロパティからHWNDを取得(VBA側で必ず設定すること)
            winrt::hstring group = sender.Group();
            HWND targetHwnd = (HWND)std::stoull(group.c_str());

            // UserInput()からすべてのキーと値のペアを取得(ここに、Input要素にて入力したパラメーター一式があります)
            auto userInputs = activatedArgs.UserInput();

            // ボタン押下によるAction要素のarguments属性値あるいは、toast要素のlaunch属性値の内容を取得
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

            //Action要素のarguments属性値あるいは、toast要素のlaunch属性値の内容と、Input要素一式、Groupプロパティから得たHWNDをExcelマクロ処理用に渡す
            ExecuteExcelMacro(argument.c_str(), InputElementsArray, targetHwnd);
        }
        catch (const hresult_error& e)
        {
            // エラーハンドリング
            MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        }
    }
}

//***************************************************************************************************
//* 機能　　：トースト通知のDismissedを処理する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：sender     通知オブジェクト
//            args       閉じられた理由などを含むイベント引数
//---------------------------------------------------------------------------------------------------
//* URL     ：・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.dismissed
//            ・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastdismissalreason
//***************************************************************************************************
void OnDismissed(ToastNotification const& sender, ToastDismissedEventArgs const& args){
    try {
        // GroupプロパティからHWNDを取得(VBA側で必ず設定すること)
        winrt::hstring group = sender.Group();
        HWND targetHwnd = (HWND)std::stoull(group.c_str());

        // 閉じられた理由を取得
        ToastDismissalReason reason = args.Reason();
        long reasonValue = static_cast<long>(reason);

        //----- デバッグ用 -----
        //std::wstring reasonName;
        //switch (reason) {
        //    case ToastDismissalReason::UserCanceled: reasonName = L"UserCanceled"; break;
        //    case ToastDismissalReason::ApplicationHidden: reasonName = L"ApplicationHidden"; break;
        //    case ToastDismissalReason::TimedOut: reasonName = L"TimedOut"; break;
        //    default: reasonName = L"Unknown"; break;
        //}

        //std::wstring message = L"通知が閉じられました。\n理由: " + reasonName;
        //MessageBoxW(nullptr, message.c_str(), L"Dismissedイベント発生", MB_OK);

        // 閉じられた理由を格納するSAFEARRAYを準備
        SAFEARRAYBOUND bounds[2];long indices[2];
        bounds[0].lLbound = 0;
        bounds[0].cElements = 1; // 1行
        bounds[1].lLbound = 0;
        bounds[1].cElements = 2; // 2列 (Tag, 理由値)
        SAFEARRAY* dismissedInfoArray = SafeArrayCreate(VT_BSTR, 2, bounds);

        // 1列目にTagプロパティを設定
        indices[0] = 0; indices[1] = 0;
        CComBSTR bstrReasonName(sender.Tag().c_str());
        SafeArrayPutElement(dismissedInfoArray, indices, bstrReasonName);

        // 2列目に理由値を設定
        indices[0] = 0; indices[1] = 1;
        CComBSTR bstrReasonValue(std::to_wstring(reasonValue).c_str());
        SafeArrayPutElement(dismissedInfoArray, indices, bstrReasonValue);

        //決められたプロシージャ名、閉じられた理由情報の2次元配列、Groupプロパティから得たHWNDをExcelマクロ処理用に渡す
        ExecuteExcelMacro(EventTriggerMacroName_ToastDismissed, dismissedInfoArray, targetHwnd);
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }
}
