/***************************************************************************************************
 *										    イベント機能
 ***************************************************************************************************
 * 以下のイベントを記述します
 * ・Activated イベント
 *      通知へのアクション(ボタン押下など)で、指定VBAマクロの実行
 * 　   →入力値、ドロップダウンリストで選択したIDも取得可能
 *
 * ・Dismissed イベント
 * ・Failed イベント
 *
 * URL
 * https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification#events
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"



//***************************************************************************************************
//                                  ■■■ 静的定数 ■■■
//***************************************************************************************************
constexpr const wchar_t* EXCEL_MAIN_CLASS_NAME = L"XLMAIN";                 //"XLMAIN"ウィンドウの名称
constexpr const wchar_t* EXCEL_DESK_CLASS_NAME = L"XLDESK";                 //"XLMAIN"ウィンドウの子名称
constexpr const wchar_t* EXCEL_SHEET_CLASS_NAME = L"EXCEL7";                //"XLDESK"の子名称
constexpr const wchar_t* EXCEL_APPLICATION_CLASS_NAME = L"Application";     //"Application"のオブジェクト名称
constexpr const wchar_t* EXCEL_APPLICATION_RUN_MethodName = L"Run";         //"Application.Run"のメソッド名称

constexpr const wchar_t* EventTriggerMacroName_ToastDismissed = L"ExcelToast_Dismissed";    //トースト通知の Dismissed イベント時に使うプロシージャ名
constexpr const wchar_t* EventTriggerMacroName_ToastFailed = L"ExcelToast_Failed";          //トースト通知の Failed イベント時に使うプロシージャ名

// EnumThreadWindowsのためのコールバック関数
struct EnumThreadWndData {
    HWND foundHwnd;
};


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
//* 機能　　：EnumThreadWindows 用のコールバック関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：※割愛します
//---------------------------------------------------------------------------------------------------
//* 注意事項：Excel の hwnd 取得に限ります。
//***************************************************************************************************
BOOL CALLBACK EnumThreadWndProc(HWND hwnd, LPARAM lParam) {
    auto& data = *reinterpret_cast<EnumThreadWndData*>(lParam);
    wchar_t className[256];
    if (GetClassNameW(hwnd, className, 256) > 0) {
        if (wcscmp(className, EXCEL_MAIN_CLASS_NAME) == 0) {
            data.foundHwnd = hwnd;
            return FALSE; // 発見したので終了
        }
    }
    return TRUE;
}

//***************************************************************************************************
//* 機能　　：PIDからHWNDを見つける関数です
//---------------------------------------------------------------------------------------------------
//* 返り値  ：PID に基づく、Excel の HWND
//* 引数　　：取得したプロセスID
//---------------------------------------------------------------------------------------------------
//* 注意事項：「#include <tlhelp32.h>」のインクルードが必要です
//***************************************************************************************************
HWND FindMainWindowFromPid(DWORD pid) {
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
    if (hSnap == INVALID_HANDLE_VALUE) return NULL;

    THREADENTRY32 te;
    te.dwSize = sizeof(te);

    HWND foundHwnd = NULL;

    if (Thread32First(hSnap, &te)) {
        do {
            if (te.th32OwnerProcessID == pid) {
                // PIDが一致するスレッドを発見
                EnumThreadWndData data = { 0 };
                EnumThreadWindows(te.th32ThreadID, EnumThreadWndProc, reinterpret_cast<LPARAM>(&data));
                if (data.foundHwnd) {
                    foundHwnd = data.foundHwnd;
                    break; // 見つかったのでループを抜ける
                }
            }
        } while (Thread32Next(hSnap, &te));
    }

    CloseHandle(hSnap);
    return foundHwnd;
}

//***************************************************************************************************
//* 機能　　：トーストオブジェクトにある Group から、埋め込まれてるであろう PID を抽出し、そこから HWND を抽出します
//---------------------------------------------------------------------------------------------------
//* 返り値  ：ブック名と、PID に基づくHWND
//* 引数　　：トーストオブジェクト
//***************************************************************************************************
std::pair<std::wstring, HWND> GetBookInfoFromGroup(ToastNotification const& Target) {
    // 1. Groupプロパティから、複合文字列を取得
    winrt::hstring groupHString = Target.Group();
    std::wstring combinedString = groupHString.c_str();

    // 2. rfind()で、最後の '|' の位置を探す
    size_t lastPipePos = combinedString.rfind(L'|');

    // 3. '|' が見つかったかどうかをチェック
    if (lastPipePos != std::wstring::npos) {
        // 見つかった場合

        // 4-1. substr()で、'|' の次の文字から最後までを切り出す
        std::wstring pidString = combinedString.substr(lastPipePos + 1);

        // 4-2. ブック名の部分も取得
        std::wstring bookNameString = combinedString.substr(0, lastPipePos);

        // デバッグ用に表示
        //MessageBoxW(nullptr, pidString.c_str(), L"抽出したPID", MB_OK);
        //MessageBoxW(nullptr, bookNameString.c_str(), L"抽出したブック名", MB_OK);

        // 5. 抽出した文字列を数値に変換
        DWORD targetPID = std::stoul(pidString);

        // 6. ブック名と見つけたHWNDをペアにして返す
        return { bookNameString, FindMainWindowFromPid(targetPID) };

    }
    else {
        // '|' が見つからなかった場合のエラー処理
        MessageBoxW(nullptr, L"Groupプロパティの形式が正しくありません。(区切り文字'|'が見つかりません)", L"エラー", MB_OK);
        return { L"", NULL };
    }
}



//***************************************************************************************************
//                                  ■■■ メイン関数 ■■■
//***************************************************************************************************
//* 機能　　：Excel マクロを実行する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：ExcelMacroPass    '{ブック名}'!{マクロ名}を想定してます。
//            UserInputs        Input要素で入力した内容、あるいはSelect要素のID名称とそれに紐づくInput要素のIDとのセットとなる2次元配列                             
//            targetHWND        プロシージャ起動先のExcel ハンドル値。ToastNotification.Group プロパティ から得る設計にしてます
//---------------------------------------------------------------------------------------------------
//* 詳細説明：ExcelのHWNDを渡すことで、複数プロセスで起動してるExcel環境でも対応できます
//***************************************************************************************************
static void ExecuteExcelMacro(const wchar_t* ExcelMacroPass, IDispatch* UserInputs, HWND targetHWND) {
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
    saVariant.vt = VT_DISPATCH;
    saVariant.pdispVal = UserInputs;      //input要素一式 2次元配列
    UserInputs->AddRef();                 // Dictionaryオブジェクトの参照カウントを増やしておく

    //---------- 4. 引数を配列として渡す ----------   
    //引数は逆の順序で表示されるため、それを考慮した代入を行う
    CComVariant argsArray[2] = { UserInputs,macroName };
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
            //---------- 1. 実行先プロシージャを特定する準備 ----------
            // 1-1. GroupプロパティからHWNDとBook名を取得(VBA側で必ずブック名とPIDを設定すること)
            auto [bookName, targetHwnd] = GetBookInfoFromGroup(sender); // ★構造化束縛(C++17↑)

            // 1-2. ボタン押下によるAction要素のarguments属性値あるいは、toast要素のlaunch属性値の内容を取得(マクロ名を想定)
            winrt::hstring argument = activatedArgs.Arguments();

            // 1-3. 完全修飾マクロ名を組み立てる
            std::wstring qualifiedMacroName = L"'" + bookName + L"'!" + argument.c_str();


            //---------- 2.Input要素のIDと値を格納するためのDictionaryを作成する準備 ----------    
            // UserInput()からすべてのキーと値のペアを取得(ここに、Input要素にて入力したパラメーター一式があります)
            auto userInputs = activatedArgs.UserInput();

            // 変数準備
            CComPtr<IDispatch> pDictionary;
            CLSID clsid;

            // 2-1. "Scripting.Dictionary"のCLSIDを取得
            HRESULT hr = CLSIDFromProgID(L"Scripting.Dictionary", &clsid);
            if (SUCCEEDED(hr)) {
                // 2-2. 空のDictionaryオブジェクトを生成！
                hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pDictionary);
            }
            if (pDictionary) {
                // 2-3. ".Add" メソッドのIDを取得
                OLECHAR* addMethodName = const_cast<OLECHAR*>(L"Add");
                DISPID dispidAdd;
                hr = pDictionary->GetIDsOfNames(IID_NULL, &addMethodName, 1, LOCALE_USER_DEFAULT, &dispidAdd);

                if (SUCCEEDED(hr)) {
                    for (auto const& input : userInputs) {
                        auto key = input.Key();                // 入力フィールドのID (キー)
                        auto value = input.Value();            // 入力された値 (IInspectable型)
                        auto inputValue = value.as<winrt::hstring>();  //扱いやすいように変換

                        // 2-4. キーと値を準備
                        CComBSTR bstrKey(key.c_str());          //Input要素のID属性を取得
                        CComBSTR bstrValue(inputValue.c_str()); //Input要素の値または、Select要素のIDを取得

                        // 2-5. Invokeのための引数配列を作成（逆順！）
                        CComVariant addArgs[2] = { bstrValue, bstrKey };
                        DISPPARAMS params = { addArgs, NULL, 2, 0 };

                        // 2-6. .Addメソッドを呼び出して、データを詰める！
                        pDictionary->Invoke(dispidAdd, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
                    }
                }
            }

            //Action要素のarguments属性値あるいは、toast要素のlaunch属性値の内容と、Input要素一式、Groupプロパティから得たHWNDとブック名を基に、Excelマクロ処理用に渡す
            ExecuteExcelMacro(qualifiedMacroName.c_str(), pDictionary, targetHwnd);
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
        //---------- 1. 実行先プロシージャを特定する準備 ----------
        // 1-1. GroupプロパティからHWNDとBook名を取得(VBA側で必ずブック名とPIDを設定すること)
        auto [bookName, targetHwnd] = GetBookInfoFromGroup(sender); // ★構造化束縛(C++17↑)

        // 1-2. 完全修飾マクロ名を組み立てる
        std::wstring qualifiedMacroName = L"'" + bookName + L"'!" + EventTriggerMacroName_ToastDismissed;


        //---------- 2.閉じられた理由を取得を作成する準備(SAFEARRAY) ----------    
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
        //ExecuteExcelMacro(qualifiedMacroName.c_str(), dismissedInfoArray, targetHwnd);
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }
}

//***************************************************************************************************
//* 機能　　：トースト通知のFailedを処理する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：sender     通知オブジェクト
//            args       エラー関係の引数
//---------------------------------------------------------------------------------------------------
//* URL     ：・https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.failed
//***************************************************************************************************
void OnFailed(ToastNotification const& sender, ToastFailedEventArgs const& args){
    try {
        //---------- 1. 実行先プロシージャを特定する準備 ----------
        // 1-1. GroupプロパティからHWNDとBook名を取得(VBA側で必ずブック名とPIDを設定すること)
        auto [bookName, targetHwnd] = GetBookInfoFromGroup(sender); // ★構造化束縛(C++17↑)

        // 1-2. 完全修飾マクロ名を組み立てる
        std::wstring qualifiedMacroName = L"'" + bookName + L"'!" + EventTriggerMacroName_ToastFailed;


        //---------- 2.なぜ失敗したのか、HRESULT形式のエラーコードを取得を作成する準備(SAFEARRAY) ----------    
        winrt::hresult errorCode = args.ErrorCode();

        // デバッグやロギングのためにエラー情報を文字列化
        _com_error err(errorCode);

        // エラー理由を格納するSAFEARRAYを準備
        SAFEARRAYBOUND bounds[2]; long indices[2];
        bounds[0].lLbound = 0;
        bounds[0].cElements = 1; // 1行
        bounds[1].lLbound = 0;
        bounds[1].cElements = 2; // 2列 (Tag, エラー内容)
        SAFEARRAY* failedInfoArray = SafeArrayCreate(VT_BSTR, 2, bounds);

        // 1列目にTagプロパティを設定
        indices[0] = 0; indices[1] = 0;
        CComBSTR bstrReasonName(sender.Tag().c_str());
        SafeArrayPutElement(failedInfoArray, indices, bstrReasonName);

        // 2列目にエラー内容を設定
        wchar_t hresultStr[20]; // "0x" + 8桁の16進数 + NULL文字
        swprintf_s(hresultStr, L"0x%08X", errorCode);

        std::wstring detailedErrorMessage = hresultStr;
        detailedErrorMessage += L"\n"; // 改行を追加
        detailedErrorMessage += err.ErrorMessage();

        // 組み立てた文字列をBSTRとして設定
        indices[0] = 0; indices[1] = 1;
        CComBSTR bstrErrorDetails(detailedErrorMessage.c_str());
        SafeArrayPutElement(failedInfoArray, indices, bstrErrorDetails);

        //決められたプロシージャ名、エラー情報の2次元配列、Groupプロパティから得たHWNDをExcelマクロ処理用に渡す
        //ExecuteExcelMacro(qualifiedMacroName.c_str(), failedInfoArray, targetHwnd);
    }
    catch (const hresult_error& e)
    {
        // エラーハンドリング
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
    }
}
