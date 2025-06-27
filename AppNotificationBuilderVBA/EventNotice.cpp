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
//* 機能　　：Excel マクロを実行する関数
//---------------------------------------------------------------------------------------------------
//* 引数　　：ExcelMacroPass     Action要素のarguments。[マクロ名-ExcelのPID]を想定してます。
//            UserInputs         Input要素で入力した内容、あるいはSelect要素のID名称とそれに紐づくInput要素のIDとのセットとなる2次元配列                             
//***************************************************************************************************
static void ExecuteExcelMacro(const wchar_t* ExcelMacroPass, SAFEARRAY* UserInputs) {
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
        return;
    }

    // 3. IUnknownからIDispatchへのキャスト
    CComPtr<IDispatch> pExcelApp;
    hr = pUnk->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&pExcelApp));

    //キャストに失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get IDispatch from Excel instance", L"Error", MB_OK);
        return;
    }

    // 4. DISPIDの取得
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Run");  // 実行するメソッド名(VBAのApplication.Run 相当)
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);

    //Runメソッドの取得に失敗した場合
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"Failed to get DISPID for Run method", L"Error", MB_OK);
        return;
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
