/***************************************************************************************************
 *										一般的な通知機能
 ***************************************************************************************************
 * 以下の機能を記述してます
 * ・即日通知と削除
 * ・スケジュール通知
 * ・通知内容の更新(プログレスバーの更新処理など)
 *
 * 
 * URL
 * https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-notifications-overview
 ***************************************************************************************************/



 //設定がまとまってるヘッダーファイルを指定
#include "AppNotificationBuilder.h"



//***************************************************************************************************
//                                 ■■■ 内部のヘルパー関数 ■■■
//***************************************************************************************************
//* 機能　　 ：SYSTEMTIMEをDateTimeに変換します
//---------------------------------------------------------------------------------------------------
//* 引数　　 ：SYSTEMTIME
//* 返り値　 ：dateTime
//***************************************************************************************************
static Windows::Foundation::DateTime SystemTimeToDateTime(const SYSTEMTIME& st) {
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
//* 機能　　：トーストのデータバインディング設定定義を行います
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastUpdata    データバインディング情報
//---------------------------------------------------------------------------------------------------
//* 機能説明：通知を表示しながら、データバインディングをサポートするプロパティ値を適用します
//* 注意事項：null pointer だと、エラーになるため、if で存在判定を行うこと。
//---------------------------------------------------------------------------------------------------
//* URL：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-progress-bar?tabs=xml
//***************************************************************************************************
static NotificationData SetDataBinding(ToastNotificationVariable* ToastUpdata) {
    NotificationData VariableParams;
    auto VariableParamsValues = VariableParams.Values();         // 戻り値の型を明示的に指定

    //最上位レベルの AdaptiveText 要素の Text プロパティ
    if (ToastUpdata->TitleText) VariableParamsValues.Insert(L"TopTextTitle", ToastUpdata->TitleText);                   //タイトル
    if (ToastUpdata->ContentsText) VariableParamsValues.Insert(L"TopTextContents", ToastUpdata->ContentsText);          //コンテンツ
    if (ToastUpdata->AttributionText) VariableParamsValues.Insert(L"TopTextAttribution", ToastUpdata->AttributionText); //属性

    //-----AdaptiveProgressのすべての属性-----
    if (ToastUpdata->ProgressTitle) VariableParamsValues.Insert(L"ProgressTitle", ToastUpdata->ProgressTitle);                               //タイトル
    if (ToastUpdata->ProgressStatus) VariableParamsValues.Insert(L"ProgressStatus", ToastUpdata->ProgressStatus);                            //左下の進行状況バーの下に表示される状態文字列
    //  進捗値の場合、負になってたら、ドットアニメーションの不確定式にします。
    if (ToastUpdata->ProgressValue < 0) {
        VariableParamsValues.Insert(L"ProgressValue", L"Indeterminate");                                     //進行状況バーの状態を「不確定」として、設定
    }
    else {
        VariableParamsValues.Insert(L"ProgressValue", std::to_wstring(ToastUpdata->ProgressValue).c_str());  //進行状況バーの状態を設定
    }
    //  文字列がない場合は、バインディング処理しません
    if (ToastUpdata->ProgressValueStringOverride) VariableParamsValues.Insert(L"ProgressValueString", ToastUpdata->ProgressValueStringOverride);           //既定のパーセンテージ文字列の代わりに表示される省略可能な文字列を取得または設定します。 これが指定されていない場合は、"70%" のようなものが表示されます。

    //返却
    return VariableParams;
}



//***************************************************************************************************
//                              ■■■ VBA側から呼び出す関数 ■■■
//***************************************************************************************************
//* 機能　　：単純なトースト通知を表示します。指定日時に通知するスケジュール機能も対応します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData            ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//            ToastNotificationVariable  データバインディング用引数。後述の「UpdateToastNotification」で使用します           
//***************************************************************************************************
void __stdcall ShowToastNotification(ToastNotificationParams* ToastConfigData, ToastNotificationVariable* ToastUpdata) {
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

        //コレクションIDが指定されてあったら、そっちのオブジェクトを使用する
        ToastNotifier toastNotifier =
            ToastConfigData->CollectionID
            ? ToastNotificationManager::GetDefault().GetToastNotifierForToastCollectionIdAsync(ToastConfigData->CollectionID).get()
            : ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID);

        //スケジュール通知モードの場合、この処理に入る(※この場合、データバインディング、アクティベート機能は使えません)
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
            toastNotifier.AddToSchedule(scheduledToast);
        }

        else {
            // 通常のトースト通知を作成
            ToastNotification toast(toastXml);

            //先ほど定義したデータバインディングを適用する
            toast.Data(SetDataBinding(ToastUpdata));

            // イベントハンドラーを設定
            toast.Activated(TypedEventHandler<ToastNotification, IInspectable>(OnActivated));               //Activated イベント
            toast.Dismissed(TypedEventHandler<ToastNotification, ToastDismissedEventArgs>(OnDismissed));    //Dismissed	イベント

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
}

//***************************************************************************************************
//* 機能　　：引数に渡された値で、トーストの内容を更新します。
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData                ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//            ToastNotificationVariable      データバインディング用引数。
//---------------------------------------------------------------------------------------------------
//* 注意事項：データバインディングに対応する箇所のみとなります。
//***************************************************************************************************
long __stdcall UpdateToastNotification(ToastNotificationParams* ToastConfigData, ToastNotificationVariable* ToastUpdata) {
    try {
        //更新結果用の変数を定義
        NotificationUpdateResult result;

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

        //CollectionIDが指定されてあったら、そっちのオブジェクトを使う
        ToastNotifier UpdateToastNotifier =
            ToastConfigData->CollectionID
            ? ToastNotificationManager::GetDefault().GetToastNotifierForToastCollectionIdAsync(ToastConfigData->CollectionID).get()
            : ToastNotificationManager::CreateToastNotifier(ToastConfigData->AppUserModelID);

        // トースト通知を更新
        result = UpdateToastNotifier.Update(SetDataBinding(ToastUpdata), ToastConfigData->Tag, ToastConfigData->Group);

        //結果値を返す
        return static_cast<long>(result);
    }
    catch (const winrt::hresult_error& e) {
        MessageBoxW(nullptr, e.message().c_str(), L"エラー", MB_OK);
        return e.code();
    }
}

//***************************************************************************************************
//* 機能　　：トースト通知を削除します
//---------------------------------------------------------------------------------------------------
//* 引数　　：ToastConfigData    ヘッダーファイルに定義した引数。ここから必要な値を使用する方針です
//---------------------------------------------------------------------------------------------------
//* 機能説明：最も細かく設定できる引数、tag,group,appid　の3つで削除を指定します。
//***************************************************************************************************
void __stdcall RemoveToastNotification(ToastNotificationParams* ToastConfigData) {
    try {

        //値Check用
        //MessageBoxW(nullptr, ToastConfigData->AppUserModelID, L"AppUserModelID", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Tag, L"Tag", MB_OK);
        //MessageBoxW(nullptr, ToastConfigData->Group, L"Group", MB_OK);

        //コレクションIDが指定されてあったら、そっちの処理を行う
        if (ToastConfigData->CollectionID) {
            // コレクションIDに紐づく指定の通知IDを削除
            ToastNotificationManager::GetDefault().GetHistoryForToastCollectionIdAsync(ToastConfigData->CollectionID).get().Remove(ToastConfigData->Tag, ToastConfigData->Group);
        }
        else {
            // ToastNotificationManagerからToastNotificationHistory.Remove メソッドを使用し、該当の通知を削除する
            ToastNotificationManager::History().Remove(ToastConfigData->Tag, ToastConfigData->Group, ToastConfigData->AppUserModelID);
        }
    }
    catch (const hresult_error& ex) {
        // エラー処理: 必要に応じてエラーメッセージを表示
        MessageBox(nullptr, ex.message().c_str(), L"Error", MB_OK);
    }
}
