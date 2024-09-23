VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_AppNotificationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   VBAから、Windows PowerShell を経由して、デスクトップ通知【トースト通知】を表示させます
'   また、DLLファイルを経由した通知手段も提供します。

' MIT License
'
' Copyright (c) 2024 eschamali
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

'---------------------------------------------------------------------------------------------------
'* URL ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-notifications-overview
'* 概要：テキストはもちろん、アイコンやサムネイル画像の設定ができる通知を提供します。
'        また、進捗状況バーの表示にも対応させます。(参考URL：https://qiita.com/relu/items/b7121487a1d5756dfcf9)
'***************************************************************************************************
Option Explicit

'***************************************************************************************************
'■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
' 機能     ：C++で書かれたDLLに、 Windows.UI.Notifications 関連の処理を埋め込ませ、Shell経由より高速に処理できます
'---------------------------------------------------------------------------------------------------
Private Declare PtrSafe Sub ShowToastNotification Lib "AppNotificationBuilderVBA.dll" (ByRef Params_String As ToastNotificationParams)  '一般的な通知
Private Declare PtrSafe Sub ShowToastNotificationWithProgressBar Lib "AppNotificationBuilderVBA.dll" (ByRef Params_String As ToastNotificationParams, ByVal ProgressStatus As LongPtr, ByVal ProgressValue As Double, ByVal ProgressTitle As LongPtr, ByVal ProgressValueStringOverride As LongPtr) 'プログレスバー付き通知
Private Declare PtrSafe Function UpdateToastNotificationWithProgressBar Lib "AppNotificationBuilderVBA.dll" (ByRef Params_String As ToastNotificationParams, ByVal ProgressStatus As LongPtr, ByVal ProgressValue As Double, ByVal ProgressTitle As LongPtr, ByVal ProgressValueStringOverride As LongPtr, SequenceNumber As Long) As Long 'プログレスバー付き通知の更新用
Private Declare PtrSafe Function CreateToastCollection Lib "AppNotificationBuilderVBA.dll" (ByRef Params_String As ToastNotificationParams, ByVal DisplayName As LongPtr, ByVal LaunchArgs As LongPtr, ByVal IconUri As LongPtr) As Long 'ToastCollectionの作成
Private Declare PtrSafe Function DeleteToastCollection Lib "AppNotificationBuilderVBA.dll" (ByRef Params_String As ToastNotificationParams) As Long 'ToastCollectionの削除



'***************************************************************************************************
'■■■ よく使うWin32API宣言セクション ■■■
'***************************************************************************************************
'  機能     ：Windows API：mciSendString  を使って任意のパスやURLの音源ファイルを再生させる機能です
'---------------------------------------------------------------------------------------------------
Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long



'***************************************************************************************************
'■■■ 一部の設定プロパティ値を列挙型で宣言 ■■■
'***************************************************************************************************
'トーストが表すシナリオを指定します。数字の代わりにこの名称でも機能します。
'---------------------------------------------------------------------------------------------------
Public Enum ToastScenario
    Default
    Reminder
    Alarm
    IncomingCall
    Urgent
End Enum

'===================================================================================================
'トースト通知を表示するときに再生するサウンドを指定します。数字の代わりにこの名称でも機能します。
'---------------------------------------------------------------------------------------------------
Public Enum MsWinSoundEvent
    NotificationDefault = 1
    NotificationIM
    NotificationMail
    NotificationReminder
    NotificationSMS
    NotificationLoopingAlarm01
    NotificationLoopingAlarm02
    NotificationLoopingAlarm03
    NotificationLoopingAlarm04
    NotificationLoopingAlarm05
    NotificationLoopingAlarm06
    NotificationLoopingAlarm07
    NotificationLoopingAlarm08
    NotificationLoopingAlarm09
    NotificationLoopingAlarm10
    NotificationLoopingCall01
    NotificationLoopingCall02
    NotificationLoopingCall03
    NotificationLoopingCall04
    NotificationLoopingCall05
    NotificationLoopingCall06
    NotificationLoopingCall07
    NotificationLoopingCall08
    NotificationLoopingCall09
    NotificationLoopingCall10
End Enum

'===================================================================================================
'ボタンのスタイルを指定します。数字の代わりにこの名称でも機能します。
'---------------------------------------------------------------------------------------------------
Public Enum ToastButtonStyle
    NoStyle
    Success
    Critical
End Enum

'---------------------------------------------------------------------------------------------------



'***************************************************************************************************
' ■■■ タイムゾーン　ユーザー定義型宣言 ■■■
'***************************************************************************************************
'* 機能　　：日付、時刻データを扱う場合は、UTC基準になる場合があるため、その設定です
'===================================================================================================
Private TargetTimeZone As TimeZone

Private Type TimeZone
    Hour    As Integer  '-23~23
    Minute  As Integer  '0-59
End Type



'***************************************************************************************************
'■■■ トーストの挙動を決めるユーザー定義型宣言 ■■■
'***************************************************************************************************
' 機能：Tag,Group,suppresspopup等といった[Windows.UI.Notifications]関連の設定値をまとめたものです
'===================================================================================================
Private ToastConfig As ToastNotification

Private Type ToastNotification
    '[Windows.UI.Notifications]の挙動を決める設定値
    AppUserModelID              As String   'https://learn.microsoft.com/ja-jp/windows/win32/shell/appids
    AllowUse_InternetImageFlag  As Boolean  'インターネット上の画像を使用する権限設定(実際は単に、Office関連のAppUserModelIDを切り替えてるだけ)
    XmlSchema                   As String   'https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-schema

    '[Windows.UI.Notifications]の主な設定値
    Tag                         As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.tag
    Group                       As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.group
    Schedule_ID                 As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtilenotification.id
    CollectionID                As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationmanagerforuser.gettoastnotifierfortoastcollectionidasync
    ExpiresOnReboot             As Boolean  'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.expiresonreboot
    SuppressPopup               As Boolean  'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.suppresspopup
    Schedule_DeliveryTime       As Date     'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.-ctor
    ExpirationTime              As Date     'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.expirationtime OR https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.expirationtime
End Type

'***************************************************************************************************
'■■■ VBA→DLLへの値参照が正しくできるようにするための中継用ユーザー定義型宣言 ■■■
'***************************************************************************************************
' 機能：スタック領域の制限の都合上、構造体としてを定義します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：C++側で、シグネチャ（型や順序）が合うようにすること。
'===================================================================================================
Private Type ToastNotificationParams
    '文字列に関するパラメーター　LongPtrを使うことで、ポインタから文字列を取得するようにします。
    AppUserModelID          As LongPtr
    XmlTemplate             As LongPtr
    Tag                     As LongPtr
    Group                   As LongPtr
    Schedule_ID             As LongPtr
    CollectionID            As LongPtr
    
    'スイッチングに関するパラメーター　「Boolean」では、上手く行かないため、longにしています
    ExpiresOnReboot         As Long
    SuppressPopup           As Long
    
    '日付に関するパラメーター
    Schedule_DeliveryTime As Double
    ExpirationTime As Double
End Type

'===================================================================================================



'***************************************************************************************************
'■■■ トースト スキーマに基づくメンバユーザー定義型宣言 ■■■
'***************************************************************************************************
'下記に基づいて順番に、宣言します
'→https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/root-elements

'===================================================================================================
'位置        ：Toast要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast
'---------------------------------------------------------------------------------------------------
Private ElementToast As AttributeToast

Private Type AttributeToast
    Duration         As String    '通知の表示時間延長フラグ
    Launch           As String    'トースト自体をクリックする際のスキーマ値
    DisplayTimestamp As String    'Windows プラットフォームによって通知が受信された時刻ではなく、通知コンテンツが実際に配信された日時を表すカスタム タイムスタンプで既定のタイムスタンプをオーバーライドします。 この値は、ISO 8601 標準を使用して書式設定された文字列である必要があります。
    ActivationType   As String    'ユーザーが特定の操作を行った際に使用されるアクティブ化の種類を決定します。公式にはないが、設定としてはあるもよう。VBA側では基本、"protocol"でしか価値がありません
    Synalio          As String    '通知のシナリオ属性
    UseButtonStyle   As Boolean   'スタイル付きボタン使用フラグ
End Type

'===================================================================================================
'位置        ：binding親要素のimage子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-image
'---------------------------------------------------------------------------------------------------
Private ElementsImage As ElementImage

'image要素で使える属性
Private Type AttributeImage
    addImageQuery As Boolean    'イメージ URI にクエリ文字列を追加できるようにするかの設定値です
    alt           As String     '支援技術のユーザー向けの画像の説明
    src           As String     'イメージ ソースの URI。ローカル画像や、httpsソースも使える
    Placement     As String     'イメージ配置
    hintcrop      As String     '円にトリミングする印
End Type

'image要素の立ち位置(属性ノード：placement  の値ごとに分けています)
Private Type ElementImage
    AppLogo As AttributeImage  'アイコン
    Inline  As AttributeImage  '画像はテキスト要素の後にインラインで表示され、ビジュアル領域の全幅が埋め込まれます。
    Hero    As AttributeImage  'トースト内で目立つように表示される注目のヒーロー画像を表示。常に通知上部に配置されます
End Type

'====================================================================================================
'位置        ：binding親要素のtext子要素で使える値と、属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-text
'----------------------------------------------------------------------------------------------------
Private ElementsText As ElementText

'text要素で使える属性
Private Type AttributeText
    HintCallScenarioCenterAlign As Boolean  '着信通知のテキストを中央に配置するには、"true" に設定します。 この値は、 シナリオ 値が "incomingCall" の通知にのみ使用されます。それ以外の場合は無視されます。
    Placement                   As String   'テキストの配置
End Type

'使える属性値と、要素の値そのもの
Private Type ParameterText
    '属性一覧
    Attribute As AttributeText

    '要素そのものの値
    Value     As String
End Type

'text要素の立ち位置
Private Type ElementText
    Title       As ParameterText    'タイトル
    Body        As ParameterText    'コンテンツ
    Attribution As ParameterText    '常に下部に表示するもの
End Type

'====================================================================================================
'位置        ：binding親要素のprogress子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-progress
'----------------------------------------------------------------------------------------------------
Private ElementProgress As AttributeProgress

Private Type AttributeProgress
    Title               As String   '省略可能なタイトル文字列。
    Status              As String   '左側の進行状況バーの下に表示されるステータス文字列。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります。
    Value               As String   'プログレスバーの値。 この値には、0.0 ~ 1.0 の浮動小数点数、または読み込みアニメーションを生成する値 "Indeterminate" を指定します。
    ValueStringOverride As String   '既定のパーセンテージ文字列の代わりに表示される省略可能な文字列。 これを指定しない場合は、"70%" などの文字が表示されます。
End Type

'====================================================================================================
'位置        ：Audio要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio
'----------------------------------------------------------------------------------------------------
Private ElementAudio As AttributeAudio

Private Type AttributeAudio
    loop   As Boolean   'ループフラグ
    silent As Boolean   '無音フラグ
    src    As String    '通知音名称
End Type

'====================================================================================================
'位置        ：actions親要素のaction子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-action
'----------------------------------------------------------------------------------------------------
'入れれる最大要素数を定数定義
Private Const MaxElementActionNum As Integer = 5    'ボタンは最大5つ、入れれます

Private ElementAction(1 To MaxElementActionNum) As AttributeAction

Private Type AttributeAction
    Content                 As String   'ボタンに表示されるコンテンツ。
    Arguments               As String   'ユーザーがこのボタンをクリックした場合にアプリが後から受け取る、アプリで定義された引数の文字列です。
    ActivationType          As String   'ユーザーが特定の操作を行った際に使用されるアクティブ化の種類を決定します。
    afterActivationBehavior As String   'ユーザーがトーストに対して操作を行ったときに、トーストが使用する動作を指定します。
    Placement               As String   '"contextMenu" に設定すると、アクションは、従来のトースト ボタンではなく、トースト通知のコンテキスト メニューに追加されたコンテキスト メニュー アクションになります。
    ImageUri                As String   'トースト ボタン アイコンのイメージ ソースの URI。
    HintInputId             As String   'input要素の横にあるボタンの ID に設定します。
    HintButtonStyle         As String   'ボタンのスタイル。
    HintToolTip             As String   'ボタンに空のコンテンツ文字列がある場合のボタンのヒント。
End Type

'====================================================================================================
'位置        ：input親要素のselection子要素で使える属性値。
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-selection
'----------------------------------------------------------------------------------------------------
'入れれる最大要素数を定数定義
Private Const MaxElementSelectionNum As Integer = 5    '選択肢は最大5つ、入れれます

Private ElementSelection(1 To MaxElementSelectionNum) As AttributeSelection

Private Type AttributeSelection
    MinuteValue As String  '再通知する時間。(単位：分)
    ChoseName   As String   '選択肢に出す名称
End Type

'====================================================================================================
'位置        ：actions親要素のinput子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-input
'----------------------------------------------------------------------------------------------------
'入れれる最大要素数を定数定義
Private Const MaxElementInputNum As Integer = 5    '入力要素は最大5つ、入れれます

Private ElementInput(1 To MaxElementInputNum) As AttributeInput

Private Type AttributeInput
    '属性一覧
    ID                 As String    '入力に関連付けられている ID
    Type               As String    'この属性には、"text"、"selection" のいずれかの値を指定できます。
    PlaceHolderContent As String    'テキスト入力用に表示されるプレースホルダー
    Title              As String    '入力のラベルとして表示されるテキスト。
    DefaultInput       As String    'デフォルトの入力値(公式にはないが、設定項目としてはあるもよう)

    '要素そのものの配列(Select要素が入ります)
    Value() As AttributeSelection
End Type

'====================================================================================================
'位置        ：header要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-header
'----------------------------------------------------------------------------------------------------
Private ElementHeader As AttributeHeader

Private Type AttributeHeader
    ID             As String    'このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。
    Title          As String    'ヘッダーのタイトルです。
    Arguments      As String    'ユーザーがこのヘッダーをクリックするとアプリに返されます。 null にすることはできません。
    ActivationType As String    'このヘッダーがクリックされた場合に使用するアクティブ化の種類。
End Type



'***************************************************************************************************
' ■■■ VBA上で使うのに相応しいAppIDをプリセット ■■■
'***************************************************************************************************
Private Const AppID_Excel        As String = "Microsoft.Office.EXCEL.EXE.15"                                            'DeskTopアプリ：Excel
Private Const AppID_Microsoft365 As String = "Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe!Microsoft.MicrosoftOfficeHub"  'Store  アプリ：Microsoft365(Office) PWA
Private AppID_Any As String                                                                                            '任意のAppID設定用変数



'***************************************************************************************************
' ■■■ PowerShell経由で実行する際の定数一式 ■■■
'***************************************************************************************************
'PowerShell起動コマンド
Private Const ActionPS As String = "powershell -Command """

'xmlコンテンツを制御するオブジェクトを定義し、xml内容を読み込むコマンド文字列
Private Const ReadXml As String = "$XmlDocument = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]::New();$XmlDocument.loadXml($xml)"



'***************************************************************************************************
' ■■■ Toast全般の設定プロパティ ■■■
'***************************************************************************************************
'* 機能    ：インターネット上の画像ソースを使うか決める変数をセットします
'---------------------------------------------------------------------------------------------------
Public Property Let AllowUse_InternetImage(ByVal ArgFlag As Boolean)
    ToastConfig.AllowUse_InternetImageFlag = ArgFlag
End Property

'===================================================================================================
'* 機能    ：この通知をどのアプリで出すかを設定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：設定すると、「AllowUse_InternetImageFlag」よりも優先して設定されます。batファイル等、VBAと直接関係なくなり、アイコンが噛み合わない時に使います。
'* 注意事項：存在しない(未インストール)AppID、無効な文字列を指定すると、Toastが発行されないのでご注意ください。
'---------------------------------------------------------------------------------------------------
Public Property Let SetToastContent_AppUserModelID(ByVal ArgAppID As String)
    AppID_Any = ArgAppID
End Property

'===================================================================================================
'* 機能    ：タイムゾーンを変更します。デフォルトは、日本基準のため日本在宅者は基本、使用する必要はありません
'---------------------------------------------------------------------------------------------------
'* 注意事項：Class_Initialize　で、日本時間にセットされます
'---------------------------------------------------------------------------------------------------
Public Sub SetTimeZone(ByVal ArgUTC_Hour As Integer, ByVal ArgUTC_Minute As Integer)
    With TargetTimeZone
        .Hour = ArgUTC_Hour
        .Minute = ArgUTC_Minute
    End With
End Sub

'===================================================================================================
'* 機能    ：設定値を纏めたユーザー定義型：ToastNotification　を基に、DLLで扱えるユーザー定義型：ToastNotificationParams　に変換します
'---------------------------------------------------------------------------------------------------
'* 注意事項：Windows API の文字列関係は、ポインターで取得するパターンが多いので、StrPtrで経由する必要があります
'---------------------------------------------------------------------------------------------------
Private Function ParamsConvertForDLL(ToastSetting As ToastNotification) As ToastNotificationParams
    With ParamsConvertForDLL
        '文字列から、文字列ポインターに変換
        .AppUserModelID = StrPtr(ToastSetting.AppUserModelID)
        .XmlTemplate = StrPtr(ToastSetting.XmlSchema)
        .Tag = StrPtr(ToastSetting.Tag)
        .Group = StrPtr(ToastSetting.Group)
        .Schedule_ID = StrPtr(ToastSetting.Schedule_ID)
        .CollectionID = StrPtr(ToastSetting.CollectionID)
        
        'スイッチングをセット。Boolean→BOOL の変換が上手く行かないため、longで代用します。
        .ExpiresOnReboot = ToastSetting.ExpiresOnReboot * -1
        .SuppressPopup = ToastSetting.SuppressPopup * -1
    
        '日付をセット。DLL(C++)での処理の場合、UTC基準になるため、設定済みのタイムゾーンを基に、補正します。
        'スケジュール通知
        .Schedule_DeliveryTime = DateAdd("h", TargetTimeZone.Hour * -1, ToastSetting.Schedule_DeliveryTime)
        .Schedule_DeliveryTime = DateAdd("n", TargetTimeZone.Minute * -1, .Schedule_DeliveryTime)
        
        '有効期限(上限は3日後)
        .ExpirationTime = DateAdd("h", TargetTimeZone.Hour * -1, ToastSetting.ExpirationTime)
        .ExpirationTime = DateAdd("n", TargetTimeZone.Minute * -1, .ExpirationTime)
    End With
End Function



'***************************************************************************************************
' ■■■ Toastを構成するスキーマ情報を決める基となるメンバ変数へ設定するプロパティ ■■■
'***************************************************************************************************
'記述順番と名称はは、下記のリンクに準拠します。
'→https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-schema
'値、固定値の設定は、ここで行う
'---------------------------------------------------------------------------------------------------
'* 機能　　：toast要素のduration属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgExtensionDuration
'               True ：long  として扱います。25秒間、表示されます。
'               False：short として扱います。省略も同等です。
'---------------------------------------------------------------------------------------------------
'* 注意事項：アクセシビリティ→視覚効果→この時間が経過したら通知を破棄する  で表示時間をオーバーライド可能です。(longの場合は、上回る場合にオーバーライド)
'***************************************************************************************************
Public Property Let SetToastContent_Duration(ByVal ArgExtensionDuration As Boolean)
    If ArgExtensionDuration Then ElementToast.Duration = "long"
End Property

'***************************************************************************************************
'* 機能　　：toast要素のlaunch、activationType属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgSchema          起動スキーマ。URLやファイルパスなど
'            ArgActivationType  このヘッダーがクリックされた場合に使用するアクティブ化の種類(VBAでは、"protocol"以外では意味をなしません)
'***************************************************************************************************
Public Property Let SetToastContent_Launch(Optional ByVal ArgActivationType As String = "protocol", ByVal ArgSchema As String)
    With ElementToast
        .Launch = ArgSchema
        .ActivationType = ArgActivationType     'プロトコルのアクティブ化を使用して別のアプリを起動します。(ArgActivationTypeがprotocolの場合)
    End With
End Property

'***************************************************************************************************
'* 機能　　：toast要素のdisplayTimestamp属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgDisplayTimestamp    カスタム タイムスタンプで既定のタイムスタンプをオーバーライドします
'---------------------------------------------------------------------------------------------------
'* 注意事項：※ISO 8601で指定する都合上、UTCを設定する引数を設けますが、JSTの場合は気にする必要はありません。
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/custom-timestamps-on-toasts
'***************************************************************************************************
Public Property Let SetToastContent_DisplayTimestamp(Optional ArgUTC_Hour As Integer = 9, Optional ArgUTC_Minute As Integer = 0, ByVal ArgDisplayTimestamp As Date)
    With WorksheetFunction
        If ArgDisplayTimestamp Then ElementToast.DisplayTimestamp = .Text(ArgDisplayTimestamp, "yyyy-mm-ddThh:mm:ss") & .Text(TargetTimeZone.Hour, "+00:;-00:") & .Text(TargetTimeZone.Minute, "00")
    End With
End Property

'***************************************************************************************************
'* 機能　　：toast要素のsynalio属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID    シナリオID。便宜上、数値に応じてシナリオ名を設定します
'***************************************************************************************************
Public Property Let SetToastScenario(ToastScenarioMode As ToastScenario)
    'ArgIDから、シナリオ名を決めます。
    With ElementToast
        Select Case ToastScenarioMode
            Case 1
                '通知を永遠に表示する。action要素がないと効果は発動しないが、任意の通知音に設定可能
                .Synalio = "reminder"

            Case 2
                '通知を永遠に表示する。action要素がないと効果は発動しないが、応答不可になっても設定問わず、必ず表示される。通知音はアラーム系(Alarm)か、ミュートのみ設定可能
                .Synalio = "alarm"

            Case 3
                '通知を永遠に表示する。action要素がなくても効果発動。最後のボタン位置のみ、Windowsのテーマ色に基づく着色が施され、位置が必ず下側になる。通知音は呼び出し系(Call)か、ミュートのみ設定可能
                .Synalio = "incomingCall"

            Case 4
                '重要な通知として表示。応答不可にしても表示されますが、設定で表示させないようにすることが可能です。とりあえず常に表示させたい場合は、シナリオ：「アラーム」にしましょう。
                .Synalio = "urgent"

            Case Else
                '通常の通知として表示。属性：duration  あるいは、アクセシビリティ→視覚効果→この時間が経過したら通知を破棄する  に基づく表示時間となります
                .Synalio = vbNullString
        End Select
    End With
End Property

'***************************************************************************************************
'* 機能　　：toast要素のuseButtonStyle属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArguseButtonStyle
'               True ：スタイル付きボタンを使用します
'               False：スタイル付きボタンを使用しません。
'---------------------------------------------------------------------------------------------------
'* 注意事項：この機能は、Windows 11 以降のみ機能します。
'***************************************************************************************************
Public Property Let AllowToastContent_UseButtonStyle(ByVal ArgUseButtonStyle As Boolean)
    ElementToast.UseButtonStyle = ArgUseButtonStyle
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のimage子要素のうち、インラインに設定する画像のパスと、丸いロゴフラグの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：Arg_InlineImageSrc    イメージ ソースの URI。ファイルパスや、http(s)画像など
'            Arg_InlineCircle
'               True ：画像は円にトリミングされます。
'               False：画像はトリミングされず、そのまんまのサイズとして表示されます。

'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Arg_Alt                画像の説明を記載します。
'***************************************************************************************************
Public Property Let SetToastContent_ImageInline(Optional ByVal Arg_InlineCircle As Integer, Optional ByVal Flag_addImageQuery As Boolean, Optional ByVal Arg_Alt As String, ByVal Arg_InlineImageSrc As String)
    With ElementsImage.Inline
        .src = Arg_InlineImageSrc
        .addImageQuery = Flag_addImageQuery
        .alt = Arg_Alt

        If Arg_InlineCircle Then .hintcrop = "circle"
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のimage子要素のうち、ヒーロー画像のパスの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：Arg_HeroImageSrc       イメージ ソースの URI。ファイルパスや、http(s)画像など
'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Arg_Alt                画像の説明を記載します。
'***************************************************************************************************
Public Property Let SetToastContent_ImageHero(Optional ByVal Flag_addImageQuery As Boolean, Optional ByVal Arg_Alt As String, ByVal Arg_HeroImageSrc As String)
    With ElementsImage.Hero
        .src = Arg_HeroImageSrc
        .addImageQuery = Flag_addImageQuery
        .alt = Arg_Alt

        .Placement = "hero"
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のimage子要素のうち、AppLogo(appLogoOverride)に設定する画像のパスと、丸いロゴフラグの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：Arg_AppLogoSrc    イメージ ソースの URI。ファイルパスや、http(s)画像など
'            Arg_LogoCircle
'               True ：画像は円にトリミングされます。
'               False：画像はトリミングされず、正方形として表示されます。

'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Arg_Alt                画像の説明を記載します。
'***************************************************************************************************
Public Property Let SetToastContent_ImageAppLogo(Optional ByVal Arg_LogoCircle As Boolean, Optional ByVal Flag_addImageQuery As Boolean, Optional ByVal Arg_Alt As String, ByVal Arg_AppLogoSrc As String)
    With ElementsImage.AppLogo
        .src = Arg_AppLogoSrc
        .addImageQuery = Flag_addImageQuery
        .alt = Arg_Alt

        .Placement = "appLogoOverride"

        If Arg_LogoCircle Then .hintcrop = "circle"
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のtext子要素に値と属性をそれぞれ設定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：TextContents                   通知内容
'            HintCallScenarioCenterAlign    テキストを中央に配置するフラグ(Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：1行目のタイトルにあたります
'* 注意事項：Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'***************************************************************************************************
Public Property Let SetToastContent_TextTitle(Optional ByVal HintCallScenarioCenterAlign As Boolean, ByVal TextContents As String)
    With ElementsText.Title
        .Value = TextContents
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のtext子要素に値と属性をそれぞれ設定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：TextContents                   通知内容
'            HintCallScenarioCenterAlign    テキストを中央に配置するフラグ(Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：2行目以降の通知内容にあたります
'* 注意事項：Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'            また、タイトルが未設定の場合、これがタイトル扱いになるため注意
'***************************************************************************************************
Public Property Let SetToastContent_TextBody(Optional ByVal HintCallScenarioCenterAlign As Boolean, ByVal TextContents As String)
    With ElementsText.Body
        .Value = TextContents
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のtext子要素に値と属性をそれぞれ設定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：TextContents                   通知内容
'            HintCallScenarioCenterAlign    テキストを中央に配置するフラグ(Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：通知の下部に設定するテキストを設定します。
'* 注意事項：Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'            placement属性に、"attribution"を常に付与しています
'***************************************************************************************************
Public Property Let SetToastContent_TextAttribute(Optional ByVal HintCallScenarioCenterAlign As Boolean, ByVal TextContents As String)
    With ElementsText.Attribution
        .Value = TextContents
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign

        .Attribute.Placement = "attribution"
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のProgress子要素に必要な属性を一括設定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Arg_Title                  タイトルの文字列 (オプション) を取得または設定します。 データ バインディングをサポートしています。
'            Arg_Value                  進行状況バーの値を取得または設定します。 データ バインディングをサポートしています。 既定値は 0 です。 0.0 ～ 1.0 の double 型で、ゲージ式。負の値にすると、ドットアニメーションの不確定式になります。
'            Arg_ValueStringOverride    割合を示す既定の文字列に代わって表示される文字列 (オプション) を取得または設定します。 これを指定しない場合は、"70%" などの文字が表示されます。
'            Arg_Status                 進行状況バーの下の左側に表示されるステータス文字列 (必須) を取得または設定します。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります。
'---------------------------------------------------------------------------------------------------
'URL       ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-progress-bar
'***************************************************************************************************
Private Sub SetToastContent_ProgressBar(Arg_Title As String, Arg_Value As Double, Arg_ValueStringOverride As String, Arg_Status As String)
    With ElementProgress
        .Title = Arg_Title
        .Status = Arg_Status

        '0未満で渡される場合、特定の値は表示されず、処理中であることをアニメーション ドットで示すフラグ「Indeterminate」を設定します
        If Arg_Value < 0 Then
            .Value = "Indeterminate"
        Else
            .Value = Arg_Value
        End If

        .ValueStringOverride = Arg_ValueStringOverride
    End With
End Sub

'***************************************************************************************************
'* 機能　　：Audio要素の属性をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgSrcID       通知音ID。便宜上、数値に応じて通知音名を設定します
'               0-24        Windows側で定められた通知音名を番号で指定
'               上記以外    ミュートにします

'            ArgLoop
'               True  ：トーストが表示されている限りサウンドが繰り返される場合に設定
'               False ：1回だけ再生する場合に設定
'***************************************************************************************************
Public Property Let SetToastAudio(Optional ByVal ArgLoop As Boolean, ArgSrcID As MsWinSoundEvent)
    'Windows側で定められた通知音数
    Const DefNoticeNum As Integer = 25

    With ElementAudio
        'audio要素のLoop属性を設定
        .loop = ArgLoop

        'ArgSrcID に応じて、ミュート判定を決める
        If 0 < ArgSrcID And ArgSrcID <= DefNoticeNum Then
            '存在する番号ならミュートしない
            .silent = False
        Else
            '存在しない番号ならミュート
            .silent = True
        End If

        'audio要素のsrc属性を設定
        'ArgSrcIDから、通知音ファイル名を決めます。Windows側で決められた範囲でしか使えないので注意
        '→https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio#attributes
            
        'ただし、toast要素のsynalio属性値によっては、この設定が効かないことがあります
        Select Case ArgSrcID
            Case 1
                .src = "ms-winsoundevent:Notification.Default"

            Case 2
                .src = "ms-winsoundevent:Notification.IM"
                
            Case 3
                .src = "ms-winsoundevent:Notification.Mail"

            Case 4
                .src = "ms-winsoundevent:Notification.Reminder"

            Case 5
                .src = "ms-winsoundevent:Notification.SMS"

            Case 6
                .src = "ms-winsoundevent:Notification.Looping.Alarm"

            Case 7
                .src = "ms-winsoundevent:Notification.Looping.Alarm2"

            Case 8
                .src = "ms-winsoundevent:Notification.Looping.Alarm3"

            Case 9
                .src = "ms-winsoundevent:Notification.Looping.Alarm4"

            Case 10
                .src = "ms-winsoundevent:Notification.Looping.Alarm5"

            Case 11
                .src = "ms-winsoundevent:Notification.Looping.Alarm6"

            Case 12
                .src = "ms-winsoundevent:Notification.Looping.Alarm7"

            Case 13
                .src = "ms-winsoundevent:Notification.Looping.Alarm8"

            Case 14
                .src = "ms-winsoundevent:Notification.Looping.Alarm9"

            Case 15
                .src = "ms-winsoundevent:Notification.Looping.Alarm10"

            Case 16
                .src = "ms-winsoundevent:Notification.Looping.Call"

            Case 17
                .src = "ms-winsoundevent:Notification.Looping.Call2"

            Case 18
                .src = "ms-winsoundevent:Notification.Looping.Call3"

            Case 19
                .src = "ms-winsoundevent:Notification.Looping.Call4"

            Case 20
                .src = "ms-winsoundevent:Notification.Looping.Call5"

            Case 21
                .src = "ms-winsoundevent:Notification.Looping.Call6"

            Case 22
                .src = "ms-winsoundevent:Notification.Looping.Call7"

            Case 23
                .src = "ms-winsoundevent:Notification.Looping.Call8"

            Case 24
                .src = "ms-winsoundevent:Notification.Looping.Call9"

            Case 25
                .src = "ms-winsoundevent:Notification.Looping.Call10"

            Case Else
                .src = vbNullString

        End Select

    End With
End Property

'***************************************************************************************************
'* 機能　　：action要素の各属性にセットしたものをactions要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgContent         ボタンに表示されるコンテンツ
'            ArgArguments       起動スキーマ。URLやファイルパスなど
'            ArgActivationType  このボタンが押下された場合に使用するアクティブ化の種類。VBAでは、基本"protocol"以外では意味をなしませんが、ここのボタンの場合、"system"にして、所々の設定を行うことで、システムの再通知と無視アクションを利用するアラーム通知を作成できます。
'            ArgPendingUpdate   ユーザーがトースト上のボタンをクリックすると、通知は "保留中の更新" 表示状態のままです。 この "更新の保留中" の表示状態が長時間続くことを避けるため、バックグラウンド タスクから即座にトーストを更新する必要があります。
'            ArgContextMenu     Trueで、従来のトースト ボタンではなく、トースト通知のコンテキスト メニューに追加されたコンテキスト メニュー アクションになります。
'            ArgIcon            トースト ボタン アイコンのイメージソース
'            ArgHintInputId     テキストBoxの横に来させたいIDを指定します
'            ArgHintButtonStyle ボタンのスタイル。現状、「緑、赤」の2色のみ対応です。便宜上、数値に応じてスタイル名を設定します。※この機能は、Windows 11 以降のみ機能します。
'            ArgHintToolTip     ボタンに空のコンテンツ文字列がある場合のボタンのヒント。
'            Index              追加位置
'***************************************************************************************************
Public Property Let SetIToastActions(ByVal ArgContent As String, ByVal ArgArguments As String, Optional ByVal ArgActivationType As String = "protocol", Optional ArgPendingUpdate As Boolean, Optional ByVal ArgContextMenu As Boolean, Optional ByVal ArgIcon As String, Optional ByVal ArgHintInputId As String, Optional ArgHintButtonStyle As ToastButtonStyle, Optional ByVal ArgHintToolTip As String, ByVal Index As Integer)
    With ElementAction(Index)

        .Content = ArgContent
        .Arguments = ArgArguments
        .ActivationType = ArgActivationType

        If ArgPendingUpdate Then .afterActivationBehavior = "pendingUpdate"
        If ArgContextMenu Then .Placement = "contextMenu"

        .ImageUri = ArgIcon
        .HintInputId = ArgHintInputId

        '属性：hint-buttonStyle  は、引数の値に応じて命名します
        Select Case ArgHintButtonStyle
            Case 1
                'ボタンを緑色にします。
                .HintButtonStyle = "Success"
            Case 2
                'ボタンを赤色にします。
                .HintButtonStyle = "Critical"
            Case Else
                '装飾は、施しません。
                .HintButtonStyle = vbNullString
        End Select

        .HintToolTip = ArgHintToolTip
    End With
End Property

'***************************************************************************************************
'* 機能　　：input要素の各属性にセットしたものをaction要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID                  入力に関連付けられている ID
'            ChoseFlag              Trueで"selection"、Falseで、"text"扱いとします
'            ArgPlaceHolderContent  テキスト入力用に表示されるプレースホルダー。
'            Argtitle               入力のラベルとして表示されるテキスト
'            ArgDefaultInput        デフォルトの入力値
'            Index                  追加位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：VBAでは、ChoseFlag：text　は無意味です。
'***************************************************************************************************
Public Property Let SetIToastInput(ByVal ArgID As String, Optional ByVal ChoseFlag As Boolean, Optional ByVal ArgPlaceHolderContent As String, Optional ByVal ArgTitle As String, Optional ByVal ArgDefaultInput As String, ByVal Index As Integer)
    With ElementInput(Index)
        .ID = ArgID

        'Typeが、"selection"の場合、予め作成されたSelection要素を入れるようにします。
        If ChoseFlag Then
            .Type = "selection"
            .Value = ElementSelection
        Else
            .Type = "text"
        End If

        .PlaceHolderContent = ArgPlaceHolderContent
        .Title = ArgTitle
        .DefaultInput = ArgDefaultInput
    End With
End Property

'***************************************************************************************************
'* 機能　　：input要素の各属性にセットしたものをaction要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ReminderMinute 何分後にリマインダー通知させるか？
'            ArgChoseName   選択項目の名称
'            Index          挿入位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：VBAでは基本、リマインダー用途しか使い所がありません
'***************************************************************************************************
Public Property Let SetToastSelectionBox(ByVal ReminderMinute As String, ByVal ArgChoseName As String, ByVal Index As Integer)
    With ElementSelection(Index)
        .MinuteValue = ReminderMinute
        .ChoseName = ArgChoseName
    End With
End Property

'***************************************************************************************************
'* 機能　　：header要素の各属性にセットしたものをtoast要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID            ：このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。(省略時：このExcelのファイル名)
'            ArgTitle         ：ヘッダーのタイトルです。(省略時：このExcelのBook名)
'            ArgArguments     ：起動スキーマ。URLやファイルパスなど(省略時：このExcelの保存フォルダパス)
'            ArgActivationType：このヘッダーがクリックされた場合に使用するアクティブ化の種類
'---------------------------------------------------------------------------------------------------
'* 注意事項：VBAでは、ArgActivationType:"protocol"以外では意味をなしません
'***************************************************************************************************
Public Property Let SetToastHeader(Optional ByVal ArgID As String, Optional ByVal ArgArguments As String, Optional ByVal ArgActivationType As String = "protocol", ByVal ArgTitle As String)
    With ElementHeader
        '省略時：このExcelのファイル名
        If StrPtr(ArgID) Then
            .ID = ArgID
        Else
            .ID = ThisWorkbook.Name
        End If

        '省略時：このExcelの保存フォルダパス
        If StrPtr(ArgArguments) Then
            .Arguments = ArgArguments
        Else
            .Arguments = ThisWorkbook.Path
        End If

        '省略時：protocol
        .ActivationType = ArgActivationType

        .Title = ArgTitle
    End With
End Property



'***************************************************************************************************
' ■■■ Toastを構成するスキーマ情報を、メンバ変数を基に、動的生成します ■■■
'***************************************************************************************************
'* 返り値　 ：トースト スキーマのxmlContentsが返る
'* 引数　 　：UseProgress    trueで、progress要素を含ませます
'---------------------------------------------------------------------------------------------------
'* 注意事項 ：「Microsoft XML. v6.0」の参照設定が必要です
'* URL      ：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/schema-root
'***************************************************************************************************
Private Function MakeXmlToastContentBuilder(UseProgress As Boolean) As String
    'IXMLDOMNodeオブジェクトを宣言
    Dim ToastContents As MSXML2.DOMDocument60



    '--------------------ルート要素(toast)--------------------
    Dim RootNode_toast As MSXML2.IXMLDOMNode: Const NameRootNode_toast As String = "toast"

    '--------------------toast要素内で使える子要素オブジェクト--------------------
    Dim Node01__visual  As MSXML2.IXMLDOMNode: Const NameNode01__visual  As String = "visual"
    Const NameNode02__audio   As String = "audio"
    Dim Node03__actions As MSXML2.IXMLDOMNode: Const NameNode03__actions As String = "actions"
    Const NameNode04__header  As String = "header"

    '--------------------visual要素内で使える子要素オブジェクト--------------------
    Dim Node01_01__binding As MSXML2.IXMLDOMNode: Const NameNode01_01__binding As String = "binding"

    '--------------------binding要素内で使う要素関係の宣言--------------------
    'image要素
    Const NameNode01_01_01__image As String = "image"
    'text要素
    Const NameNode01_01_02__text As String = "text"
    'progress要素
    Const NameNode01_01_03__progress As String = "progress"

    '--------------------action要素内で使える要素関係の宣言--------------------
    'input要素
    Dim Node03_01__input As MSXML2.IXMLDOMNode: Const NameNode03_01__input As String = "input"
    'action要素
    Const NameNode03_02__action As String = "action"

    '--------------------input要素内で使える要素関係の宣言--------------------
     Const NameNode03_01_01__selection As String = "selection"



    '初期化
    Set ToastContents = New MSXML2.DOMDocument60
    ToastContents.async = False



    '-------------------------------------------------------Root要素を設定-------------------------------------------------------
    'ルートノード：toast  を作成して最初のルートノードとして追加します。
    Set RootNode_toast = ToastContents.appendChild(ToastContents.createNode(NODE_ELEMENT, NameRootNode_toast, ""))

    With RootNode_toast
        '属性ノード：duration  にTrueを設定してあったら、値を設定し、ルートノードの属性に追加します。
        If StrPtr(ElementToast.Duration) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "duration", "")).NodeValue = ElementToast.Duration
        '属性ノード：launch  に値を設定してあったら、値を設定し、ルートノードの属性に追加します。
        If StrPtr(ElementToast.Launch) Then
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "launch", "")).NodeValue = ElementToast.Launch
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "activationType", "")).NodeValue = ElementToast.ActivationType
        End If
        '属性ノード：displayTimestamp  に値を設定してあったら、値を設定し、ルートノードの属性に追加します。
        If StrPtr(ElementToast.DisplayTimestamp) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "displayTimestamp", "")).NodeValue = ElementToast.DisplayTimestamp
        '属性ノード：scenario  に値を設定してあったら、値を設定し、ルートノードの属性に追加します。
        If StrPtr(ElementToast.Synalio) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "scenario", "")).NodeValue = ElementToast.Synalio
        '属性ノード：useButtonStyle  にTrueを設定してあったら、値を設定し、ルートノードの属性に追加します。
        If ElementToast.UseButtonStyle Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "useButtonStyle", "")).NodeValue = LCase(ElementToast.UseButtonStyle)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします
    End With



    '-------------------------------------------------------visual,binding要素を設定-------------------------------------------------------
    'ノード：visual  を作成してルートノードの子ノードとして追加します。
    Set Node01__visual = RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01__visual, ""))

    'ノード：binding  を作成して「visual」ノードの子ノードとして追加します。
    Set Node01_01__binding = Node01__visual.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01__binding, ""))

    '属性ノード：template  を作成して値を設定し、先程の「binding」ノードの属性に追加します。
    '※この値は "ToastGeneric" に設定する必要があります
    Node01_01__binding.Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "template", "")).NodeValue = "ToastGeneric"



    '-------------------------------------------------------image要素関係の設定-------------------------------------------------------
    'ノード：image  を作成して「binding」ノードの子ノードとして追加します。
    With Node01_01__binding
        'ロゴ(アイコン)画像が設定してあったら、1つ目の「image」ノードを追加します
        If StrPtr(ElementsImage.AppLogo.src) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_01__image, ""))
                '属性ノード：addImageQuery  に値があったら、先ほどのノード：image  の属性に追加します
                If ElementsImage.AppLogo.addImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsImage.AppLogo.addImageQuery)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

                '属性ノード：alt  に値があったら、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.AppLogo.alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.AppLogo.alt

                '属性ノード：src  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementsImage.AppLogo.src

                '属性ノード：placement  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placement", "")).NodeValue = ElementsImage.AppLogo.Placement

                '円にトリミングするフラグがあったら、属性ノード：hint-crop  に値を設定したのを、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.AppLogo.hintcrop) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-crop", "")).NodeValue = ElementsImage.AppLogo.hintcrop
            End With
        End If

        'インライン画像が設定してあったら、2つ目の「image」ノードを追加します
        If StrPtr(ElementsImage.Inline.src) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_01__image, ""))
                '属性ノード：addImageQuery  に値があったら、先ほどのノード：image  の属性に追加します
                If ElementsImage.Inline.addImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsImage.Inline.addImageQuery)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

                '属性ノード：alt  に値があったら、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.Inline.alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.Inline.alt

                '属性ノード：src  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementsImage.Inline.src

                '円にトリミングするフラグがあったら、属性ノード：hint-crop  に値を設定したのを、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.Inline.hintcrop) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-crop", "")).NodeValue = ElementsImage.Inline.hintcrop
            End With
        End If

        'ヒーロー画像が設定してあったら、3つ目の「image」ノードを追加します
        If StrPtr(ElementsImage.Hero.src) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_01__image, ""))
                '属性ノード：addImageQuery  に値があったら、先ほどのノード：image  の属性に追加します
                If ElementsImage.Hero.addImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsImage.Hero.addImageQuery)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

                '属性ノード：alt  に値があったら、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.Hero.alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.Hero.alt

                '属性ノード：src  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementsImage.Hero.src

                '属性ノード：placement  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placement", "")).NodeValue = ElementsImage.Hero.Placement
            End With
        End If

    End With



    '-------------------------------------------------------text要素関係の設定-------------------------------------------------------
    'ノード：text  を作成して「binding」ノードの子ノードとして追加します。
    With Node01_01__binding
        'タイトルを設定してあったら、1つ目の「text」ノードを追加します。
        If StrPtr(ElementsText.Title.Value) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_02__text, ""))
                '先ほどのノード：text  に値を設定します
                .Text = ElementsText.Title.Value

                 '属性ノード：hint-callScenarioCenterAlign  に値を設定してあったら、先ほどのノード：text  の属性に追加します
                If ElementsText.Title.Attribute.HintCallScenarioCenterAlign Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-callScenarioCenterAlign", "")).NodeValue = LCase(ElementsText.Title.Attribute.HintCallScenarioCenterAlign)
            End With
        End If

        '内容を設定してあったら、2つ目の「text」ノードを追加を追加します
        If StrPtr(ElementsText.Body.Value) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_02__text, ""))
                '先ほどのノード：text  に値を設定します
                .Text = ElementsText.Body.Value

                 '属性ノード：hint-callScenarioCenterAlign  に値を設定してあったら、先ほどのノード：text  の属性に追加します
                If ElementsText.Body.Attribute.HintCallScenarioCenterAlign Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-callScenarioCenterAlign", "")).NodeValue = LCase(ElementsText.Body.Attribute.HintCallScenarioCenterAlign)
            End With
        End If

        'フッターを設定してあったら、3つ目の「text」ノードを追加します
        If StrPtr(ElementsText.Attribution.Value) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_02__text, ""))
                '先ほどのノード：text  に値を設定します
                .Text = ElementsText.Attribution.Value

                '属性ノード：placement  に値を設定したのを、先ほどのノード：text  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placement", "")).NodeValue = ElementsText.Attribution.Attribute.Placement

                '属性ノード：hint-callScenarioCenterAlign  に値を設定してあったら、先ほどのノード：text  の属性に追加します
                If ElementsText.Attribution.Attribute.HintCallScenarioCenterAlign Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-callScenarioCenterAlign", "")).NodeValue = LCase(ElementsText.Attribution.Attribute.HintCallScenarioCenterAlign)
            End With
        End If
    End With



    '-------------------------------------------------------progress要素関係の設定-------------------------------------------------------
    '引数：UseProgress  がTrueなら、「progress」ノードを追加します
    If UseProgress Then
        With Node01_01__binding.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_03__progress, ""))
            '属性ノード：title  に値を設定したのを、先ほどのノード：progress  の属性に追加します
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "title", "")).NodeValue = "{progressTitle}"

            '属性ノード：status  に値を設定したのを、先ほどのノード：progress  の属性に追加します
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "status", "")).NodeValue = "{progressStatus}"

            '属性ノード：value  に値を設定したのを、先ほどのノード：progress  の属性に追加します
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "value", "")).NodeValue = "{progressValue}"

            '属性ノード：valueStringOverride  に値を設定してあったら、先ほどのノード：progress  の属性に追加します
            If StrPtr(ElementProgress.ValueStringOverride) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "valueStringOverride", "")).NodeValue = "{progressValueString}"
        End With
    End If



    '-------------------------------------------------------audio要素を設定-------------------------------------------------------
    'audio要素の属性：src,silent  いずれかに有効な値があったら、「audio」ノードを追加します
    If StrPtr(ElementAudio.src) Or ElementAudio.silent Then
        With RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode02__audio, ""))
            '属性ノード：silent  がtrueなら、作成して値を設定し、先程の「audio」ノードの属性に追加します。
            'この場合、他の属性値は無視します。
            If ElementAudio.silent Then
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "silent", "")).NodeValue = LCase(ElementAudio.silent) 'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします
            Else
                '属性ノード：loop  がtrueなら、作成して値を設定し、先程の「audio」ノードの属性に追加します。
                If ElementAudio.loop Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "loop", "")).NodeValue = LCase(ElementAudio.loop) 'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

                '属性ノード：src  を作成して値を設定し、先程の「audio」ノードの属性に追加します。
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementAudio.src

            End If
        End With
    End If



    '-------------------------------------------------------actions要素を設定-------------------------------------------------------
    'action要素に、1番目のContent属性または、input要素に、1番目のID属性が設定されてあったら、ノード：actions  を作成してルートノードの子ノードとして追加します。
    If StrPtr(ElementAction(1).Content) Or StrPtr(ElementInput(1).ID) Then
        Set Node03__actions = RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03__actions, ""))


        '-------------------------------------------------------input要素を設定-------------------------------------------------------
        '各属性ノードを作成して値を設定し、「input」ノードの属性に追加します。
        '5つまで、繰り返す
        Dim i As Integer, j As Integer
        For i = 1 To MaxElementInputNum
            '属性ノード：ID  がない場合は、作成をSKIPします
            If StrPtr(ElementInput(i).ID) Then
                Set Node03_01__input = Node03__actions.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03_01__input, ""))
                With Node03_01__input.Attributes
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementInput(i).ID
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "type", "")).NodeValue = ElementInput(i).Type
                    If StrPtr(ElementInput(i).PlaceHolderContent) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placeHolderContent", "")).NodeValue = ElementInput(i).PlaceHolderContent
                    If StrPtr(ElementInput(i).Title) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "title", "")).NodeValue = ElementInput(i).Title
                    If StrPtr(ElementInput(i).DefaultInput) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "defaultInput", "")).NodeValue = ElementInput(i).DefaultInput

                    '現在のInput要素にて、type属性が"selection"なら、要素番号に対応するSelection要素配列を挿入する
                    If ElementInput(i).Type = "selection" Then
                        For j = 1 To MaxElementSelectionNum
                            'Input要素のID属性がない場合、挿入をSKIPする
                            If StrPtr(ElementInput(i).Value(j).MinuteValue) Then
                                With Node03_01__input.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03_01_01__selection, "")).Attributes
                                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementInput(i).Value(j).MinuteValue
                                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "content", "")).NodeValue = ElementInput(i).Value(j).ChoseName
                                End With
                            End If
                        Next
                    End If
                End With
            End If
        Next



        '-------------------------------------------------------action要素を設定-------------------------------------------------------
        '各属性ノードを作成して値を設定し、「action」ノードの属性に追加します。
        '5つまで、繰り返す
        For i = 1 To MaxElementActionNum
            '属性ノード：Content  がない場合は、作成をSKIPします
            If StrPtr(ElementAction(i).Content) Then
                With Node03__actions.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03_02__action, "")).Attributes
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "content", "")).NodeValue = ElementAction(i).Content
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "arguments", "")).NodeValue = ElementAction(i).Arguments
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "activationType", "")).NodeValue = ElementAction(i).ActivationType
                    If StrPtr(ElementAction(i).afterActivationBehavior) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "afterActivationBehavior", "")).NodeValue = ElementAction(i).afterActivationBehavior
                    If StrPtr(ElementAction(i).Placement) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placement", "")).NodeValue = ElementAction(i).Placement
                    If StrPtr(ElementAction(i).ImageUri) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "imageUri", "")).NodeValue = ElementAction(i).ImageUri
                    If StrPtr(ElementAction(i).HintInputId) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-inputId", "")).NodeValue = ElementAction(i).HintInputId
                    If StrPtr(ElementAction(i).HintButtonStyle) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-buttonStyle", "")).NodeValue = ElementAction(i).HintButtonStyle
                    If StrPtr(ElementAction(i).HintToolTip) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-toolTip", "")).NodeValue = ElementAction(i).HintToolTip
                End With
            End If
        Next
    End If



    '-------------------------------------------------------header要素を設定-------------------------------------------------------
    'header要素に、id属性が設定されてあったら、ノード：header  を作成してルートノードの子ノードとして追加します。
    If StrPtr(ElementHeader.ID) Then
        With RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode04__header, ""))
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementHeader.ID
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "title", "")).NodeValue = ElementHeader.Title
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "arguments", "")).NodeValue = ElementHeader.Arguments
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "activationType", "")).NodeValue = ElementHeader.ActivationType
        End With
    End If



    '上記の処理で生成されたxmlコンテンツを返します(末尾に改行が入らないように置換を行います)
    'Debug.Print ToastContents.XML
    MakeXmlToastContentBuilder = Replace(ToastContents.XML, "</toast>" & vbCrLf, "</toast>")

End Function



'***************************************************************************************************
' ■■■ メンバ変数を元に、設定用の変数コマンド文字列を生成 ■■■
'***************************************************************************************************
'* 機能　　：コマンドプロンプト(shell関数など)で認識できるように、xmlコンテンツを整形(エスケープ)し、それをセットするコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：エスケープ後のトーストスキーマコンテンツ(xml)
'* 引数　　：元のトーストスキーマコンテンツ(xml)
'***************************************************************************************************
Private Function SetToastShemaXml(XmlContents As String) As String
    'コマンドプロンプトで実行するため、以下の文字列をエスケープしてます
    '  「"」→「\"」
    '  「'」→「''」
    '  ちなみに、改行コードが直接入っても、shell関数なら正しく動く模様
    SetToastShemaXml = "$xml = '" & Replace(Replace(XmlContents, Chr(34), "\"""), Chr(39), "''") & Chr(39)
End Function

'***************************************************************************************************
'* 機能　　：どのアプリから通知を発行するかを設定するAppIDを設定するコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：AppIDを設定するコマンド文字列
'* 引数　　：UseDll
'               True ：カスタマイズしたDllで処理する際に指定。AppIDをそのまま返します
'               False：PowerShellコマンドで実行する際に指定します。変数格納用の「$AppId」が追加されます(規定)
'---------------------------------------------------------------------------------------------------
'* 機能説明：ネット上の画像も使用して通知を表示するかを設定します
'* 注意事項：厳密には、Settingがc:internetを含むAppIDでないとインターネット上の画像が使えない
'**************************************************************************************************
Private Function SetValueAppUserModelID(Optional UseDll As Boolean) As String
    '任意のAppIDの設定があったら、そちらを優先します。
    If StrPtr(AppID_Any) Then
        ToastConfig.AppUserModelID = AppID_Any

    '未設定の場合は、フラグ値に基づいて、プリセットAppIDをセットします。
    ElseIf ToastConfig.AllowUse_InternetImageFlag Then
        ToastConfig.AppUserModelID = AppID_Microsoft365
    Else
        ToastConfig.AppUserModelID = AppID_Excel
    End If



    'AppIDを決めるコマンド文字列を返す
    If UseDll Then
        SetValueAppUserModelID = ToastConfig.AppUserModelID
    Else
        SetValueAppUserModelID = "$AppId = '" & ToastConfig.AppUserModelID & Chr(39)
    End If

End Function

'***************************************************************************************************
'* 機能　　：「System.Collections.Generic.Dictionary」を利用して、可変値部分を設定、更新を行うコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：UpdateMode
'               True    更新時は、これに設定
'               False   最初に進捗Toastを発行する場合は、これを設定
'               ※変数の格納先を変更する都合上、このフラグが必要です。
'***************************************************************************************************
Private Function Binding_AdaptiveProgressBar(UpdateMode As Boolean) As String
    '「System.Collections.Generic.Dictionary」というオブジェクトを定義するコマンド文字列
    Const New__System_Collections_Generic__Dictionary As String = "$Dictionary = [System.Collections.Generic.Dictionary[String, String]]::New()"

    '事前に定義した「$Dictionary」から、トーストの進行状況バーの可変値に、パラメーターを設定します
    '→https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.data
    Dim Set__Windows_UI_Notifications__ToastNotificationData As String
    If UpdateMode Then
        Set__Windows_UI_Notifications__ToastNotificationData = "$ToastNotificationData = [Windows.UI.Notifications.NotificationData]::New($Dictionary)"
    Else
        Set__Windows_UI_Notifications__ToastNotificationData = "$ToastNotification.Data = [Windows.UI.Notifications.NotificationData]::New($Dictionary)"
    End If


    With ElementProgress
        'タイトル属性を設定するコマンド文字列
        Dim SetProgressBar_Title               As String: SetProgressBar_Title = "$Dictionary.Add('progressTitle', '" & .Title & "')"

        '進行状況バーの値属性を設定するコマンド文字列
        '※0未満になってると、「Indeterminate」扱いとなります。
        Dim SetProgressBar_Value               As String: SetProgressBar_Value = "$Dictionary.Add('progressValue', '" & .Value & "')"

        '割合を示す文字列属性を設定するコマンド文字列
        '  Null文字として渡された場合、「progress」要素に「valueStringOverride」属性自体を挿入しないようにします。これにより規定の"70%" などの文字が表示されます。
        '  ただし、更新時に改めて、「valueStringOverride」属性を指定しても、無視されます。
        Dim SetProgressBar_ValueStringOverride As String
        If StrPtr(.ValueStringOverride) Then SetProgressBar_ValueStringOverride = "$Dictionary.Add('progressValueString', '" & .ValueStringOverride & "')"

        '進行状況バーの下の左側に表示されるステータス文字列属性を設定するコマンド文字列
        Dim SetProgressBar_Status              As String: SetProgressBar_Status = "$Dictionary.Add('progressStatus', '" & .Status & "')"
    End With


    '「$Dictionary」をAddするコマンド文字列を返す
    Binding_AdaptiveProgressBar = WorksheetFunction.TextJoin(";", True, New__System_Collections_Generic__Dictionary, SetProgressBar_Title, SetProgressBar_Value, SetProgressBar_ValueStringOverride, SetProgressBar_Status, Set__Windows_UI_Notifications__ToastNotificationData)
End Function




'***************************************************************************************************
' ■■■ [Windows.UI.Notifications]に関する動的設定 ■■■
'***************************************************************************************************
'* 機能　　：トースト通知のコンテンツ、関連するメタデータとイベント、有効期限を定義します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に応じた[Windows.UI.Notifications.ScheduledToastNotification]あるいは、[Windows.UI.Notifications.ToastNotification]の設定コマンド文字列
'* 引数　　：ArgScheduleDate    Toastを出す日時
'            ArgScheduleID      特定のスケジュールされたトーストを識別するために使用される開発者指定の値

'---------------------------------------------------------------------------------------------------
'* 機能説明：引数があると、スケジュールモードとして扱います(Windows.UI.Notifications.ScheduledToastNotification)。そうじゃない場合は、通常の通知(Windows.UI.Notifications.ToastNotification)
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification
'            https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification
'***************************************************************************************************
Private Function CreateObject__Windows_UI_Notifications(ArgScheduleDate As Date, ArgScheduleID As String) As String
    '引数が有効なら、スケジュールとしてClassを定義します
    If ArgScheduleDate Then
        'スケジュールを設定しているが過去になっている場合、VBAエラーを起こして、処理を停止させます
        If Now() >= ArgScheduleDate Then Err.Raise Number:=513, description:="設定したスケジュールは、現在のシステム日時に対して過去の日時に設定されます。" & vbCrLf & "設定したスケジュール通知日時が、将来であることを確認してください。 "

        'スケジュールされた時刻に表示されるトースト通知を定義する XML が含まれます。更に識別子も付与します
        '→https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification
        CreateObject__Windows_UI_Notifications = WorksheetFunction.TextJoin(";", False, _
            "$ToastNotification = [Windows.UI.Notifications.ScheduledToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime]::New($XmlDocument,'" & WorksheetFunction.Text(ArgScheduleDate, "yyyy/mm/dd hh:mm:ss") & "')", _
            "$ToastNotification.id = '" & ArgScheduleID & "'")
    Else
        'トースト通知のコンテンツ、関連するメタデータとイベント、有効期限を定義します。
        '→https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification
        CreateObject__Windows_UI_Notifications = "$ToastNotification = [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime]::New($XmlDocument)"
    End If
End Function

'***************************************************************************************************
'* 機能　　：この通知の一意識別子を設定するコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：通知の一意識別子を設定するコマンド文字列
'* 引数　　：ArgTag     通知の一意識別子
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.tag
'***************************************************************************************************
Private Function Set__Windows_UI_Notifications__ToastNotificationTag(ArgTag As String) As String
    Set__Windows_UI_Notifications__ToastNotificationTag = "$ToastNotification.Tag = '" & ArgTag & Chr(39)
End Function

'***************************************************************************************************
'* 機能　　：この通知のグループ内の一意識別子を設定するコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：通知のグループ内の一意識別子を設定するコマンド文字列
'* 引数　　：ArgGroup     Toastのグループ一意識別子を設定する値
'---------------------------------------------------------------------------------------------------
'* 注意事項：引数は、このExcelファイル名にすることを推奨
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.group
'***************************************************************************************************
Private Function Set__Windows_UI_Notifications__ToastNotificationGroup(ArgGroup As String) As String
    Set__Windows_UI_Notifications__ToastNotificationGroup = "$ToastNotification.Group = '" & ArgGroup & Chr(39)
End Function

'***************************************************************************************************
'* 機能　　：通知の有効期限を設定します。基本的に、スケジュール通知用のプロパティです。通常の通知でも使えます。
'---------------------------------------------------------------------------------------------------
'* 返り値　：通知の有効期限あるいは、再起動時通知消滅フラグを設定するコマンド文字列
'* 引数　　：ArgExpirationDate      アクションセンターから自動削除する日時。負にすると、コンピューターの再起動時に通知の有効期限が切れる設定になります
'            ArgScheduleDate        チェック用。「コンピューターの再起動時に通知の有効期限が切れる設定」は、Scheduleモードでは動作しないため、そのチェック用となります

'---------------------------------------------------------------------------------------------------
'* 注意事項：有効期限が切れても、10~30sの誤差が発生します。
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.expirationtime
'            https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.expiresonreboot
'***************************************************************************************************
Private Function Set__Windows_UI_Notifications__ToastNotificationExpirationTime(ArgExpirationDate As Date, ArgScheduleDate As Date) As String
    '有効期限が負かつ、ArgScheduleDateが未設定の場合、コンピューターの再起動時に通知の有効期限が切れる設定のコマンド文字列を返す
    If ArgExpirationDate < 0 And ArgScheduleDate = 0 Then
        Set__Windows_UI_Notifications__ToastNotificationExpirationTime = "$ToastNotification.ExpiresOnReboot = 1"

    '有効な有効期限なら、それを設定
    ElseIf ArgExpirationDate Then
        Set__Windows_UI_Notifications__ToastNotificationExpirationTime = "$ToastNotification.ExpirationTime = '" & WorksheetFunction.Text(ArgExpirationDate, "yyyy/mm/dd hh:mm:ss") & Chr(39)
    End If
End Function

'***************************************************************************************************
'* 機能　　：トーストのポップアップ UI をユーザーの画面に表示するかどうかを設定します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：通知のポップアップを設定するコマンド文字列
'* 引数　　：ArgHideFlag
'               True    ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'               False   ポップアップ メッセージを表示し、トースト通知がアクション センターにも通知されます。
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.suppresspopup
'***************************************************************************************************
Private Function Set__Windows_UI_Notifications__ToastNotificationSuppressPopup(ArgSuppress As Boolean) As String
    If ArgSuppress Then Set__Windows_UI_Notifications__ToastNotificationSuppressPopup = "$ToastNotification.SuppressPopup = " & ArgSuppress * -1
End Function

'***************************************************************************************************
'* 機能　　：更新対象のToastTagを指定して、通知内容を更新するコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：更新対象のToastTagを指定するコマンド文字列
'* 引数　　：ArgGroup   通知内容を更新させるToastGroup
'            ArgTag     通知内容を更新させるToastGroup内のToastTag
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotifier.update
'***************************************************************************************************
Private Function Run__Windows_UI_Notifications__ToastNotifierUpdate(ArgTag As String, ArgGroup As String) As String
    Run__Windows_UI_Notifications__ToastNotifierUpdate = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::CreateToastNotifier($AppId).Update($ToastNotificationData, '" & WorksheetFunction.TextJoin("','", True, ArgTag, ArgGroup) & "')"
End Function




'***************************************************************************************************
'* 機能　　：スケジュールの有無に応じて実行メソッドを変えます
'---------------------------------------------------------------------------------------------------
'* 返り値　：通知を発行するコマンド文字列
'* 引数　　：ArgSchedule    通知発行予定日
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotifier.show
'            https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotifier.addtoschedule
'            https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationmanager.createtoastnotifier
'***************************************************************************************************
Private Function Run__Windows_UI_Notifications__ToastNotifier(ByVal ArgSchedule As Boolean) As String
    'ToastNotificationManager.CreateToastNotifier メソッドをPowershellで定義するコマンド文字列
    '→https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationmanager.createtoastnotifier
    Const Def_CreateToastNotifier As String = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::CreateToastNotifier($AppId)"

    'スケジュールの有無に応じて、ToastNotifier  のメソッドを適切に割り当てます
    If ArgSchedule Then
        'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotifier.addtoschedule
        Run__Windows_UI_Notifications__ToastNotifier = Def_CreateToastNotifier & ".addToSchedule($ToastNotification)"
    Else
        'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotifier.show
        Run__Windows_UI_Notifications__ToastNotifier = Def_CreateToastNotifier & ".Show($ToastNotification)"
    End If
End Function

'***************************************************************************************************
'* 機能　　：コレクションに送信し、Toast表示するまでの、一式処理です。
'---------------------------------------------------------------------------------------------------
'* 返り値　：コレクションに送信し、Toast表示するコマンド文字列
'* 引数　　：CollectionID   通知のコレクションID
'***************************************************************************************************
Private Function Run__Windows_UI_Notifications__ToastNotifierShow_Collection(CollectionID As String) As String
    '指定した通知グループのバインドされたユーザーとアプリの新しい ToastNotification を作成します。
    Dim IAsyncOperation_ToastNotifier As String
    IAsyncOperation_ToastNotifier = "$WinRtTask = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::GetDefault().GetToastNotifierForToastCollectionIdAsync('" & CollectionID & "');$ResultType = [Windows.UI.Notifications.ToastNotifier]"

    'await部分を擬似的に処理
    '→https://qiita.com/7shi/items/e8de755a5738f862e09f
    Const EmulateAwait As String = "Add-Type -AssemblyName System.Runtime.WindowsRuntime;" & _
        "$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0];" & _
        "$asTask = $asTaskGeneric.MakeGenericMethod($ResultType);" & _
        "$netTask = $asTask.Invoke($null, @($WinRtTask))"

    '上記で返却されたオブジェクトを使って、トーストを表示
    Const Show As String = "$ToastNotifier = $netTask.Result;$ToastNotifier.Show($ToastNotification)"



    '上記でコマンド文字列を生成
    Run__Windows_UI_Notifications__ToastNotifierShow_Collection = WorksheetFunction.TextJoin(";", False, IAsyncOperation_ToastNotifier, EmulateAwait, Show)
End Function



'***************************************************************************************************
' ■■■ Windows API mciSendString  に関するメソッド ■■■
'***************************************************************************************************
'* 機能　　：トーストの通知音を自分好みの音源ファイルに指定して鳴らす
'---------------------------------------------------------------------------------------------------
'* 引数　　：SoundURI   音源ファイルパス(ローカルパスはもちろん、インターネット上も指定可能です)
'---------------------------------------------------------------------------------------------------
'* 注意事項：※事前に、トースト自体の通知音をミュートにし、トーストを呼び出す後に実行すること。トーストの通知音を無音にして、このメソッドを呼ぶことで擬似的に指定した通知音を鳴らしています。
'***************************************************************************************************
Public Property Let CustomNotificationSoundPlay(ByVal SoundURI As String)
    'メモリの蓄積を防ぐため、最初にClose処理を施す
    mciSendString "Close """ & SoundURI & Chr(34), "", 0, 0

    '開く
    mciSendString "Open """ & SoundURI & Chr(34), "", 0, 0

    '再生
    mciSendString "Play """ & SoundURI & Chr(34), "", 0, 0
End Property



'***************************************************************************************************
' ■■■ Toast表示させるコマンド文字列を返すメソッド ■■■
'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastNotifierShow
'* 機能　　：引数に渡された値で、単純なトースト通知を表示するコマンド文字列を返します。指定日時に通知するスケジュール機能も対応します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、トーストを表示させるコマンド文字列を返します
'* 引数　　：ToastTag       ToastTagを指定します。複数通知を提供する際に必要です
'            CollectionID   トーストコレクションIDがあるなら、それを指定
'            ScheduleDate   指定日時に通知したい場合は、「Date型」として指定することでスケジュール通知として登録されます。省略時、即時通知とします
'            ExpirationDate アクションセンターから自動削除される日時を指定。最大30sの遅延バグが発生します。
'            Suppress       「True」で、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'---------------------------------------------------------------------------------------------------
'* 注意事項：あくまで「コマンド実行の文字列」を返すだけで、呼び出すだけでは実行されません。別途、Shell関数や「Windows Script Host Object Model」を参照設定してRUNコマンドを実行することで、Toastを発行できます。
'**************************************************************************************************
Public Property Get GenerateCmd_ToastNotifierShow(ByVal ToastTag As String, Optional ByVal CollectionID As String, Optional ByVal ScheduleDate As Date, Optional ByVal ExpirationDate As Date, Optional ByVal Suppress As Boolean) As String
    With ToastConfig
        '設定値をまとめる
        '.AppUserModelID = SetValueAppUserModelID(False)  '後続で、代入
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(False)
        
        .Tag = ToastTag
        .Group = ThisWorkbook.Name
        .Schedule_ID = "ExcelSchedule"
        .CollectionID = CollectionID
        .ExpiresOnReboot = (ExpirationDate < 0)
        .SuppressPopup = Suppress
        .Schedule_DeliveryTime = ScheduleDate
        .ExpirationTime = ExpirationDate
        
        
        
        '共通の設定コマンドを生成
        ' 1.事前に設定したメンバ変数から、xmlコンテンツを生成し、それを読み込む。
        ' 2.読み込んだxmlコンテンツから、ToastNotificationの構造を設定
        ' 3.サイレント通知フラグを設定
        ' 4.有効期限を設定
        ' 5.識別Group,Tagを設定
        GenerateCmd_ToastNotifierShow = ActionPS & WorksheetFunction.TextJoin(";", True, _
            SetToastShemaXml(.XmlSchema), ReadXml, _
            CreateObject__Windows_UI_Notifications(.Schedule_DeliveryTime, .Schedule_ID), _
            Set__Windows_UI_Notifications__ToastNotificationSuppressPopup(.SuppressPopup), _
            Set__Windows_UI_Notifications__ToastNotificationExpirationTime(.ExpirationTime, .Schedule_DeliveryTime), _
            Set__Windows_UI_Notifications__ToastNotificationGroup(.Group), Set__Windows_UI_Notifications__ToastNotificationTag(.Tag))


        'コレクションIDの有無に応じで、設定コマンドを切り替える
        If StrPtr(CollectionID) Then
            '6.コレクションIDを指定して、Toast通知実行
            GenerateCmd_ToastNotifierShow = WorksheetFunction.TextJoin(";", False, GenerateCmd_ToastNotifierShow, _
                Run__Windows_UI_Notifications__ToastNotifierShow_Collection(.CollectionID)) & Chr(34)
        Else
            '6.AppIDを設定
            '7.Toast通知実行
            GenerateCmd_ToastNotifierShow = WorksheetFunction.TextJoin(";", False, GenerateCmd_ToastNotifierShow, _
                SetValueAppUserModelID, _
                Run__Windows_UI_Notifications__ToastNotifier(.Schedule_DeliveryTime)) & Chr(34)
        End If
    End With
End Property

'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastNotifierShow_Progress
'* 機能　　：引数に渡された値で、トーストの進行状況バーを表示するコマンド文字列を返します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、進捗状況ゲージ込みのトーストを表示させるコマンド文字列を返します
'* 引数　　：ToastTag                   ToastTagを指定します。以後の進捗状況を更新する際に必要です
'            Arg_Status                 進行状況バーの下の左側に表示されるステータス文字列 (必須) を設定します。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります
'            Arg_Value                  進行状況バーの値を取得または設定します。 既定値は 0 です。 0.0 ～ 1.0 の double型とします。負の値を指定すると、「Indeterminate」と扱います。
'            Arg_Title                  タイトルの文字列 (オプション) を設定します
'            Arg_ValueStringOverride    割合を示す既定の文字列に代わって表示される文字列 (オプション) を設定します。 これを指定しない場合は、"70%" などの文字が表示されます。
'            Suppress                   「True」で、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：初めて、トーストの進行状況バーを表示するときはまずこちらを使います。
'***************************************************************************************************
Public Property Get GenerateCmd_ToastNotifierShow_Progress(ByVal ToastTag As String, ByVal Arg_Status As String, Optional Arg_Value As Double = 0, Optional ByVal Arg_Title As String, Optional ByVal Arg_ValueStringOverride As String, Optional ByVal ArgSuppress As Boolean) As String
    '進捗用ユーザー定義型のメンバー変数にセット
    SetToastContent_ProgressBar Arg_Title, Arg_Value, Arg_ValueStringOverride, Arg_Status

    With ToastConfig
        '設定値をまとめる
        '.AppUserModelID = SetValueAppUserModelID(False)  '後続で、代入
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(True)

        .Tag = ToastTag
        .Group = ThisWorkbook.Name
        .ExpiresOnReboot = False
        .Schedule_ID = vbNullString
        .CollectionID = vbNullString
        
        .ExpiresOnReboot = False
        .SuppressPopup = ArgSuppress
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False



        'トーストの進行状況バーを表示させるコマンド文字列を生成
        ' 1.事前に設定したメンバ変数から、xmlコンテンツを生成し、それを読み込む。
        ' 2.読み込んだxmlコンテンツから、ToastNotificationの構造を設定
        ' 3.識別Group,Tagを設定
        ' 4.サイレント通知フラグを設定
        ' 5.進捗Toastパラメーターを設定
        ' 6.AppIDを設定
        ' 7.Toast通知実行
        GenerateCmd_ToastNotifierShow_Progress = ActionPS & WorksheetFunction.TextJoin(";", True, _
            SetToastShemaXml(.XmlSchema), ReadXml, _
            CreateObject__Windows_UI_Notifications(False, vbNullString), _
            Set__Windows_UI_Notifications__ToastNotificationGroup(.Group), Set__Windows_UI_Notifications__ToastNotificationTag(.Tag), _
            Set__Windows_UI_Notifications__ToastNotificationSuppressPopup(.SuppressPopup), _
            Binding_AdaptiveProgressBar(False), _
            SetValueAppUserModelID, _
            Run__Windows_UI_Notifications__ToastNotifier(False)) & Chr(34)
    
    End With
End Property

'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastNotifierUpdate_Progress
'* 機能　　：引数に渡された値で、トーストの進行状況バーを更新します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、既存の進捗状況ゲージ込みのトーストを更新させるコマンド文字列を返します
'* 引数　　：ToastTag                   ToastTagを指定します。以後の進捗状況を更新する際に必要です
'            Arg_Status                 進行状況バーの下の左側に表示されるステータス文字列 (必須) を設定します。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります
'            Arg_Value                  進行状況バーの値を取得または設定します。 既定値は 0 です。 0.0 ～ 1.0 の double型とします。負の値を指定すると、「Indeterminate」と扱います。
'            Arg_Title                  タイトルの文字列 (オプション) を設定します
'            Arg_ValueStringOverride    割合を示す既定の文字列に代わって表示される文字列 (オプション) を設定します。 これを指定しない場合は、"70%" などの文字が表示されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：既にトーストの進行状況バーを表示しているとき、その時に指定したタグを指定することで、更新が可能です。
'***************************************************************************************************
Public Property Get GenerateCmd_ToastNotifierUpdate_Progress(ByVal ToastTag As String, ByVal Arg_Status As String, Optional ByVal Arg_Value As Double = 0, Optional ByVal Arg_Title As String, Optional ByVal Arg_ValueStringOverride As String) As String
    '進捗用ユーザー定義型のメンバー変数にセット
    SetToastContent_ProgressBar Arg_Title, Arg_Value, Arg_ValueStringOverride, Arg_Status

    With ToastConfig
        '設定値をまとめる
        '.AppUserModelID = SetValueAppUserModelID(False)  '後続で、代入
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

        .Tag = ToastTag
        .Group = ThisWorkbook.Name
        .Schedule_ID = vbNullString
        .CollectionID = vbNullString
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
       


        'トーストの進行状況バーを更新させるコマンド文字列を生成
        ' 1.進捗Toastを更新するパラメータを設定
        ' 2.AppIDを設定
        ' 3.更新用Toast通知実行
        GenerateCmd_ToastNotifierUpdate_Progress = ActionPS & WorksheetFunction.TextJoin(";", True, _
            Binding_AdaptiveProgressBar(True), _
            SetValueAppUserModelID, _
            Run__Windows_UI_Notifications__ToastNotifierUpdate(.Tag, .Group)) & Chr(34)
    
    End With
End Property

'***************************************************************************************************
'* 処理名　：Get GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync
'* 機能　　：引数で渡された値で、コレクションを使用したトースト通知のグループ化を作成するコマンド文字列を返します。この段階ではまだ(疑似的に)、作成のみです。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、既存の進捗状況ゲージ込みのトーストを更新させるコマンド文字列を返します
'* 引数　　：CollectionId   グループの ID。 ID を使用して、ToastCollectionManager を使用してグループを操作できます
'            DisplayName    アクション センターに表示するグループのタイトル
'            LaunchArgs     ユーザーがアクション センターで通知グループのタイトルをクリックしたときにアプリに渡される起動引数。
'            IconUri        アクション センターのタイトルの横に表示するアイコン。
'            FixMode        Trueで、有効期限切れのToastの発行処理を含めるコマンド文字列を生成させます。これを入れないと動作しないWindows バージョンがあるそうなのでそれ用です
'---------------------------------------------------------------------------------------------------
'* 注意事項：Desktopアプリだと、「CollectionId」が「displayName」になる問題があるため、Storeアプリでの使用を推奨します。
'* 機能説明：コレクション作成→トースト発行 を同プロセスで実行しないとうまく動作しないため､有効期限を最古の値に設定して擬似的に作成のみの処理にしています｡
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-collections
'***************************************************************************************************
Public Property Get GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync(CollectionID As String, DisplayName As String, LaunchArgs As String, IconUri As String, Optional FixMode As Boolean) As String
    With ToastConfig
        '設定値をまとめる
        '.AppUserModelID = SetValueAppUserModelID(False)  '後続で、代入
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(False)

        .Tag = vbNullString
        .Group = vbNullString
        .Schedule_ID = vbNullString
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = #1/1/1900#
        
        
        
        '引数から、新しいトースト通知グループを作成します。
        Dim New__Windows_UI_Notifications__ToastCollection As String
        New__Windows_UI_Notifications__ToastCollection = "$ToastCollection = [Windows.UI.Notifications.ToastCollection, Windows.UI.Notifications, ContentType = WindowsRuntime]::New('" & WorksheetFunction.TextJoin("','", False, CollectionID, DisplayName, LaunchArgs, IconUri) & "')"

        '指定したトースト通知グループの非同期作成または更新を開始します。
        Const Run__Windows_UI_Notifications__ToastCollectionManager_SaveToastCollectionAsync As String = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::GetDefault().GetToastCollectionManager($AppId).SaveToastCollectionAsync($ToastCollection)"



        '共通の設定コマンド文字列を生成
        'コレクションを使用したトースト通知のグループ化を作成するコマンド文字列を生成。最後の方でToast通知実行しますが、有効期限の設定上、表示はされず結果的に作成のみの動作っぽくなります。こうしないとうまくいかないパターンがあるためです。
        ' 1.新しいトースト通知グループを作成
        ' 2.AppIDを設定
        ' 3.指定したトースト通知グループを作成
        GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync = ActionPS & WorksheetFunction.TextJoin(";", True, _
            New__Windows_UI_Notifications__ToastCollection, _
            SetValueAppUserModelID, _
            Run__Windows_UI_Notifications__ToastCollectionManager_SaveToastCollectionAsync)

        'true  で不具合対応板として、動作させます。
        '最後の方でToast通知実行しますが、有効期限の設定上、表示はされず結果的に作成のみの動作っぽくなります。こうしないとうまくいかないパターンがあるためです。
        If FixMode Then
            ' 4.事前に設定したメンバ変数から、xmlコンテンツを生成し、それを読み込む。
            ' 5.読み込んだxmlコンテンツから、ToastNotificationの構造を設定
            ' 6.有効期限をExcelで扱える最低値にします
            ' 7.Toast通知実行しますが、先ほどの設定により、表示はされません。
            GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync = WorksheetFunction.TextJoin(";", True, GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync, _
                SetToastShemaXml(.XmlSchema), ReadXml, _
                CreateObject__Windows_UI_Notifications(.Schedule_DeliveryTime, .Schedule_ID), _
                Set__Windows_UI_Notifications__ToastNotificationExpirationTime(.ExpirationTime, .Schedule_DeliveryTime), _
                Run__Windows_UI_Notifications__ToastNotifierShow_Collection(.CollectionID)) & Chr(34)
        Else
            'パッチコマンドなしで、返します。
            GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync = GenerateCmd_ToastCollectionManagerSaveToastCollectionAsync & Chr(34)
        End If
    
    End With
End Property

'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastCollectionManagerRemoveToastCollectionAsync
'* 機能　　：AppIDに紐づく、指定Toastコレクションを削除します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：AppIDに紐づく、指定Toastコレクションを削除するコマンド文字列
'* 引数　　：CollectionId   コレクションID。省略すると、現在のアプリIDのToastコレクションを全て削除します
'---------------------------------------------------------------------------------------------------
'* 注意事項：無闇にToastコレクションを増やすと  設定→システム→通知  の項目が増えてしまうので、これを実行して整理して下さい。
'***************************************************************************************************
Public Property Get GenerateCmd_ToastCollectionManagerRemoveToastCollectionAsync(Optional CollectionID As String) As String
    '全削除用コマンド文字列
    Const Run__Windows_UI_Notifications__ToastCollectionManager_RemoveAllToastCollections As String = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::GetDefault().GetToastCollectionManager($AppId).RemoveAllToastCollectionsAsync()"

    '指定ToastコレクションIDの削除
    Dim Run__Windows_UI_Notifications__ToastCollectionManager_RemoveToastCollections As String: Run__Windows_UI_Notifications__ToastCollectionManager_RemoveToastCollections = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::GetDefault().GetToastCollectionManager($AppId).RemoveToastCollectionAsync('" & CollectionID & "')"


    '引数の有無に応じて、削除モードを切り替えます
    Dim RemoveMode As String
    If StrPtr(CollectionID) Then RemoveMode = Run__Windows_UI_Notifications__ToastCollectionManager_RemoveToastCollections Else RemoveMode = Run__Windows_UI_Notifications__ToastCollectionManager_RemoveAllToastCollections


    '削除コマンド文字列を生成します
    GenerateCmd_ToastCollectionManagerRemoveToastCollectionAsync = ActionPS & WorksheetFunction.TextJoin(";", False, _
        SetValueAppUserModelID, _
        RemoveMode) & Chr(34)
End Property

'***************************************************************************************************
'* 処理名　：PresetReminder
'* 機能　　：リマインダー機能をプリセットします。システム側で用意されているパラメーターをセットすることで、システムの再通知と無視アクションを利用するアラーム通知を作成できます。
'---------------------------------------------------------------------------------------------------
'* 引数　　：RemindMinute1-5    再通知する時間を分で指定
'            Message            タイトル

'* 注意事項：このプロシージャを呼び出す前に定義した下記の要素は、上書きされますのでご注意ください
'           ・1,2つ目のaction要素
'           ・1つ目のinput要素

'* URL 　　：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#snoozedismiss
'***************************************************************************************************
Public Sub PresetReminder(RemindMinute1 As LongPtr, Optional RemindMinute2 As LongPtr, Optional RemindMinute3 As LongPtr, Optional RemindMinute4 As LongPtr, Optional RemindMinute5 As LongPtr, Optional Message As String)
    '紐付け用識別子(なんでもOK)
    Const ReminderID As String = "snoozeTime"


    '----------------------------1.各Selection要素を用意----------------------------
    '選択肢を用意
    SetToastSelectionBox(RemindMinute1, RemindMinute1 & " 分後") = 1
    SetToastSelectionBox(RemindMinute2, RemindMinute2 & " 分後") = 2
    SetToastSelectionBox(RemindMinute3, RemindMinute3 & " 分後") = 3
    SetToastSelectionBox(RemindMinute4, RemindMinute4 & " 分後") = 4
    SetToastSelectionBox(RemindMinute5, RemindMinute5 & " 分後") = 5


    '----------------------------2.1 で用意したSelection要素をinput要素を作成する際のIndex値を指定し、Type=True でset----------------------------
    '選択肢を1つ目のinput要素に格納
    SetIToastInput(ReminderID, True, , Message, RemindMinute1) = 1


    '----------------------------3.1,2つ目のaction要素に、それぞれプリセット----------------------------
    SetIToastActions("", "snooze", "system", , , , ReminderID) = 1
    SetIToastActions("", "dismiss", "system") = 2
End Sub



'***************************************************************************************************
' ■■■ 専用DLL C++からWindows.UI.Notificationsを呼び出すメソッド ■■■
'***************************************************************************************************
'* 処理名　：RunDll_ToastNotifierShow
'* 機能　　：引数に渡された値で、単純なトースト通知を表示します。指定日時に通知するスケジュール機能も対応します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、トーストを表示させるコマンド文字列を返します
'* 引数　　：ToastTag           ToastTagを指定します。複数通知を提供する際に必要です
'            CollectionID       トーストコレクションIDがあるなら、それを指定
'            ScheduleDate       指定日時に通知したい場合は、「Date型」として指定することでスケジュール通知として登録されます。省略時、即時通知とします
'            ExpirationDate     アクションセンターから自動削除される日時を指定。最大30sの遅延バグが発生します。
'            Suppress           「True」で、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'---------------------------------------------------------------------------------------------------
'* 注意事項：これを実行すると、即座にDll内で処理を行います。また、日時系データを扱う場合は、UTC基準になるので注意
'**************************************************************************************************
Public Sub RunDll_ToastNotifierShow(ByVal ToastTag As String, Optional ByVal CollectionID As String, Optional ByVal ScheduleDate As Date, Optional ByVal ExpirationDate As Date, Optional ByVal Suppress As Boolean)
    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = SetValueAppUserModelID(True)
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(False)
        
        .Tag = ToastTag
        .Group = ThisWorkbook.Name
        .Schedule_ID = "ExcelSchedule"
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = (ExpirationDate < 0)
        .SuppressPopup = Suppress
        
        .Schedule_DeliveryTime = ScheduleDate
        .ExpirationTime = ExpirationDate
    End With


    'Dll内の「ShowToastNotification」関数を実行します
    ShowToastNotification ParamsConvertForDLL(ToastConfig)
End Sub

'***************************************************************************************************
'* 処理名　：RunDll_ToastNotifierShow_Progress
'* 機能　　：引数に渡された値で、トーストの進行状況バーを表示します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、進捗状況ゲージ込みのトーストを表示させるコマンド文字列を返します
'* 引数　　：ToastTag                   ToastTagを指定します。以後の進捗状況を更新する際に必要です
'            Arg_Status                 進行状況バーの下の左側に表示されるステータス文字列 (必須) を設定します。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります
'            Arg_Value                  進行状況バーの値を取得または設定します。 既定値は 0 です。 0.0 ～ 1.0 の double型とします。負の値を指定すると、「Indeterminate」と扱います。
'            Arg_Title                  タイトルの文字列 (オプション) を設定します
'            Arg_ValueStringOverride    割合を示す既定の文字列に代わって表示される文字列 (オプション) を設定します。 これを指定しない場合は、"70%" などの文字が表示されます。
'            Suppress                   「True」で、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：初めて、トーストの進行状況バーを表示するときはまずこちらを使います。
'* 注意事項：これを実行すると、即座にDll内で処理を行います。また、プログレスバーの特性上、スケジュールトースト、有効期限(既定の有効期限と最大有効期限は3日間)の任意設定は設けません。
'***************************************************************************************************
Sub RunDll_ToastNotifierShow_Progress(ByVal ToastTag As String, ByVal Arg_Status As String, Optional Arg_Value As Double = 0, Optional ByVal Arg_Title As String = "", Optional ByVal Arg_ValueStringOverride As String, Optional ByVal ArgSuppress As Boolean)
    '進捗用ユーザー定義型のメンバー変数にセット
    SetToastContent_ProgressBar Arg_Title, Arg_Value, Arg_ValueStringOverride, Arg_Status

    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = SetValueAppUserModelID(True)
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(True)

        .Tag = ToastTag
        .Group = ThisWorkbook.Name
        .ExpiresOnReboot = False
        .Schedule_ID = vbNullString
        .CollectionID = vbNullString
        
        .ExpiresOnReboot = False
        .SuppressPopup = ArgSuppress
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With


    'Dll内の「ShowToastNotificationWithProgressBar」関数を実行します
    ShowToastNotificationWithProgressBar ParamsConvertForDLL(ToastConfig), StrPtr(Arg_Status), Arg_Value, StrPtr(Arg_Title), StrPtr(Arg_ValueStringOverride)
End Sub

'***************************************************************************************************
'* 処理名　：RunDll_ToastNotifierUpdate_Progress
'* 機能　　：引数に渡された値で、トーストの進行状況バーを更新します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.notificationupdateresult
'            0(Succeeded)               通知が更新されました。
'            1(Failed)                  通知の更新に失敗しました。
'            2(NotificationNotFound)    指定した通知が見つかりませんでした。
'            ※上記以外                  予期せぬエラー

'* 引数　　：ToastTag                   ToastTagを指定します。以後の進捗状況を更新する際に必要です
'            Arg_Status                 進行状況バーの下の左側に表示されるステータス文字列 (必須) を設定します。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります
'            Arg_Value                  進行状況バーの値を取得または設定します。 既定値は 0 です。 0.0 ～ 1.0 の double型とします。負の値を指定すると、「Indeterminate」と扱います。
'            Arg_Title                  タイトルの文字列 (オプション) を設定します
'            Arg_ValueStringOverride    割合を示す既定の文字列に代わって表示される文字列 (オプション) を設定します。 これを指定しない場合は、"70%" などの文字が表示されます。
'            SequenceNumber             順序外の更新を防ぐための、シーケンス番号を指定します。
'---------------------------------------------------------------------------------------------------
'* 機能説明：既にトーストの進行状況バーを表示しているとき、その時に指定したタグを指定することで、更新が可能です。Dllの場合、返り値に対応し、これに応じた分岐処理も可能です。
'***************************************************************************************************
Public Property Get RunDll_ToastNotifierUpdate_Progress(ByVal ToastTag As String, ByVal Arg_Status As String, Optional ByVal Arg_Value As Double = 0, Optional ByVal Arg_Title As String = "", Optional ByVal Arg_ValueStringOverride As String, Optional SequenceNumber As Long = 2) As Long
    '進捗用ユーザー定義型のメンバー変数にセット
    SetToastContent_ProgressBar Arg_Title, Arg_Value, Arg_ValueStringOverride, Arg_Status

    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = SetValueAppUserModelID(True)
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

        .Tag = ToastTag
        .Group = ThisWorkbook.Name
        .Schedule_ID = vbNullString
        .CollectionID = vbNullString
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With
     
     
    'Dll内の「UpdateToastNotificationWithProgressBar」関数を実行します
    RunDll_ToastNotifierUpdate_Progress = UpdateToastNotificationWithProgressBar(ParamsConvertForDLL(ToastConfig), StrPtr(Arg_Status), Arg_Value, StrPtr(Arg_Title), StrPtr(Arg_ValueStringOverride), SequenceNumber)
End Property

'***************************************************************************************************
'* 処理名　：RunDll_ToastCollectionManagerSaveToastCollectionAsync
'* 機能　　：引数で渡された値で、コレクションを使用したトースト通知のグループ化を作成します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、コレクションを使用したトースト通知のグループ化を作成するコマンド文字列を返します
'* 引数　　：CollectionId   コレクションの ID。 ID を使用して、ToastCollectionManager を使用してグループを操作できます
'            DisplayName    アクション センターに表示するグループのタイトル
'            LaunchArgs     ユーザーがアクション センターで通知グループのタイトルをクリックしたときにアプリに渡される起動引数。
'            IconUri        アクション センターのタイトルの横に表示するアイコン。
'---------------------------------------------------------------------------------------------------
'* 注意事項：Desktopアプリだと、「CollectionId」が「displayName」になる問題があるため、Storeアプリでの使用を推奨します。
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-collections
'***************************************************************************************************
Public Property Get RunDll_ToastCollectionManagerSaveToastCollectionAsync(ByVal CollectionID As String, ByVal DisplayName As String, ByVal LaunchArgs As String, IconUri As String) As Long
    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = SetValueAppUserModelID(True)
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(False)

        .Tag = vbNullString
        .Group = vbNullString
        .Schedule_ID = vbNullString
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With
    
    
    'ToastCollectionを作成
    RunDll_ToastCollectionManagerSaveToastCollectionAsync = CreateToastCollection(ParamsConvertForDLL(ToastConfig), StrPtr(DisplayName), StrPtr(LaunchArgs), StrPtr(IconUri))
End Property

'***************************************************************************************************
'* 処理名　：RunDll_ToastCollectionManagerRemoveToastCollectionAsync
'* 機能　　：引数で渡された値で、コレクションを使用したトースト通知のグループ化を削除します。なお、省略時、全削除となります
'---------------------------------------------------------------------------------------------------
'* 返り値　：0で、成功扱いです。
'* 引数　　：CollectionId   コレクションの ID。 ID を使用して、ToastCollectionManager を使用してグループを操作できます
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-collections
'***************************************************************************************************
Public Property Get RunDll_ToastCollectionManagerRemoveToastCollectionAsync(Optional ByVal CollectionID As String) As Long
    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = SetValueAppUserModelID(True)
        'AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(False)

        .Tag = vbNullString
        .Group = vbNullString
        .Schedule_ID = vbNullString
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With
    
    
    'ToastCollectionを削除
    RunDll_ToastCollectionManagerRemoveToastCollectionAsync = DeleteToastCollection(ParamsConvertForDLL(ToastConfig))
End Property



'***************************************************************************************************
' ■■■ インスタンス生成時に自動で設定する内容 ■■■
'***************************************************************************************************
'* 機能　　：1.トーストのヘッダー情報を設定します。不要ならコメントアウトしてください。デフォルトで以下にしています
'               Header_ID           このExcelのファイル名
'               Header_Title        このExcelのBook名
'               Header_Arguments    このExcelの保存フォルダパス
'
'            2.タイムゾーンを日本基準にします。
'**************************************************************************************************
Private Sub Class_Initialize()
    With ThisWorkbook
        '拡張子の位置を調べる
        Dim FileTypePos As Long: FileTypePos = InStrRev(.Name, ".")

        '拡張子抜きのBook名を抽出(拡張子がない場合は、そのまま)
        Dim SetHeaderName As String: If FileTypePos Then SetHeaderName = Left(.Name, FileTypePos - 1) Else SetHeaderName = .Name

        'ヘッダー情報をセット
        SetToastHeader = SetHeaderName
    End With

    'タイムゾーンを日本基準にします
    SetTimeZone 9, 0
End Sub
