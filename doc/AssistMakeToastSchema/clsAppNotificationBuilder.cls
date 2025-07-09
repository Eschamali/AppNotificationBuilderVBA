VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppNotificationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   VBAから、Windows PowerShell を経由して、デスクトップ通知【トースト通知】を表示させます
'   また、DLLファイルを経由した通知手段も提供します。Ver1.5.0

' MIT License
'
' Copyright (c) 2025 eschamali
' https://github.com/Eschamali/AppNotificationBuilderVBA
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
'               ■■■ VBA用にカスタマイズした専用DLL 内部関数宣言セクション ■■■
'***************************************************************************************************
' 機能     ：C++で書かれたDLLに、 Windows.UI.Notifications 関連の処理を埋め込ませ、Shell経由より高速に処理できます
'---------------------------------------------------------------------------------------------------
Private Declare PtrSafe Sub ShowToastNotification Lib "AppNotificationBuilderVBA" (ByRef Params_String As ToastNotificationParams, ByRef Params_Binding As ToastNotificationVariable) '通知表示
Private Declare PtrSafe Function UpdateToastNotification Lib "AppNotificationBuilderVBA" (ByRef Params_String As ToastNotificationParams, ByRef Params_Binding As ToastNotificationVariable) As Long  '通知更新
Private Declare PtrSafe Sub RemoveToastNotification Lib "AppNotificationBuilderVBA" (ByRef Params_String As ToastNotificationParams)  '通知を削除
Private Declare PtrSafe Function CreateToastCollection Lib "AppNotificationBuilderVBA" (ByRef Params_String As ToastNotificationParams, ByVal DisplayName As LongPtr, ByVal LaunchArgs As LongPtr, ByVal IconUri As LongPtr) As Long 'ToastCollectionの作成
Private Declare PtrSafe Function DeleteToastCollection Lib "AppNotificationBuilderVBA" (ByRef Params_String As ToastNotificationParams) As Long 'ToastCollectionの削除
Private Declare PtrSafe Function CheckNotificationSetting Lib "AppNotificationBuilderVBA" (ByRef Params_String As ToastNotificationParams) As Long '通知check
Private Declare PtrSafe Function ExecuteSQLite Lib "AppNotificationBuilderVBA" (ByVal DB_Path As LongPtr, ByVal sql As LongPtr) As String    'SQLite操作用
Private Declare PtrSafe Function AttemptToWriteRegistry Lib "AppNotificationBuilderVBA" (ByVal hwnd As LongPtr, ByVal hinst As LongPtr, ByVal lpszCmdLine As LongPtr, ByVal nCmdShow As Long) As Long   'レジストリ操作用



'***************************************************************************************************
'                           ■■■ ユーザー定義エラー番号文言 ■■■
'***************************************************************************************************
Private Const vbObjectErrorID_0001 As String = "設定したスケジュールは、現在のシステム日時に対して過去の日時に設定されてます。" & vbCrLf & "設定したスケジュール通知日時が、未来であることを確認してください。 "
Private Const vbObjectErrorID_0002 As String = "設定したパラメーターは、定義されていません。"



'***************************************************************************************************
'                       ■■■ システムで使える特別なaction要素引数名 ■■■
'***************************************************************************************************
Const ActionArguments再通知 As String = "snooze"
Const ActionArguments解除   As String = "dismiss"



'***************************************************************************************************
'                           ■■■ データバインディング名 ■■■
'***************************************************************************************************
'最上位テキスト要素
Private Const DataBindingTopText_Title     As String = "{TopTextTitle}"
Private Const DataBindingTopText_Contents  As String = "{TopTextContents}"
Private Const DataBindingTopText_Attribute As String = "{TopTextAttribution}"

'Progress要素
Private Const DataBindingProgress_Title         As String = "{ProgressTitle}"
Private Const DataBindingProgress_Status        As String = "{ProgressStatus}"
Private Const DataBindingProgress_Value         As String = "{ProgressValue}"
Private Const DataBindingProgress_ValueString   As String = "{ProgressValueString}"



'***************************************************************************************************
'                               ■■■ 通知に関連した拡張機能 ■■■
'***************************************************************************************************
'  機能     ：Windows API 定義です。
'            ・mciSendString            任意のパスやURLの音源ファイルを通知音として再生させる機能に使います。
'            ・URLDownloadToFile        インターネット上の画像を通知に載せる機能に使います。
'---------------------------------------------------------------------------------------------------
Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long



'***************************************************************************************************
'                          ■■■ 一部の設定プロパティ値を列挙型で宣言 ■■■
'***************************************************************************************************
'  機能     ：トーストが表すシナリオを指定します。数字の代わりにこの名称でも機能します。
'---------------------------------------------------------------------------------------------------
Public Enum EnumToastScenario
    tsDefault
    tsReminder
    tsAlarm
    tsIncomingCall
    tsUrgent
End Enum

'===================================================================================================
'  機能     ：トースト通知を表示するときに再生するサウンドを指定します。数字の代わりにこの名称でも機能します。
'---------------------------------------------------------------------------------------------------
'Windows側で定められた通知音数
Private Const DefNoticeSoundNum As Integer = 25

Public Enum EnumMsWinSoundEvent
    NotificationMute
    NotificationDefault
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
'  機能     ：ボタンのスタイルを指定します。数字の代わりにこの名称でも機能します。
'---------------------------------------------------------------------------------------------------
Public Enum EnumToastButtonStyle
    tbsSuccess = 1
    tbsCritical
End Enum

'===================================================================================================
'  機能     ：ユーザーが特定の操作を行った際に使用されるアクティブ化の種類を決定します。
'---------------------------------------------------------------------------------------------------
Public Enum EnumToastActivationType
    taForeground = 1  'フォアグラウンド アプリが起動します。
    taBackground      '対応するバックグラウンド タスクがトリガーされ、ユーザーを中断することなくバックグラウンドでコードを実行できます。
    taProtocol        'プロトコルのアクティブ化を使用して別のアプリを起動します。
    taSystem          'Action要素でのリマインダーモード用です
End Enum
'---------------------------------------------------------------------------------------------------



'***************************************************************************************************
'                       ■■■ タイムゾーン　ユーザー定義型宣言 ■■■
'***************************************************************************************************
'* 機能　　：日付、時刻データを扱う場合は、UTC基準になる場合があるため、その設定です
'===================================================================================================
Private TargetTimeZone As TimeZone

Private Type TimeZone
    Hour    As Integer  '-23~23
    Minute  As Integer  '0-59
End Type



'***************************************************************************************************
'                   ■■■ トーストの挙動を決めるユーザー定義型宣言 ■■■
'***************************************************************************************************
' 機能：Tag,Group,suppresspopup等といった[Windows.UI.Notifications]関連の設定値をまとめたものです
'===================================================================================================
Private ToastConfig As ToastNotification

Private Type ToastNotification
    '[Windows.UI.Notifications]の挙動を決める設定値
    AppUserModelID              As String   'https://learn.microsoft.com/ja-jp/windows/win32/shell/appids
    AllowUse_InternetImageFlag  As Boolean  'インターネット上の画像を使用する権限設定
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
'       ■■■ VBA→DLLへの値参照が正しくできるようにするための中継用ユーザー定義型宣言 ■■■
'***************************************************************************************************
' 機能：スタック領域の制限の都合上、構造体としてを定義します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：・DLL側で、シグネチャ（型や順序）が合うようにすること。
'          ：・複数の型を混在させる場合、散在せずまとめるようにすること
'          ：・Type 名は、DLL側の struct と同一にすること
'          ：・Type 内の変数名も、DLL側の struct 内と同一にすること
'          ：・DLL 側の型が、 BOOL であっても、こちら VBA 側は、 long 型にすること
'===================================================================================================
Private Type ToastNotificationParams
    '文字列に関するパラメーター　LongPtrを使うことで、ポインタから文字列を取得(const wchar_t*)するようにします。
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

'***************************************************************************************************
' 機能：データバインディングの定義を用意します
'===================================================================================================
Private Type ToastNotificationVariable
    TitleText                   As LongPtr
    ContentsText                As LongPtr
    AttributionText             As LongPtr
    ProgressTitle               As LongPtr
    ProgressValueStringOverride As LongPtr
    ProgressStatus              As LongPtr
    ProgressValue               As Double
End Type

'===================================================================================================



'***************************************************************************************************
'                         ■■■ "wpndatabase.db"に関する定数/変数一式 ■■■
'***************************************************************************************************
'-----基本-----
Private LatestExecuteSQLiteNum      As Long
Private Const 設定DB保存相対パス    As String = "\Microsoft\Windows\Notifications\wpndatabase.db"
'-----WNSId 一覧(今のところ、システムでは2種類用意されてる)-----
Public Enum Wpndatabase_WNSId
    wsCustom                '好きな文字列指定用
    wsSystem                'システム
    wsNonImmersivePackage   '非パッケージ
End Enum
'-----HandlerType 一覧(今のところ、3種類用意されてる)-----
Public Enum Wpndatabase_HandlerType
    htDesktop       'Win32アプリ(非パッケージ)
    htImmersive     'パッケージアプリ(UWP)
    htSystem        'システムアプリ
End Enum
'-----SettingKey 一覧(今のところ、28種類用意されてる)-----
'ただし、一度でもアプリ側が対応するパラメーターを使用しないと出てこない仕様のため、まだまだあるかも…?
Public Enum Wpndatabase_SettingKey
    skS_tickle
    skC_tile
    skR_badge
    skS_lock_toast
    skS_tile
    skS_lock_badge
    skS_banner                              'レジストリとの連携必須
    skS_lock_tile
    skR_tile
    skS_badge
    skS_audio
    skS_voip
    skS_stopCloud
    skS_listenerEnabled
    skC_cloud
    skC_toast
    skC_internet
    skC_badge
    skC_ringing
    skC_tickle
    skC_voip
    skC_storage_toast
    skS_toast                               'レジストリとの連携必須
    skC_storage_tile
    skM_SecondaryTileMigrationComplete
    skS_cycle_medium
    skS_cycle_wide
    skS_cycle_large
End Enum



'***************************************************************************************************
'                         ■■■ レジストリに関する定数/変数一式 ■■■
'***************************************************************************************************
Private Const パラメーター区切り文字        As String = "|"
Private 最後に実行したレジストリ登録結果    As Long
Private Const HKCU配下の通知関連フルパス    As String = "Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\"
'登録型一覧
Enum RegistryType
    REG_Delete = -1 '削除用
    REG_DWORD
    REG_SZ
    REG_EXPAND_SZ
End Enum
'レジストリで設定できるパラメーターを用意
Enum RegistrySetting
    'HKEY_CLASSES_ROOT 用
    ShowInSettings
    DisplayName
    IconUri
    IconBackgroundColor
    
    'HKEY_CURRENT_USER 用
    '--wpndatabase.db 連携必須--
    Enabled
    ShowBanner
    '--wpndatabase.db 連携不要--
    SoundFile
    AllowContentAboveLock
    Rank
    ShowInActionCenter
    AllowUrgentNotifications
End Enum



'***************************************************************************************************
'                   ■■■ トースト スキーマに基づくメンバユーザー定義型宣言 ■■■
'***************************************************************************************************
'下記に基づいて順番に、宣言します
'→https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/root-elements

'===================================================================================================
'位置        ：Toast要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast
'---------------------------------------------------------------------------------------------------
Private ElementToast As ToastVisual

Private Type ToastVisual
    Duration         As String    '通知の表示時間延長フラグ
    Launch           As String    'トースト自体をクリックする際のスキーマ値
    DisplayTimestamp As String    'Windows プラットフォームによって通知が受信された時刻ではなく、通知コンテンツが実際に配信された日時を表すカスタム タイムスタンプで既定のタイムスタンプをオーバーライドします。 この値は、ISO 8601 標準を使用して書式設定された文字列である必要があります。
    ActivationType   As String    'ユーザーが特定の操作を行った際に使用されるアクティブ化の種類を決定します。公式にはないが、設定としてはあるもよう。VBA側では基本、"protocol"でしか価値がありません
    Synalio          As String    '通知のシナリオ属性
    UseButtonStyle   As Boolean   'スタイル付きボタン使用フラグ
End Type

'===================================================================================================
'位置        ：binding要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-binding
'---------------------------------------------------------------------------------------------------
Private ElementBinding As ToastBindingGeneric

Private Type ToastBindingGeneric
    Baseuri     As String   'イメージ ソース属性の相対 URI と組み合わせた既定のベース URI。
    template    As String   'この値は "ToastGeneric" に設定する必要があります
End Type

'===================================================================================================
'位置        ：binding親要素のimage子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-image
'---------------------------------------------------------------------------------------------------
Private ElementsImage As TopAdaptiveImage

'image要素で使える属性
Private Type AttributeImage
    AddImageQuery As Boolean    'イメージ URI にクエリ文字列を追加できるようにするかの設定値です
    alt           As String     '支援技術のユーザー向けの画像の説明
    src           As String     'イメージ ソースの URI。ローカル画像や、httpsソースも使える
    Placement     As String     'イメージ配置
    HintCrop      As String     '円にトリミングする印
End Type

'image要素の立ち位置(属性ノード：placement  の値ごとに分けています)
Private Type TopAdaptiveImage
    AppLogo As AttributeImage  'アイコン
    Inline  As AttributeImage  '画像はテキスト要素の後にインラインで表示され、ビジュアル領域の全幅が埋め込まれます。
    Hero    As AttributeImage  'トースト内で目立つように表示される注目のヒーロー画像を表示。常に通知上部に配置されます
End Type

'====================================================================================================
'位置        ：binding親要素のtext子要素で使える値と、属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-text
'----------------------------------------------------------------------------------------------------
Private ElementsText As TopAdaptiveText

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
Private Type TopAdaptiveText
    Title       As ParameterText    'タイトル
    Body        As ParameterText    'コンテンツ
    Attribution As ParameterText    '常に下部に表示するもの
End Type

'====================================================================================================
'位置        ：subgroup要素に付けれる属性値と値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-subgroup
'----------------------------------------------------------------------------------------------------
Private ElementsSubgroup()      As AdaptiveSubgroup, SubgroupTotal As Long
Private SubgroupInnerTexts()    As SubAdaptiveText, SubgroupInnerTextsTotal As Long
Private SubgroupInnerImages()   As SubAdaptiveImage, SubgroupInnerImagesTotal As Long
Private SubgroupInnerOrder()    As SubgroupInnerOrder, SubgroupInnerTotal As Long

'グループ/サブグループの子として配置されるテキスト要素
'https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-schema#adaptivetext
Private Type SubAdaptiveText
    '各属性
    HintStyle       As String   'テキストのフォント サイズ、太さ、および不透明度を制御します。
    HintWrap        As Boolean  'true に設定すると、テキストの折り返しが有効になります。
    HintMaxLines    As Integer  '表示が許可される、テキスト要素の最大行数です。
    HintMinLines    As Integer  '表示する必要のある、テキスト要素の最小行数です。
    HintAlign       As String   'テキストの水平方向の配置。
    Language        As String   'XML ペイロードの対象ロケールです。
    
    '値
    Value          As String   '表示するテキスト。
End Type

'グループ/サブグループの子として配置されるインライン画像
'https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-schema#adaptiveimage
Private Type SubAdaptiveImage
    '各属性
    src                 As String   '画像パス
    HintCrop            As String   'イメージの目的のトリミングを制御
    HintRemoveMargin    As Boolean  'このプロパティを true に設定することで余白を削除できます。
    HintAlign           As String   '画像の水平方向の配置です。
    AlternateText       As String   'アクセシビリティ対応目的で使用される、画像を説明する代替テキストです。
    AddImageQuery       As Boolean  'クエリ文字列を追加できるようになります
End Type

'配置順の管理
Private Type SubgroupInnerOrder
    UseText           As Boolean  'テキストか画像かのフラグ値
    TargetElementNo   As Long     '対象の要素No
End Type

'subgroupに付けれる属性と値、順序データ
Private Type AdaptiveSubgroup
    '各属性
    HintWeight As Integer       '別のサブグループを基準として太さを指定することで、このサブグループの列の幅を制御します。
    HintTextStacking As String  'このサブグループのコンテンツの垂直方向の配置を制御します。

    '順序データ
    QueueOrder() As SubgroupInnerOrder

    '値
    QueueText() As SubAdaptiveText
    QueueImage() As SubAdaptiveImage
End Type

'<subgroup用の便利な列挙型>
'画像トリミングを指定します。
Public Enum EnumAdaptiveImageCrop
    ImageCrop_Default   '既定値です。トリミングの動作がレンダラーによって決定されます。
    ImageCrop_None      '画像がトリミングされません。
    ImageCrop_Circle    '画像が円形にトリミングされます。
End Enum

'画像の水平方向の配置を指定します。
Public Enum EnumAdaptiveImageAlign
    ImageAlign_Default  '既定値です。配置の動作がレンダラーによって決定されます。
    ImageAlign_Stretch  '利用可能な幅いっぱいに画像が拡大されます (また同時に、画像が配置される位置に応じて、利用可能な高さいっぱいに拡大されることもあります)。
    ImageAlign_Left     '画像を左側に配置し、ネイディブの解像度で表示します。
    ImageAlign_Center   '画像を中央に水平方向に配置し、ネイティブの解像度で表示します。
    ImageAlign_Right    '画像を右側に配置し、ネイディブの解像度で表示します。
End Enum

'フォント サイズ、太さ、および不透明度を制御します。 "Subtle" の不透明度は 60% の不透明度になります。
Public Enum EnumAdaptiveTextStyle
    TextStyle_Default           '既定値です。 スタイルがレンダラーによって決定されます。
    TextStyle_Caption           '段落のフォント サイズより小さいサイズです。
    TextStyle_CaptionSubtle     'Caption と同じですが、不透明度が Subtle です。
    TextStyle_Body              '段落本文のフォント サイズです。
    TextStyle_BodySubtle        'Body と同じですが、不透明度が Subtle です。
    TextStyle_Base              '段落本文のフォント サイズで、太字です。 基本的には、Body の太字バージョンと言えます。
    TextStyle_BaseSubtle        'Base と同じですが、不透明度が Subtle です。
    TextStyle_Subtitle          'H4 のフォント サイズです。
    TextStyle_SubtitleSubtle    'Subtitle と同じですが、不透明度が Subtle です。
    TextStyle_Title             'H3 のフォント サイズです。
    TextStyle_TitleSubtle       'Title と同じですが、不透明度が Subtle です。
    TextStyle_TitleNumeral      'Title と同じですが、上/下のパディングが削除されます。
    TextStyle_Subheader         'H2 のフォント サイズです。
    TextStyle_SubheaderSubtle   'Subheader と同じですが、不透明度が Subtle です。
    TextStyle_SubheaderNumeral  'Subheader と同じですが、上/下のパディングが削除されます。
    TextStyle_Header            'H1 のフォント サイズです。
    TextStyle_HeaderSubtle      'Header と同じですが、不透明度が Subtle です。
    TextStyle_HeaderNumeral     'Header と同じですが、上/下のパディングが削除されます。
End Enum

'テキストの水平方向の配置を制御します。
Public Enum EnumAdaptiveTextAlign
    TextAlign_Default '既定値です。 配置がレンダラーによって自動的に決定されます。
    TextAlign_Auto    '配置が現在の言語とカルチャによって決定されます。
    TextAlign_Left    'テキストを左側に水平方向に配置します。
    TextAlign_Center  'テキストを中央に水平方向に配置します。
    TextAlign_Right   'テキストを右側に水平方向に配置します。
End Enum

'コンテンツの垂直方向の配置を指定します。
Public Enum AdaptiveSubgroupTextStacking
    TextStacking_Default    '既定値です。 レンダラーが既定の垂直方向の配置を自動的に選択します。
    TextStacking_Top        '上に合わせて垂直に配置されます。
    TextStacking_Center     '中央に合わせて垂直に配置されます。
    TextStacking_Bottom     '下に合わせて垂直に配置されます。
End Enum



'====================================================================================================
'位置        ：binding親要素のprogress子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-progress
'----------------------------------------------------------------------------------------------------
Private ElementProgress As AdaptiveProgressBar

Private Type AdaptiveProgressBar
    Title               As String   '省略可能なタイトル文字列。
    Status              As String   '左側の進行状況バーの下に表示されるステータス文字列。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります。
    Value               As Double   'プログレスバーの値。 この値には、0.0 ~ 1.0 の浮動小数点数、または読み込みアニメーションを生成する値 "Indeterminate" を指定します。なおVBAの場合、負の場合で"Indeterminate"扱いとします。
    ValueStringOverride As String   '既定のパーセンテージ文字列の代わりに表示される省略可能な文字列。 これを指定しない場合は、"70%" などの文字が表示されます。
End Type

'====================================================================================================
'位置        ：Audio要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio
'----------------------------------------------------------------------------------------------------
Private ElementAudio As ToastAudio

Private Type ToastAudio
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

Private ElementAction(1 To MaxElementActionNum) As IToastActions

Private Type IToastActions
    Content                 As String   'ボタンに表示されるコンテンツ。
    Arguments               As String   'ユーザーがこのボタンをクリックした場合にアプリが後から受け取る、アプリで定義された引数の文字列です。
    Type                    As String   '関連付けられたアプリに渡すことができる引数文字列。ユーザー アクションに応答して実行する必要があるアクションに関する詳細を提供します。
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

Private ElementSelection(1 To MaxElementSelectionNum) As ToastSelectionBox

Private Type ToastSelectionBox
    SelectionID As String   '識別子名
    ChoseName   As String   '選択肢に出す名称
End Type

'====================================================================================================
'位置        ：actions親要素のinput子要素で使える属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-input
'----------------------------------------------------------------------------------------------------
'入れれる最大要素数を定数定義
Private Const MaxElementInputNum As Integer = 5    '入力要素は最大5つ、入れれます

Private ElementInput(1 To MaxElementInputNum) As IToastInput

Private Type IToastInput
    '属性一覧
    id                 As String    '入力に関連付けられている ID
    Type               As String    'この属性には、"text"、"selection" のいずれかの値を指定できます。
    PlaceHolderContent As String    'テキスト入力用に表示されるプレースホルダー
    Title              As String    '入力のラベルとして表示されるテキスト。
    DefaultInput       As String    'デフォルトの入力値(公式にはないが、設定項目としてはあるもよう)

    '要素そのものの配列(Select要素が入ります)
    Value() As ToastSelectionBox
End Type

'====================================================================================================
'位置        ：header要素に付けれる属性値
'ドキュメント：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-header
'----------------------------------------------------------------------------------------------------
Private ElementHeader As ToastHeader

Private Type ToastHeader
    id             As String    'このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。
    Title          As String    'ヘッダーのタイトルです。
    Arguments      As String    'ユーザーがこのヘッダーをクリックするとアプリに返されます。 null にすることはできません。
    ActivationType As String    'このヘッダーがクリックされた場合に使用するアクティブ化の種類。
End Type



'***************************************************************************************************
' ■■■ VBA上で使うのに相応しいAppUserModelIDをプリセット ■■■
'***************************************************************************************************
Private Const AppUserModelID_Excel  As String = "Microsoft.Office.EXCEL.EXE.15" 'DeskTopアプリ：Excel
Private AppUserModelID_Any          As String                                   '任意のAppID設定用変数



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
'* 機能説明：設定すると、ExcelのAppUserModelIDよりも優先して設定されます。batファイル等、VBAと直接関係なくなり、アイコンが噛み合わない時に使います。
'* 注意事項：存在しない(未インストール)AppID、無効な文字列を指定すると、Toastが発行されないのでご注意ください。
'---------------------------------------------------------------------------------------------------
Public Property Let SetAppUserModelID(ByVal ArgAppID As String)
    AppUserModelID_Any = ArgAppID
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

'===================================================================================================
'* 機能    ：DLLで扱えるユーザー定義型：ToastNotificationVariable　に変換します
'---------------------------------------------------------------------------------------------------
'* 注意事項：Windows API の文字列関係は、ポインターで取得するパターンが多いので、StrPtrで経由する必要があります
'---------------------------------------------------------------------------------------------------
Private Function ReturnDataBinding() As ToastNotificationVariable
    With ReturnDataBinding
        '最上位レベルのテキスト要素
        .TitleText = StrPtr(ElementsText.Title.Value)
        .ContentsText = StrPtr(ElementsText.Body.Value)
        .AttributionText = StrPtr(ElementsText.Attribution.Value)
        
        '進行状況バー要素
        '  Statusが設定済みの場合のみ定義する(公式ドキュメント準拠)
        If StrPtr(ElementProgress.Status) Then
            .ProgressTitle = StrPtr(ElementProgress.Title)
            .ProgressStatus = StrPtr(ElementProgress.Status)
            .ProgressValue = ElementProgress.Value
            .ProgressValueStringOverride = StrPtr(ElementProgress.ValueStringOverride)
        End If
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
    If ArgExtensionDuration Then ElementToast.Duration = "long" Else ElementToast.Duration = vbNullString
End Property

'***************************************************************************************************
'* 機能　　：toast要素のlaunch、activationType属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgSchema          起動スキーマ。URLやファイルパスなど
'            ActivationType     このヘッダーがクリックされた場合に使用するアクティブ化の種類(DLLを使用しない場合、"protocol"以外は機能しません)
'***************************************************************************************************
Public Property Let SetToastContent_Launch(Optional ByVal ActivationType As EnumToastActivationType = EnumToastActivationType.taProtocol, ByVal ArgSchema As String)
    With ElementToast
        .Launch = ArgSchema
        .ActivationType = SetToastActivationType(ActivationType)
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
        If ArgDisplayTimestamp Then ElementToast.DisplayTimestamp = .Text(ArgDisplayTimestamp, "yyyy-mm-ddThh:mm:ss") & .Text(TargetTimeZone.Hour, "+00:;-00:") & .Text(TargetTimeZone.Minute, "00") Else ElementToast.DisplayTimestamp = vbNullString
    End With
End Property

'***************************************************************************************************
'* 機能　　：toast要素のsynalio属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID    シナリオID。便宜上、数値に応じてシナリオ名を設定します
'***************************************************************************************************
Public Property Let SetToastContent_Scenario(ToastScenarioMode As EnumToastScenario)
    'ArgIDから、シナリオ名を決めます。
    With ElementToast
        Select Case ToastScenarioMode
            Case EnumToastScenario.tsReminder
                '通知を永遠に表示する。action要素がないと効果は発動しないが、任意の通知音に設定可能
                .Synalio = "reminder"

            Case EnumToastScenario.tsAlarm
                '通知を永遠に表示する。action要素がないと効果は発動しないが、応答不可になっても設定問わず、必ず表示される。通知音はアラーム系(Alarm)か、ミュートのみ設定可能
                .Synalio = "alarm"

            Case EnumToastScenario.tsIncomingCall
                '通知を永遠に表示する。action要素がなくても効果発動。最後のボタン位置のみ、Windowsのテーマ色に基づく着色が施され、位置が必ず下側になる。通知音は呼び出し系(Call)か、ミュートのみ設定可能
                .Synalio = "incomingCall"

            Case EnumToastScenario.tsUrgent
                '重要な通知として表示。応答不可にしても表示されますが、設定で表示させないようにすることが可能です。とりあえず常に表示させたい場合は、シナリオ：「アラーム」にしましょう。
                .Synalio = "urgent"

            Case Else
                '通常の通知として表示。属性：duration  あるいは、アクセシビリティ→視覚効果→この時間が経過したら通知を破棄する  に基づく表示時間となります
                .Synalio = vbNullString
        End Select
    End With
End Property

'***************************************************************************************************
'* 機能    ：引数に応じて、アクティブ化の種類を文字列で返します
'***************************************************************************************************
Private Function SetToastActivationType(ByVal activationID As EnumToastActivationType) As String
    Select Case activationID
        Case EnumToastActivationType.taForeground
            SetToastActivationType = "foreground"
        Case EnumToastActivationType.taBackground
            SetToastActivationType = "background"
        Case EnumToastActivationType.taProtocol
            SetToastActivationType = "protocol"
        Case EnumToastActivationType.taSystem
            SetToastActivationType = "system"
        Case Else
            SetToastActivationType = vbNullString
    End Select
End Function

'***************************************************************************************************
'* 機能　　：toast要素のuseButtonStyle属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArguseButtonStyle
'               True ：スタイル付きボタンを使用します
'               False：スタイル付きボタンを使用しません。
'---------------------------------------------------------------------------------------------------
'* 注意事項：この機能は、Windows 11 以降のみ機能します。
'***************************************************************************************************
Public Property Let SetToastContent_UseButtonStyle(ByVal ArgUseButtonStyle As Boolean)
    ElementToast.UseButtonStyle = ArgUseButtonStyle
End Property

'***************************************************************************************************
'* 機能　　binding要素の各属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgTemp            イメージ ソース属性の相対 URI と組み合わせた既定のベース URI。
'            ArgBaseImageURI    この値は "ToastGeneric" に設定する必要があります
'***************************************************************************************************
Public Property Let SetToastBindingGeneric(Optional ByVal ArgTemp As String = "ToastGeneric", ByVal ArgBaseImageURI As String)
    With ElementBinding
        .Baseuri = ArgBaseImageURI
        .template = ArgTemp     '※この値は "ToastGeneric" に設定する必要があります
    End With
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
Public Property Let SetToastGenericInlineImage(Optional ByVal Arg_InlineCircle As Integer, Optional ByVal Flag_addImageQuery As Boolean, Optional ByVal Arg_Alt As String, ByVal Arg_InlineImageSrc As String)
    With ElementsImage.Inline
        .src = UseHttpImageCheck(Arg_InlineImageSrc)
        .AddImageQuery = Flag_addImageQuery
        .alt = Arg_Alt

        If Arg_InlineCircle Then .HintCrop = "circle" Else .HintCrop = vbNullString
    End With
End Property

'***************************************************************************************************
'* 機能　　：binding親要素のimage子要素のうち、ヒーロー画像のパスの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：Arg_HeroImageSrc       イメージ ソースの URI。ファイルパスや、http(s)画像など
'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Arg_Alt                画像の説明を記載します。
'***************************************************************************************************
Public Property Let SetToastGenericHeroImage(Optional ByVal Flag_addImageQuery As Boolean, Optional ByVal Arg_Alt As String, ByVal Arg_HeroImageSrc As String)
    With ElementsImage.Hero
        .src = UseHttpImageCheck(Arg_HeroImageSrc)
        .AddImageQuery = Flag_addImageQuery
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
Public Property Let SetToastGenericAppLogo(Optional ByVal Arg_LogoCircle As Boolean, Optional ByVal Flag_addImageQuery As Boolean, Optional ByVal Arg_Alt As String, ByVal Arg_AppLogoSrc As String)
    With ElementsImage.AppLogo
        .src = UseHttpImageCheck(Arg_AppLogoSrc)
        .AddImageQuery = Flag_addImageQuery
        .alt = Arg_Alt

        .Placement = "appLogoOverride"

        If Arg_LogoCircle Then .HintCrop = "circle" Else .HintCrop = vbNullString
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
Public Property Let SetToastGenericTitleText(Optional ByVal HintCallScenarioCenterAlign As Boolean, ByVal TextContents As String)
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
Public Property Let SetToastGenericContentsText(Optional ByVal HintCallScenarioCenterAlign As Boolean, ByVal TextContents As String)
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
Public Property Let SetToastGenericAttributionText(Optional ByVal HintCallScenarioCenterAlign As Boolean, ByVal TextContents As String)
    With ElementsText.Attribution
        .Value = TextContents
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign

        .Attribute.Placement = "attribution"
    End With
End Property

'***************************************************************************************************
'* 機能　　：subgroup要素に属性と、事前に定義したtext,image要素をそれぞれ追加します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Weight         別のサブグループを基準として太さを指定することで、このサブグループの列の幅を制御します。
'            TextStacking   このサブグループのコンテンツの垂直方向の配置を制御します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：このプロパティを使用すると以前、「AddAdaptiveSubgroupImage」、「AddAdaptiveSubgroupText」で呼び出して設定した配列、
'             SubgroupInnerImages,SubgroupInnerTexts の中身をリセットします。
'***************************************************************************************************
Public Sub AddAdaptiveSubgroup(Optional ByVal Weight As Integer, Optional ByVal TextStacking As AdaptiveSubgroupTextStacking)
    '加算処理(0基準)
    SubgroupTotal = SubgroupTotal + 1

    '配列を増やします
    ReDim Preserve ElementsSubgroup(SubgroupTotal)
    
    'パラメーターを設定します
    With ElementsSubgroup(SubgroupTotal)
        '各属性
        .HintWeight = Weight
        
        Select Case TextStacking

            Case AdaptiveSubgroupTextStacking.TextStacking_Top
                .HintTextStacking = "top"
            
            Case AdaptiveSubgroupTextStacking.TextStacking_Center
                .HintTextStacking = "center"
            
            Case AdaptiveSubgroupTextStacking.TextStacking_Bottom
                .HintTextStacking = "bottom"
            
            Case Else
                .HintTextStacking = vbNullString
        End Select
        
        
        '順序データ
        .QueueOrder = SubgroupInnerOrder
        
        '各値
        .QueueText = SubgroupInnerTexts
        .QueueImage = SubgroupInnerImages
    End With

    '以前定義したsubgroup要素内のtext,image要素をリセットする
    '  カウントリセット
    SubgroupInnerTextsTotal = -1
    SubgroupInnerImagesTotal = -1
    SubgroupInnerTotal = -1
    '  配列リセット
    ReDim SubgroupInnerTexts(SubgroupInnerTextsTotal + 1)
    ReDim SubgroupInnerImages(SubgroupInnerImagesTotal + 1)
    ReDim SubgroupInnerOrder(SubgroupInnerTotal + 1)
End Sub

'***************************************************************************************************
'* 機能　　：subgroup親要素のimage子要素に値と属性をそれぞれ追加します
'---------------------------------------------------------------------------------------------------
'* 引数　　：       src                 As String   '画像パス
'                   HintCrop            As String   'イメージの目的のトリミングを制御
'                   HintRemoveMargin    As Boolean  'このプロパティを true に設定することで余白を削除できます。
'                   HintAlign           As String   '画像の水平方向の配置です。
'                   AlternateText       As String   'アクセシビリティ対応目的で使用される、画像を説明する代替テキストです。
'                   AddImageQuery       As Boolean  'クエリ文字列を追加できるようになります
'---------------------------------------------------------------------------------------------------
'* 機能説明：アダプティブを使用した画像をレンダリング出来ます。
'***************************************************************************************************
Public Property Let AddAdaptiveSubgroupImage(Optional ByVal Crop As EnumAdaptiveImageCrop, Optional ByVal RemoveMargin As Boolean, Optional ByVal Align As EnumAdaptiveImageAlign, Optional ByVal AltText As String, Optional ByVal AddImageQuery As Boolean, ImageSrc As String)
    '加算処理(0基準)
    SubgroupInnerImagesTotal = SubgroupInnerImagesTotal + 1     '画像挿入数
    SubgroupInnerTotal = SubgroupInnerTotal + 1                 'Subgroup要素内の挿入要素数
    
    '配列を増やします
    ReDim Preserve SubgroupInnerImages(SubgroupInnerImagesTotal)
    
    '順序データを設定します
    ReDim Preserve SubgroupInnerOrder(SubgroupInnerTotal)
    SubgroupInnerOrder(SubgroupInnerTotal).TargetElementNo = SubgroupInnerImagesTotal
    SubgroupInnerOrder(SubgroupInnerTotal).UseText = False

    'パラメーターを設定します
    With SubgroupInnerImages(SubgroupInnerImagesTotal)
        '各属性
        .src = UseHttpImageCheck(ImageSrc)
        
        Select Case Crop
            Case EnumAdaptiveImageCrop.ImageCrop_Circle
                .HintCrop = "circle"
            
            Case EnumAdaptiveImageCrop.ImageCrop_None
                .HintCrop = "none"
            
            Case Else
                .HintCrop = vbNullString
        End Select
        
        .HintRemoveMargin = RemoveMargin
        
        Select Case Align
            Case EnumAdaptiveImageAlign.ImageAlign_Stretch
                .HintAlign = "stretch"
            
            Case EnumAdaptiveImageAlign.ImageAlign_Left
                .HintAlign = "left"
            
            Case EnumAdaptiveImageAlign.ImageAlign_Center
                .HintAlign = "center"
            
            Case EnumAdaptiveImageAlign.ImageAlign_Right
                .HintAlign = "right"
            
            Case Else
                .HintAlign = vbNullString
        End Select
        
        .AlternateText = AltText
        .AddImageQuery = AddImageQuery
    End With
End Property

'***************************************************************************************************
'* 機能　　：subgroup親要素のtext子要素に値と属性をそれぞれ追加します
'---------------------------------------------------------------------------------------------------
'* 引数　　：   Style        テキストのフォント サイズ、太さ、および不透明度を制御します。
'               Wrap         true に設定すると、テキストの折り返しが有効になります。
'               MaxLines     表示が許可される、テキスト要素の最大行数です。
'               MinLines     表示する必要のある、テキスト要素の最小行数です。
'               Align        テキストの水平方向の配置。
'               Arg_Language XML ペイロードの対象ロケールです。
'               TextContents 表示テキスト
'---------------------------------------------------------------------------------------------------
'* 機能説明：アダプティブを使用したテキストをレンダリング出来ます。
'***************************************************************************************************
Public Property Let AddAdaptiveSubgroupText(Optional ByVal Style As EnumAdaptiveTextStyle, Optional ByVal Wrap As Boolean, Optional ByVal MaxLines As Integer, Optional ByVal MinLines As Integer, Optional ByVal Align As EnumAdaptiveTextAlign, Optional ByVal Arg_Language As String, ByVal TextContents As String)
    '加算処理(0基準)
    SubgroupInnerTextsTotal = SubgroupInnerTextsTotal + 1   'テキスト挿入数
    SubgroupInnerTotal = SubgroupInnerTotal + 1             'Subgroup要素内の挿入要素数
    
    '配列を増やします
    ReDim Preserve SubgroupInnerTexts(SubgroupInnerTextsTotal)
    
    '順序データを設定します
    ReDim Preserve SubgroupInnerOrder(SubgroupInnerTotal)
    SubgroupInnerOrder(SubgroupInnerTotal).TargetElementNo = SubgroupInnerTextsTotal
    SubgroupInnerOrder(SubgroupInnerTotal).UseText = True

    'パラメーターを設定します
    With SubgroupInnerTexts(SubgroupInnerTextsTotal)
        '各属性
        Select Case Style
            Case EnumAdaptiveTextStyle.TextStyle_Caption
                .HintStyle = "caption"

            Case EnumAdaptiveTextStyle.TextStyle_CaptionSubtle
                .HintStyle = "captionsubtle"

            Case EnumAdaptiveTextStyle.TextStyle_Body
                .HintStyle = "body"

            Case EnumAdaptiveTextStyle.TextStyle_BodySubtle
                .HintStyle = "bodysubtle"

            Case EnumAdaptiveTextStyle.TextStyle_Base
                .HintStyle = "base"

            Case EnumAdaptiveTextStyle.TextStyle_BaseSubtle
                .HintStyle = "basesubtle"

            Case EnumAdaptiveTextStyle.TextStyle_Subtitle
                .HintStyle = "subtitle"

            Case EnumAdaptiveTextStyle.TextStyle_SubtitleSubtle
                .HintStyle = "subtitlesubtle"

            Case EnumAdaptiveTextStyle.TextStyle_Title
                .HintStyle = "title"

            Case EnumAdaptiveTextStyle.TextStyle_TitleSubtle
                .HintStyle = "titlesubtle"

            Case EnumAdaptiveTextStyle.TextStyle_TitleNumeral
                .HintStyle = "titlenumeral"

            Case EnumAdaptiveTextStyle.TextStyle_Subheader
                .HintStyle = "subheader"

            Case EnumAdaptiveTextStyle.TextStyle_SubheaderSubtle
                .HintStyle = "subheadersubtle"

            Case EnumAdaptiveTextStyle.TextStyle_SubheaderNumeral
                .HintStyle = "subheadernumeral"

            Case EnumAdaptiveTextStyle.TextStyle_Header
                .HintStyle = "header"

            Case EnumAdaptiveTextStyle.TextStyle_HeaderSubtle
                .HintStyle = "headersubtle"

            Case EnumAdaptiveTextStyle.TextStyle_HeaderNumeral
                .HintStyle = "headernumeral"
            
            Case Else
                .HintStyle = vbNullString
        
        End Select

        .HintWrap = Wrap
        .HintMaxLines = MaxLines
        .HintMinLines = MinLines
        
        Select Case Align
            Case EnumAdaptiveTextAlign.TextAlign_Auto
                .HintAlign = "auto"
            
            Case EnumAdaptiveTextAlign.TextAlign_Left
                .HintAlign = "left"
            
            Case EnumAdaptiveTextAlign.TextAlign_Center
                .HintAlign = "center"
            
            Case EnumAdaptiveTextAlign.TextAlign_Right
                .HintAlign = "right"
            
            Case Else
                .HintAlign = vbNullString
        End Select
        
        .Language = Arg_Language
        
        '値
        .Value = TextContents
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
Public Property Let SetAdaptiveProgressBar(Optional Arg_Value As Double, Optional Arg_Title As String, Optional Arg_ValueStringOverride As String, Arg_Status As String)
    With ElementProgress
        .Title = Arg_Title & ""
        .Status = Arg_Status
        .Value = Arg_Value
        .ValueStringOverride = Arg_ValueStringOverride
    End With
End Property

'***************************************************************************************************
'* 機能　　：Audio要素の属性をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgSrcID       通知音ID。便宜上、数値に応じて通知音名を設定します
'               1-25        Windows側で定められた通知音名を番号で指定
'               上記以外    ミュートにします

'            ArgLoop
'               True  ：トーストが表示されている限りサウンドが繰り返される場合に設定
'               False ：1回だけ再生する場合に設定
'***************************************************************************************************
Public Property Let SetToastAudio(Optional ByVal ArgLoop As Boolean, ArgSrcID As EnumMsWinSoundEvent)
    With ElementAudio
        'audio要素のLoop属性を設定
        .loop = ArgLoop

        'ArgSrcID に応じて、ミュート判定を決める
        If 0 < ArgSrcID And ArgSrcID <= DefNoticeSoundNum Then
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
            Case EnumMsWinSoundEvent.NotificationDefault
                .src = "ms-winsoundevent:Notification.Default"

            Case EnumMsWinSoundEvent.NotificationIM
                .src = "ms-winsoundevent:Notification.IM"
                
            Case EnumMsWinSoundEvent.NotificationMail
                .src = "ms-winsoundevent:Notification.Mail"

            Case EnumMsWinSoundEvent.NotificationReminder
                .src = "ms-winsoundevent:Notification.Reminder"

            Case EnumMsWinSoundEvent.NotificationSMS
                .src = "ms-winsoundevent:Notification.SMS"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm01
                .src = "ms-winsoundevent:Notification.Looping.Alarm"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm02
                .src = "ms-winsoundevent:Notification.Looping.Alarm2"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm03
                .src = "ms-winsoundevent:Notification.Looping.Alarm3"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm04
                .src = "ms-winsoundevent:Notification.Looping.Alarm4"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm05
                .src = "ms-winsoundevent:Notification.Looping.Alarm5"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm06
                .src = "ms-winsoundevent:Notification.Looping.Alarm6"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm07
                .src = "ms-winsoundevent:Notification.Looping.Alarm7"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm08
                .src = "ms-winsoundevent:Notification.Looping.Alarm8"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm09
                .src = "ms-winsoundevent:Notification.Looping.Alarm9"

            Case EnumMsWinSoundEvent.NotificationLoopingAlarm10
                .src = "ms-winsoundevent:Notification.Looping.Alarm10"

            Case EnumMsWinSoundEvent.NotificationLoopingCall01
                .src = "ms-winsoundevent:Notification.Looping.Call"

            Case EnumMsWinSoundEvent.NotificationLoopingCall02
                .src = "ms-winsoundevent:Notification.Looping.Call2"

            Case EnumMsWinSoundEvent.NotificationLoopingCall03
                .src = "ms-winsoundevent:Notification.Looping.Call3"

            Case EnumMsWinSoundEvent.NotificationLoopingCall04
                .src = "ms-winsoundevent:Notification.Looping.Call4"

            Case EnumMsWinSoundEvent.NotificationLoopingCall05
                .src = "ms-winsoundevent:Notification.Looping.Call5"

            Case EnumMsWinSoundEvent.NotificationLoopingCall06
                .src = "ms-winsoundevent:Notification.Looping.Call6"

            Case EnumMsWinSoundEvent.NotificationLoopingCall07
                .src = "ms-winsoundevent:Notification.Looping.Call7"

            Case EnumMsWinSoundEvent.NotificationLoopingCall08
                .src = "ms-winsoundevent:Notification.Looping.Call8"

            Case EnumMsWinSoundEvent.NotificationLoopingCall09
                .src = "ms-winsoundevent:Notification.Looping.Call9"

            Case EnumMsWinSoundEvent.NotificationLoopingCall10
                .src = "ms-winsoundevent:Notification.Looping.Call10"

            Case Else
                .src = vbNullString

        End Select

    End With
End Property

'***************************************************************************************************
'* 機能　　：action要素の各属性にセットしたものをactions要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Content            ボタンに表示されるコンテンツ
'            Arguments          起動スキーマ。URLやファイルパスなど
'            ArgType            関連付けられたアプリに渡すことができる引数文字列。ユーザー アクションに応答して実行する必要があるアクションに関する詳細を提供します。
'            ActivationType     このボタンが押下された場合に使用するアクティブ化の種類。DLLを使用しない場合、基本"protocol"以外では意味をなしませんが、ここのボタンの場合、"system"にして、所々の設定を行うことで、システムの再通知と無視アクションを利用するアラーム通知を作成できます。
'            PendingUpdate      ユーザーがトースト上のボタンをクリックすると、通知は "保留中の更新" 表示状態のままです。 この "更新の保留中" の表示状態が長時間続くことを避けるため、バックグラウンド タスクから即座にトーストを更新する必要があります。
'            ContextMenu        Trueで、従来のトースト ボタンではなく、トースト通知のコンテキスト メニューに追加されたコンテキスト メニュー アクションになります。
'            IconPath           トースト ボタン アイコンのイメージソース
'            HintInputId        テキストBoxの横に来させたいIDを指定します
'            HintButtonStyle    ボタンのスタイル。現状、「緑、赤」の2色のみ対応です。便宜上、数値に応じてスタイル名を設定します。※この機能は、Windows 11 以降のみ機能します。
'            HintToolTip        ボタンに空のコンテンツ文字列がある場合のボタンのヒント。
'            Index              追加位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：公式Document上には、type属性があるようですがここでは一旦無視します。
'***************************************************************************************************
Public Property Let SetIToastActions(ByVal Content As String, ByVal Arguments As String, Optional ByVal ActivationType As EnumToastActivationType = EnumToastActivationType.taProtocol, Optional PendingUpdate As Boolean, Optional ByVal ContextMenu As Boolean, Optional ByVal IconPath As String, Optional ByVal HintInputId As String, Optional HintButtonStyle As EnumToastButtonStyle, Optional ByVal HintToolTip As String, ByVal Index As Integer)
    With ElementAction(Index)
        .Content = Content
        .Arguments = Arguments
        .ActivationType = SetToastActivationType(ActivationType)
        '.Type = ArgType

        If PendingUpdate Then .afterActivationBehavior = "pendingUpdate" Else .afterActivationBehavior = vbNullString
        If ContextMenu Then .Placement = "contextMenu" Else .Placement = vbNullString

        .ImageUri = UseHttpImageCheck(IconPath)
        .HintInputId = HintInputId

        '属性：hint-buttonStyle  は、引数の値に応じて命名します
        Select Case HintButtonStyle
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

        .HintToolTip = HintToolTip
    End With
End Property

'***************************************************************************************************
'* 機能　　：input要素の各属性にセットしたものをaction要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID                  入力に関連付けられている ID
'            ChoseFlag              Trueで"SelectionBox"、Falseで、"TextBox"扱いとします
'            ArgPlaceHolderContent  テキスト入力用に表示されるプレースホルダー。
'            Argtitle               入力のラベルとして表示されるテキスト
'            ArgDefaultInput        デフォルトの入力値
'            Index                  追加位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：「ChoseFlag = True」でこのプロパティを使用すると以前、「SetToastSelectionBox」で呼び出して設定した配列、
'            ElementSelection の中身をリセットします。
'***************************************************************************************************
Public Property Let SetIToastInput(ByVal ArgID As String, Optional ByVal ChoseFlag As Boolean, Optional ByVal ArgPlaceHolderContent As String, Optional ByVal ArgTitle As String, Optional ByVal ArgDefaultInput As String, ByVal Index As Integer)
    With ElementInput(Index)
        .id = ArgID

        'Typeが、"selection"の場合、予め作成されたSelection要素を入れるようにします。
        If ChoseFlag Then
            .Type = "selection"
            .Value = ElementSelection
            
            '挿入後、同じ変数で使い回すため、初期化する
            Dim i As Long
            For i = 1 To MaxElementSelectionNum
                ElementSelection(i).SelectionID = vbNullString
                ElementSelection(i).ChoseName = vbNullString
            Next
        Else
            .Type = "text"
        End If

        .PlaceHolderContent = ArgPlaceHolderContent
        .Title = ArgTitle
        .DefaultInput = ArgDefaultInput
    End With
End Property

'***************************************************************************************************
'* 機能　　：input要素のtype属性が"selection" の時に挿入する"selection"要素を定義します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgSelectionID    識別子。「ToastNotification.Activated」を利用する場合、これが"UserInput"の値として返されます。
'                              なお、数値のみの指定にすると、リマインダー用途としても使用可能です(単位：分)
'
'            ArgChoseName      選択肢名称
'            Index             挿入位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：DLLを使用しない場合、リマインダー用途しか使い所がありません
'***************************************************************************************************
Public Property Let SetToastSelectionBox(ByVal ArgSelectionID As String, ByVal ArgChoseName As String, ByVal Index As Integer)
    With ElementSelection(Index)
        .SelectionID = ArgSelectionID
        .ChoseName = ArgChoseName
    End With
End Property

'***************************************************************************************************
'* 機能　　：header要素の各属性にセットしたものをtoast要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID            ：このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。(省略時：このExcelのファイル名)
'            ArgArguments     ：起動スキーマ。URLやファイルパスなど(省略時：このExcelの保存フォルダパス)
'            ArgActivationType：このヘッダーがクリックされた場合に使用するアクティブ化の種類
'            ArgTitle         ：ヘッダーのタイトルです。(推奨：このExcelのBook名　ThisWorkbook.Name)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ArgTitle を vbnullstring で渡すとこの header要素自体を除去できます。
'* 注意事項：Microsoft の思想設計上、ActivationType は、Protocol が推奨されてます。
'***************************************************************************************************
Public Property Let SetToastHeader(Optional ByVal ArgID As String, Optional ByVal ArgArguments As String, Optional ByVal ArgActivationType As EnumToastActivationType = EnumToastActivationType.taProtocol, ByVal ArgTitle As String)
    With ElementHeader
        '省略時：このExcelのファイル名
        If StrPtr(ArgID) Then .id = ArgID Else .id = ThisWorkbook.Name
        
        '省略時：このExcelの保存フォルダパス
        If StrPtr(ArgArguments) Then .Arguments = ArgArguments Else .Arguments = ThisWorkbook.Path

        '引数の中身そのまんま
        .ActivationType = SetToastActivationType(ArgActivationType)
        .Title = ArgTitle
    End With
End Property



'***************************************************************************************************
' ■■■ Toastを構成するスキーマ情報を、メンバ変数を基に、動的生成します ■■■
'***************************************************************************************************
'* 返り値　 ：トースト スキーマのxmlContentsが返る
'* 引数　　 ：DataBindingMode   True ：トーストの更新ができるような値の指定を行います(例：{UpdateParameter})
'*                              False：値を直接記入します。スケジュール通知や、リマインド通知といった更新処理が使えないシチュエーションに指定します
'---------------------------------------------------------------------------------------------------
'* 注意事項 ：「Microsoft XML. v6.0」の参照設定が必要です
'* URL      ：https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/schema-root
'***************************************************************************************************
Private Function MakeXmlToastContentBuilder(DataBindingMode As Boolean) As String
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
    'group要素
    Const NameNode01_01_04__group As String = "group"

    '--------------------group要素内で使える要素関係の宣言--------------------
    Const NameNode01_01_04_01__subgroup As String = "subgroup"
    'subgroub用text要素
    Const NameNode01_01_04_01_01__text As String = "text"
    'subgroub用text要素
    Const NameNode01_01_04_01_02__image As String = "image"
    
    '--------------------actions要素内で使える要素関係の宣言--------------------
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



    '-------------------------------------------------------visual要素を設定-------------------------------------------------------
    'ノード：visual  を作成してルートノードの子ノードとして追加します。
    Set Node01__visual = RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01__visual, ""))

    'ノード：binding  を作成して「visual」ノードの子ノードとして追加します。
    Set Node01_01__binding = Node01__visual.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01__binding, ""))



    '-------------------------------------------------------binding要素を設定-------------------------------------------------------
    With Node01_01__binding
        '属性ノード：template  を作成して値を設定し、先程の「binding」ノードの属性に追加します。
        .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "template", "")).NodeValue = ElementBinding.template

        '属性ノード：baseUri  を作成して値を設定し、先程の「binding」ノードの属性に追加します。
        If StrPtr(ElementBinding.Baseuri) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "baseUri", "")).NodeValue = ElementBinding.Baseuri
    End With



    '-------------------------------------------------------image要素関係の設定-------------------------------------------------------
    'ノード：image  を作成して「binding」ノードの子ノードとして追加します。
    With Node01_01__binding
        'ロゴ(アイコン)画像が設定してあったら、1つ目の「image」ノードを追加します
        If StrPtr(ElementsImage.AppLogo.src) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_01__image, ""))
                '属性ノード：addImageQuery  に値があったら、先ほどのノード：image  の属性に追加します
                If ElementsImage.AppLogo.AddImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsImage.AppLogo.AddImageQuery)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

                '属性ノード：alt  に値があったら、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.AppLogo.alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.AppLogo.alt

                '属性ノード：src  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementsImage.AppLogo.src

                '属性ノード：placement  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placement", "")).NodeValue = ElementsImage.AppLogo.Placement

                '円にトリミングするフラグがあったら、属性ノード：hint-crop  に値を設定したのを、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.AppLogo.HintCrop) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-crop", "")).NodeValue = ElementsImage.AppLogo.HintCrop
            End With
        End If

        'インライン画像が設定してあったら、2つ目の「image」ノードを追加します
        If StrPtr(ElementsImage.Inline.src) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_01__image, ""))
                '属性ノード：addImageQuery  に値があったら、先ほどのノード：image  の属性に追加します
                If ElementsImage.Inline.AddImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsImage.Inline.AddImageQuery)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

                '属性ノード：alt  に値があったら、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.Inline.alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.Inline.alt

                '属性ノード：src  に値を設定したのを、先ほどのノード：image  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementsImage.Inline.src

                '円にトリミングするフラグがあったら、属性ノード：hint-crop  に値を設定したのを、先ほどのノード：image  の属性に追加します
                If StrPtr(ElementsImage.Inline.HintCrop) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-crop", "")).NodeValue = ElementsImage.Inline.HintCrop
            End With
        End If

        'ヒーロー画像が設定してあったら、3つ目の「image」ノードを追加します
        If StrPtr(ElementsImage.Hero.src) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_01__image, ""))
                '属性ノード：addImageQuery  に値があったら、先ほどのノード：image  の属性に追加します
                If ElementsImage.Hero.AddImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsImage.Hero.AddImageQuery)   'フラグ値は、全て小文字でないと正しく機能しないので変換をかまします

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
                '先ほどのノード：text  にバインディング名を設定します
                '※スケジュール通知の場合は、バインディング機能が使えないため、そのまま記入します
                If DataBindingMode Then .Text = DataBindingTopText_Title Else .Text = ElementsText.Title.Value
                
                 '属性ノード：hint-callScenarioCenterAlign  に値を設定してあったら、先ほどのノード：text  の属性に追加します
                If ElementsText.Title.Attribute.HintCallScenarioCenterAlign Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-callScenarioCenterAlign", "")).NodeValue = LCase(ElementsText.Title.Attribute.HintCallScenarioCenterAlign)
            End With
        End If

        '内容を設定してあったら、2つ目の「text」ノードを追加を追加します
        If StrPtr(ElementsText.Body.Value) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_02__text, ""))
                '先ほどのノード：text  にバインディング名を設定します
                '※スケジュール通知の場合は、バインディング機能が使えないため、そのまま記入します
                If DataBindingMode Then .Text = DataBindingTopText_Contents Else .Text = ElementsText.Body.Value

                 '属性ノード：hint-callScenarioCenterAlign  に値を設定してあったら、先ほどのノード：text  の属性に追加します
                If ElementsText.Body.Attribute.HintCallScenarioCenterAlign Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-callScenarioCenterAlign", "")).NodeValue = LCase(ElementsText.Body.Attribute.HintCallScenarioCenterAlign)
            End With
        End If

        'フッターを設定してあったら、3つ目の「text」ノードを追加します
        If StrPtr(ElementsText.Attribution.Value) Then
            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_02__text, ""))
                '先ほどのノード：text  にバインディング名を設定します
                '※スケジュール通知の場合は、バインディング機能が使えないため、そのまま記入します
                If DataBindingMode Then .Text = DataBindingTopText_Attribute Else .Text = ElementsText.Attribution.Value

                '属性ノード：placement  に値を設定したのを、先ほどのノード：text  の属性に追加します
                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placement", "")).NodeValue = ElementsText.Attribution.Attribute.Placement

                '属性ノード：hint-callScenarioCenterAlign  に値を設定してあったら、先ほどのノード：text  の属性に追加します
                If ElementsText.Attribution.Attribute.HintCallScenarioCenterAlign Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-callScenarioCenterAlign", "")).NodeValue = LCase(ElementsText.Attribution.Attribute.HintCallScenarioCenterAlign)
            End With
        End If
    End With



    '-------------------------------------------------------group,subgroup要素関係の設定-------------------------------------------------------
    'subgroup要素が定義されてたら、挿入します。
    Dim i As Integer, j As Integer
    If SubgroupTotal >= 0 Then
        'ノード：group  を作成して「binding」ノードの子ノードとして追加します。
        With Node01_01__binding.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_04__group, ""))
            '要素位置格納用
            Dim GetPos As Long
        
            'subgroup要素が定義されている分、ノード：subgroup  を作成して「group」ノードの子ノードとして追加します。
            For i = 0 To SubgroupTotal
                With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_04_01__subgroup, ""))
                    '各属性
                    If StrPtr(ElementsSubgroup(i).HintWeight) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-weight", "")).NodeValue = ElementsSubgroup(i).HintWeight
                    If StrPtr(ElementsSubgroup(i).HintTextStacking) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-textStacking", "")).NodeValue = ElementsSubgroup(i).HintTextStacking
                
                
                    '「text」、「image」を「subgroup」ノードの子ノードとして順番に追加します。
                    For j = 0 To UBound(ElementsSubgroup(i).QueueOrder)
                        '現在の順番での要素位置を取得
                        GetPos = ElementsSubgroup(i).QueueOrder(j).TargetElementNo

                        '画像かテキストを判別
                        If ElementsSubgroup(i).QueueOrder(j).UseText Then
                            'text要素を挿入します
                            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_04_01_01__text, ""))
                                '値
                                .Text = ElementsSubgroup(i).QueueText(GetPos).Value
                            
                                '各属性
                                If StrPtr(ElementsSubgroup(i).QueueText(GetPos).HintStyle) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-style", "")).NodeValue = ElementsSubgroup(i).QueueText(GetPos).HintStyle
                                If ElementsSubgroup(i).QueueText(GetPos).HintWrap Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-wrap", "")).NodeValue = LCase(ElementsSubgroup(i).QueueText(GetPos).HintWrap)
                                If ElementsSubgroup(i).QueueText(GetPos).HintMaxLines Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-maxLines", "")).NodeValue = ElementsSubgroup(i).QueueText(GetPos).HintMaxLines
                                If ElementsSubgroup(i).QueueText(GetPos).HintMinLines Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-minLines", "")).NodeValue = ElementsSubgroup(i).QueueText(GetPos).HintMinLines
                                If StrPtr(ElementsSubgroup(i).QueueText(GetPos).HintAlign) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-align", "")).NodeValue = ElementsSubgroup(i).QueueText(GetPos).HintAlign
                                If StrPtr(ElementsSubgroup(i).QueueText(GetPos).Language) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "language", "")).NodeValue = ElementsSubgroup(i).QueueText(GetPos).Language
                            End With
                        Else
                            'image要素を挿入します
                            With .appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_04_01_02__image, ""))
                                '各属性
                                .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "src", "")).NodeValue = ElementsSubgroup(i).QueueImage(GetPos).src
                                
                                If StrPtr(ElementsSubgroup(i).QueueImage(GetPos).HintCrop) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-crop", "")).NodeValue = ElementsSubgroup(i).QueueImage(GetPos).HintCrop
                                If ElementsSubgroup(i).QueueImage(GetPos).HintRemoveMargin Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-removeMargin", "")).NodeValue = LCase(ElementsSubgroup(i).QueueImage(GetPos).HintRemoveMargin)
                                If StrPtr(ElementsSubgroup(i).QueueImage(GetPos).HintAlign) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "hint-align", "")).NodeValue = ElementsSubgroup(i).QueueImage(GetPos).HintAlign
                                If StrPtr(ElementsSubgroup(i).QueueImage(GetPos).AlternateText) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alternateText", "")).NodeValue = ElementsSubgroup(i).QueueImage(GetPos).AlternateText
                                If ElementsSubgroup(i).QueueImage(GetPos).AddImageQuery Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "addImageQuery", "")).NodeValue = LCase(ElementsSubgroup(i).QueueImage(GetPos).AddImageQuery)
                            End With
                        End If
                    Next
                End With
            Next
        End With
    End If



    '-------------------------------------------------------progress要素関係の設定-------------------------------------------------------
    'スケジュール通知じゃないかつ、属性ノード：status  に値を設定してあったら「progress」ノードを追加します(Microsoft ドキュメント準拠)
    If StrPtr(ElementProgress.Status) And DataBindingMode Then
        With Node01_01__binding.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode01_01_03__progress, ""))
            '属性ノード：title  にバインディング設定したのを、先ほどのノード：progress  の属性に追加します
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "title", "")).NodeValue = DataBindingProgress_Title

            '属性ノード：status  にバインディング設定したのを、先ほどのノード：progress  の属性に追加します
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "status", "")).NodeValue = DataBindingProgress_Status

            '属性ノード：value  にバインディング設定したのを、先ほどのノード：progress  の属性に追加します
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "value", "")).NodeValue = DataBindingProgress_Value

            '属性ノード：valueStringOverride  に値を設定してあったら、先ほどのノード：progress  の属性に追加します
            If StrPtr(ElementProgress.ValueStringOverride) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "valueStringOverride", "")).NodeValue = DataBindingProgress_ValueString
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
    If StrPtr(ElementAction(1).Content) Or StrPtr(ElementInput(1).id) Then
        Set Node03__actions = RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03__actions, ""))


        '-------------------------------------------------------input要素を設定-------------------------------------------------------
        '各属性ノードを作成して値を設定し、「input」ノードの属性に追加します。
        '5つまで、繰り返す
        For i = 1 To MaxElementInputNum
            '属性ノード：ID  がない場合は、作成をSKIPします
            If StrPtr(ElementInput(i).id) Then
                Set Node03_01__input = Node03__actions.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03_01__input, ""))
                With Node03_01__input.Attributes
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementInput(i).id
                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "type", "")).NodeValue = ElementInput(i).Type
                    If StrPtr(ElementInput(i).PlaceHolderContent) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "placeHolderContent", "")).NodeValue = ElementInput(i).PlaceHolderContent
                    If StrPtr(ElementInput(i).Title) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "title", "")).NodeValue = ElementInput(i).Title
                    If StrPtr(ElementInput(i).DefaultInput) Then .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "defaultInput", "")).NodeValue = ElementInput(i).DefaultInput

                    '現在のInput要素にて、type属性が"selection"なら、要素番号に対応するSelection要素配列を挿入する
                    If ElementInput(i).Type = "selection" Then
                        For j = 1 To MaxElementSelectionNum
                            'Selection要素のID属性がない場合、挿入をSKIPする
                            If StrPtr(ElementInput(i).Value(j).SelectionID) Then
                                With Node03_01__input.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03_01_01__selection, "")).Attributes
                                    .setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementInput(i).Value(j).SelectionID
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
                    '.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "type", "")).NodeValue = ElementAction(i).Type
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
    'header要素に、Title属性が設定されてあったら、ノード：header  を作成してルートノードの子ノードとして追加します。
    If StrPtr(ElementHeader.Title) Then
        With RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode04__header, ""))
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementHeader.id
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
'               True ：カスタマイズしたDllで処理する際に指定。AppUserModelIDをそのまま返します
'               False：PowerShellコマンドで実行する際に指定します。変数格納用の「$AppId」が追加されます(規定)
'---------------------------------------------------------------------------------------------------
'* 機能説明：ネット上の画像も使用して通知を表示するかを設定します
'* 注意事項：厳密には、Settingがc:internetを含むAppIDでないとインターネット上の画像が使えない
'**************************************************************************************************
Private Function DecisionAppUserModelID(Optional UseDll As Boolean) As String
    '任意のAppIDの設定があったら、そちらを優先します。
    If StrPtr(AppUserModelID_Any) Then
        ToastConfig.AppUserModelID = AppUserModelID_Any

    '未設定の場合は、ExcelのAppUserModelIDをセットします。
    Else
        ToastConfig.AppUserModelID = AppUserModelID_Excel
    End If



    'AppIDを決めるコマンド文字列を返す
    If UseDll Then
        DecisionAppUserModelID = ToastConfig.AppUserModelID
    Else
        DecisionAppUserModelID = "$AppId = '" & ToastConfig.AppUserModelID & Chr(39)
    End If

End Function

'***************************************************************************************************
'* 機能　　：「System.Collections.Generic.Dictionary」を利用して、可変値部分を設定、更新を行うコマンド文字列を生成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：UpdateMode
'               True    更新時は、これに設定
'               False   最初に進捗Toastを発行する場合は、これを設定
'               ※powershell内変数の格納先を変更する都合上、このフラグが必要です。
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

    'Progress要素
    With ElementProgress
        'タイトル属性を設定するコマンド文字列
        Dim SetDataBindingProgressBar_Title               As String: SetDataBindingProgressBar_Title = "$Dictionary.Add('" & Mid(DataBindingProgress_Title, 2, Len(DataBindingProgress_Title) - 2) & "', '" & .Title & "')"

        '進行状況バーの値属性を設定するコマンド文字列
        '※0未満になってると、「Indeterminate」扱いとなります。
        Dim SetDataBindingProgressBar_Value               As String:
        If .Value >= 0 Then
            SetDataBindingProgressBar_Value = "$Dictionary.Add('" & Mid(DataBindingProgress_Value, 2, Len(DataBindingProgress_Value) - 2) & "', '" & .Value & "')"
        Else
            SetDataBindingProgressBar_Value = "$Dictionary.Add('" & Mid(DataBindingProgress_Value, 2, Len(DataBindingProgress_Value) - 2) & "', 'Indeterminate')"
        End If

        '割合を示す文字列属性を設定するコマンド文字列
        '  Null文字として渡された場合、「progress」要素に「valueStringOverride」属性自体を挿入しないようにします。これにより規定の"70%" などの文字が表示されます。
        '  ただし、更新時に改めて、「valueStringOverride」属性を指定しても、無視されます。
        Dim SetDataBindingProgressBar_ValueStringOverride As String
        If StrPtr(.ValueStringOverride) Then SetDataBindingProgressBar_ValueStringOverride = "$Dictionary.Add('" & Mid(DataBindingProgress_ValueString, 2, Len(DataBindingProgress_ValueString) - 2) & "', '" & .ValueStringOverride & "')"

        '進行状況バーの下の左側に表示されるステータス文字列属性を設定するコマンド文字列
        Dim SetDataBindingProgressBar_Status              As String: SetDataBindingProgressBar_Status = "$Dictionary.Add('" & Mid(DataBindingProgress_Status, 2, Len(DataBindingProgress_Status) - 2) & "', '" & .Status & "')"
    End With

    '最上位テキスト
    With ElementsText
        Dim SetDataBindingTopText_Title     As String: SetDataBindingTopText_Title = "$Dictionary.Add('" & Mid(DataBindingTopText_Title, 2, Len(DataBindingTopText_Title) - 2) & "', '" & .Title.Value & "')"
        Dim SetDataBindingTopText_Contents  As String: SetDataBindingTopText_Contents = "$Dictionary.Add('" & Mid(DataBindingTopText_Contents, 2, Len(DataBindingTopText_Contents) - 2) & "', '" & .Body.Value & "')"
        Dim SetDataBindingTopText_Attribute As String: SetDataBindingTopText_Attribute = "$Dictionary.Add('" & Mid(DataBindingTopText_Attribute, 2, Len(DataBindingTopText_Attribute) - 2) & "', '" & .Attribution.Value & "')"
    End With

    '「$Dictionary」をAddするコマンド文字列を返す
    Binding_AdaptiveProgressBar = WorksheetFunction.TextJoin(";", True, New__System_Collections_Generic__Dictionary, SetDataBindingProgressBar_Title, SetDataBindingProgressBar_Value, SetDataBindingProgressBar_ValueStringOverride, SetDataBindingProgressBar_Status, SetDataBindingTopText_Title, SetDataBindingTopText_Contents, SetDataBindingTopText_Attribute, Set__Windows_UI_Notifications__ToastNotificationData)
End Function




'***************************************************************************************************
'               ■■■ [Windows.UI.Notifications]に関する動的設定 ■■■
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
'* 機能　　：削除対象のToastTag等を指定して、通知をアクションセンターから削除します
'---------------------------------------------------------------------------------------------------
'* 返り値　：更新対象のToastTagを指定するコマンド文字列
'* 引数　　：ArgGroup       通知内容を更新させるToastGroup
'            ArgTag         通知内容を更新させるToastGroup内のToastTag
'            AppUserModelID アプリID
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationhistory.remove
'***************************************************************************************************
Private Function Run__Windows_UI_Notifications__ToastRemove(ArgTag As String, ArgGroup As String, AppUserModelID As String) As String
    Run__Windows_UI_Notifications__ToastRemove = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::History.Remove('" & WorksheetFunction.TextJoin("','", False, ArgTag, ArgGroup, AppUserModelID) & "')"
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
'                           ■■■ URLDownloadToFile  に関するメソッド ■■■
'***************************************************************************************************
'* 機能　　：インターネット上の画像ファイルをダウンロードし、トーストの画像ソースとして使用できる準備をします
'---------------------------------------------------------------------------------------------------
'* 返り値　：保存先ローカルパス
'* 引数　　：SourceURI          画像イメージURL
'---------------------------------------------------------------------------------------------------
'* 機能説明：httpソースが使えないAppUserModelIDでも、WindowsAPI:URLDownloadToFile で擬似的にできるようにします。
'* 注意事項：任意の保存先指定は設けません。「C:\Users\XXX\AppData\Local\Temp」固定です。
'          ：「Windows Terminal」等のマニフェストにインターネット機能があるパッケージアプリのAppUserModelIDの場合は、ToastConfig.AllowUse_InternetImageFlag を False にしないと予期せぬ挙動を起こすかもしれません
'***************************************************************************************************
Private Function UseHttpImageCheck(Source As String) As String
    'ネット上の画像を使用するように設定しており尚且つ、httpソースとして指定されている場合、次の処理に移ります
    Dim BaseUriHttpSourceFlag As Boolean: BaseUriHttpSourceFlag = (Left(ElementBinding.Baseuri, 4) = "http")
    Dim HttpSourceFlag As Boolean: HttpSourceFlag = (Left(Source, 4) = "http")
    If ToastConfig.AllowUse_InternetImageFlag And (HttpSourceFlag Or BaseUriHttpSourceFlag) Then

        'Sourceが空なら、即抜け
        If Source = "" Then
            UseHttpImageCheck = Source
            Exit Function
        End If

        'URLの「/」区切りの最後の部分をファイル名として指定
        Dim SaveImageFileName As String, tmp
        tmp = Split(Source, "/")
    
        '保存時の禁則文字を防ぐため、URLエンコードを使います。
        SaveImageFileName = WorksheetFunction.EncodeURL(tmp(UBound(tmp)))

        '規定のtempフォルダパスを取得し、保存先パスを取得します
        Dim TargetFullPass As String: TargetFullPass = WorksheetFunction.TextJoin("\", True, Environ("tmp"), SaveImageFileName)

        'httpソースのベースURIになってたら、連結します
        If BaseUriHttpSourceFlag And Not (HttpSourceFlag) Then Source = ElementBinding.Baseuri & Source

        'キャッシュが消えてたら再度ダウンロードして保存
        If Dir(TargetFullPass, vbNormal) = "" Then
            'ダウンロード開始
            URLDownloadToFile 0, Source, TargetFullPass, 0, 0
        End If
    
        '返却
        UseHttpImageCheck = TargetFullPass
    Else
        'そのまんま返却
        UseHttpImageCheck = Source
    End If
End Function



'***************************************************************************************************
'         ■■■ コマンドプロンプトで、Toast制御させるコマンド文字列を返すメソッド ■■■
'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastNotifierShow
'* 機能　　：引数に渡された値で、トースト通知を表示するコマンド文字列を返します。指定日時に通知するスケジュール機能も対応します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、トーストを表示させるコマンド文字列を返します
'* 引数　　：ToastTag           ToastTagを指定します。複数通知を提供する際に必要です
'            CollectionID       トーストコレクションIDがあるなら、それを指定
'            ScheduleDate       指定日時に通知したい場合は、「Date型」として指定することでスケジュール通知として登録されます。省略時、即時通知とします
'            ExpirationDate     アクションセンターから自動削除される日時を指定。最大30sの遅延バグが発生します。
'            Suppress           Trueで、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'            ExpiresOnReboot    True で、再起動すると同時に、トースト通知が消えます
'---------------------------------------------------------------------------------------------------
'* 注意事項：あくまで「コマンド実行の文字列」を返すだけで、呼び出すだけでは実行されません。別途、Shell関数や「Windows Script Host Object Model」を参照設定してRUNコマンドを実行することで、Toastを発行できます。
'**************************************************************************************************
Public Property Get GenerateCmd_ToastNotifierShow(ByVal ToastTag As String, Optional ByVal CollectionID As String, Optional ByVal ScheduleDate As Date, Optional ByVal ExpirationDate As Date, Optional ByVal Suppress As Boolean, Optional ExpiresOnReboot As Boolean) As String
    'データバインディングを使用するか判定する
    Dim UseDataBinding As Boolean, i As Long
    
    'スケジュール設定：あり
    If ScheduleDate > 0 Then
        'データバインディングを、無効化する
        UseDataBinding = False
            
        '過去になっている場合、VBAエラーを起こして、処理を停止させます
        If Now() >= ScheduleDate Then Err.Raise Number:=vbObjectError + 1, Description:=vbObjectErrorID_0001
    
    'スケジュール設定：なし
    Else
        'データバインディングを、有効化する
        UseDataBinding = True
        
        'ただし、action要素に特定の組み合わせを検知したら、データバインディングを、無効化する
        For i = 1 To MaxElementActionNum
            If ElementAction(i).ActivationType = SetToastActivationType(taSystem) And ElementAction(i).Arguments = ActionArguments再通知 Then UseDataBinding = False: Exit For
        Next
    End If
    
    
    With ToastConfig
        '設定値をまとめる
        '.AppUserModelID = DecisionAppUserModelID(True)     '後続で、設定
        '.AllowUse_InternetImageFlag                        '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(UseDataBinding)
        
        .Tag = ToastTag
        .Group = Application.hwnd
        .Schedule_ID = "ExcelSchedule"
        .CollectionID = CollectionID
        .ExpiresOnReboot = ExpiresOnReboot
        .SuppressPopup = Suppress
        .Schedule_DeliveryTime = ScheduleDate
        .ExpirationTime = ExpirationDate
        
        
        
        '共通の設定コマンドを生成
        ' 1.事前に設定したメンバ変数から、xmlコンテンツを生成し、それを読み込む。
        ' 2.読み込んだxmlコンテンツから、ToastNotificationの構造を設定
        ' 3.サイレント通知フラグを設定
        ' 4.有効期限を設定
        ' 5.識別Group,Tagを設定
        ' 6.データバインディングを設定
        GenerateCmd_ToastNotifierShow = ActionPS & WorksheetFunction.TextJoin(";", True, _
            SetToastShemaXml(.XmlSchema), ReadXml, _
            CreateObject__Windows_UI_Notifications(.Schedule_DeliveryTime, .Schedule_ID), _
            Set__Windows_UI_Notifications__ToastNotificationSuppressPopup(.SuppressPopup), _
            Set__Windows_UI_Notifications__ToastNotificationExpirationTime(.ExpirationTime, .Schedule_DeliveryTime), _
            Set__Windows_UI_Notifications__ToastNotificationGroup(.Group), Set__Windows_UI_Notifications__ToastNotificationTag(.Tag), _
            Binding_AdaptiveProgressBar(False))


        'コレクションIDの有無に応じで、設定コマンドを切り替える
        If StrPtr(CollectionID) Then
            '6.コレクションIDを指定して、Toast通知実行
            GenerateCmd_ToastNotifierShow = WorksheetFunction.TextJoin(";", False, GenerateCmd_ToastNotifierShow, _
                Run__Windows_UI_Notifications__ToastNotifierShow_Collection(.CollectionID)) & Chr(34)
        Else
            '6.AppIDを設定
            '7.Toast通知実行
            GenerateCmd_ToastNotifierShow = WorksheetFunction.TextJoin(";", False, GenerateCmd_ToastNotifierShow, _
                DecisionAppUserModelID, _
                Run__Windows_UI_Notifications__ToastNotifier(.Schedule_DeliveryTime)) & Chr(34)
        End If
    End With
End Property

'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastNotifierUpdate
'* 機能　　：引数に渡された値で、データバインディングを使った更新をします。
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、既存の進捗状況ゲージ込みのトーストを更新させるコマンド文字列を返します
'* 引数　　：ToastTag       ToastTagを指定します。複数通知を提供する際に必要です
'---------------------------------------------------------------------------------------------------
'* 機能説明：既にトーストの進行状況バーを表示しているとき、その時に指定したタグを指定することで、更新が可能です。
'* 注意事項：コレクションIDによる更新は、ここでは非対応です(ﾂｸﾙﾉﾒﾝﾄﾞｲ)
'***************************************************************************************************
Public Property Get GenerateCmd_ToastNotifierUpdate(ByVal ToastTag As String) As String
    With ToastConfig
        '設定値をまとめる
        '.AppUserModelID = DecisionAppUserModelID(True)     '後続で、設定
        '.AllowUse_InternetImageFlag                        '事前に、設定済み
        .XmlSchema = vbNullString

        .Tag = ToastTag
        .Group = Application.hwnd
        .Schedule_ID = vbNullString
        .CollectionID = vbNullString
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
       

        'トーストの進行状況バーを更新させるコマンド文字列を生成
        ' 1.データバインディングを設定
        ' 2.AppIDを設定
        ' 3.更新用Toast通知実行
        GenerateCmd_ToastNotifierUpdate = ActionPS & WorksheetFunction.TextJoin(";", True, _
            Binding_AdaptiveProgressBar(True), _
            DecisionAppUserModelID, _
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
        '.AppUserModelID = DecisionAppUserModelID(True)     '後続で、設定
        '.AllowUse_InternetImageFlag                        '事前に、設定済み
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
            DecisionAppUserModelID, _
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
        DecisionAppUserModelID, _
        RemoveMode) & Chr(34)
End Property

'***************************************************************************************************
'* 処理名　：GenerateCmd_ToastNotificationHistoryRemove
'* 機能　　：ToastTagに紐づく通知をアクションセンターから削除します
'---------------------------------------------------------------------------------------------------
'* 返り値　：ToastTagに紐づく通知をアクションセンターから削除するコマンド文字列
'* 引数　　：ToastTag   削除するToastTagを指定します。
'***************************************************************************************************
Public Property Get GenerateCmd_ToastNotificationHistoryRemove(RemoveTag As String) As String
    With ToastConfig
        '設定値をまとめる
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                      '事前に、設定済み

        .Tag = RemoveTag
        .Group = Application.hwnd

        '削除コマンドを生成
        GenerateCmd_ToastNotificationHistoryRemove = ActionPS & Run__Windows_UI_Notifications__ToastRemove(.Tag, .Group, .AppUserModelID) & Chr(34)
    End With
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
    SetIToastActions("", ActionArguments再通知, taSystem, , , , ReminderID) = 1
    SetIToastActions("", ActionArguments解除, taSystem) = 2
End Sub



'***************************************************************************************************
'             ■■■ 専用DLL C++からWindows.UI.Notificationsを呼び出すメソッド ■■■
'***************************************************************************************************
'* 処理名　：RunDll_ToastNotifierShow
'* 機能　　：引数に渡された値で、単純なトースト通知を表示します。指定日時に通知するスケジュール機能も対応します
'---------------------------------------------------------------------------------------------------
'* 返り値　：引数に渡された値で、トーストを表示させるコマンド文字列を返します
'* 引数　　：ToastTag           ToastTagを指定します。複数通知を提供する際に必要です
'            CollectionID       トーストコレクションIDがあるなら、それを指定
'            ScheduleDate       指定日時に通知したい場合は、「Date型」として指定することでスケジュール通知として登録されます。省略時、即時通知とします
'            ExpirationDate     アクションセンターから自動削除される日時を指定。最大30sの遅延バグが発生します。
'            Suppress           True で、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'            ExpiresOnReboot    True で、再起動すると同時に、トースト通知が消えます
'---------------------------------------------------------------------------------------------------
'* 注意事項：これを実行すると、即座にDll内で処理を行います。また、日時系データを扱う場合は、UTC基準になるので注意
'**************************************************************************************************
Public Sub RunDll_ToastNotifierShow(ByVal ToastTag As String, Optional ByVal CollectionID As String, Optional ByVal ScheduleDate As Date, Optional ByVal ExpirationDate As Date, Optional ByVal Suppress As Boolean, Optional ExpiresOnReboot As Boolean)
    'データバインディングを使用するか判定する
    Dim UseDataBinding As Boolean, i As Long
    
    'スケジュール設定：あり
    If ScheduleDate > 0 Then
        'データバインディングを、無効化する
        UseDataBinding = False
            
        '過去になっている場合、VBAエラーを起こして、処理を停止させます
        If Now() >= ScheduleDate Then Err.Raise Number:=vbObjectError + 1, Description:=vbObjectErrorID_0001
    
    'スケジュール設定：なし
    Else
        'データバインディングを、有効化する
        UseDataBinding = True
        
        'ただし、action要素に特定の組み合わせを検知したら、データバインディングを、無効化する
        For i = 1 To MaxElementActionNum
            If ElementAction(i).ActivationType = SetToastActivationType(taSystem) And ElementAction(i).Arguments = ActionArguments再通知 Then UseDataBinding = False: Exit For
        Next
    End If

    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = MakeXmlToastContentBuilder(UseDataBinding)
        
        .Tag = ToastTag
        .Group = Application.hwnd
        .Schedule_ID = "ExcelSchedule"
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = ExpiresOnReboot
        .SuppressPopup = Suppress
        
        .Schedule_DeliveryTime = ScheduleDate
        .ExpirationTime = ExpirationDate
    End With

    'データバインディングを設定しつつ、Dll内の「ShowToastNotification」関数を実行します
    ShowToastNotification ParamsConvertForDLL(ToastConfig), ReturnDataBinding
End Sub

'***************************************************************************************************
'* 処理名　：RunDll_ToastNotifierUpdate
'* 機能　　：引数に渡された値で、トーストの進行状況バーを更新します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.notificationupdateresult
'            0(Succeeded)               通知が更新されました。
'            1(Failed)                  通知の更新に失敗しました。
'            2(NotificationNotFound)    指定した通知が見つかりませんでした。
'            ※上記以外                  予期せぬエラー

'* 引数　　：ToastTag                   ToastTagを指定します。以後の進捗状況を更新する際に必要です
'---------------------------------------------------------------------------------------------------
'* 機能説明：既にトーストの進行状況バーを表示しているとき、その時に指定したタグを指定することで、更新が可能です。Dllの場合、返り値に対応し、これに応じた分岐処理も可能です。
'***************************************************************************************************
Public Property Get RunDll_ToastNotifierUpdate(ByVal ToastTag As String, Optional ByVal CollectionID As String) As Long
    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

        .Tag = ToastTag
        .Group = Application.hwnd
        .Schedule_ID = vbNullString
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With


    'Dll内の「UpdateToastNotification」関数を実行します
    RunDll_ToastNotifierUpdate = UpdateToastNotification(ParamsConvertForDLL(ToastConfig), ReturnDataBinding)
End Property

'**************************************************************************************************
'* 処理名　：RunDll_ToastNotificationHistoryRemove
'* 機能　　：タグから、トースト通知を削除します
'---------------------------------------------------------------------------------------------------
'* 引数　　：CollectionID       コレクション通知IDの方を削除したい時に指定
'            ToastTag           削除するToastTagを指定します。指定しない場合、アクティブExcelで出したトースト通知を全て削除します。
'            AllDel             True で、AppUserModelID に紐づく全てのトースト通知を削除します
'---------------------------------------------------------------------------------------------------
'* 注意事項：AllDel が True になると、ToastTag を無視します。
'**************************************************************************************************
Public Sub RunDll_ToastNotificationHistoryRemove(Optional ByVal ToastTag As String, Optional ByVal CollectionID As String, Optional AllDel As Boolean)
    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

        'AllDelモードなら、NULL 指定
        If AllDel Then
            .Tag = vbNullString
            .Group = vbNullString
        Else
            .Tag = ToastTag
            .Group = Application.hwnd
        End If
        .Schedule_ID = vbNullString
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With


    'Dll内の「RemoveToastNotification」関数を実行します
    RemoveToastNotification ParamsConvertForDLL(ToastConfig)
End Sub

'***************************************************************************************************
'* 処理名　：RunDll_ToastCollectionManagerSaveToastCollectionAsync
'* 機能　　：引数で渡された値で、コレクションを使用したトースト通知のグループ化を作成します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：0で、成功扱いです。
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
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

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
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

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
'* 処理名　：RunDll_CheckNotificationSetting
'* 機能　　：トースト通知の表示に関する制限を確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：0~4 の値。意味は公式サイトを御覧ください。
'* 引数　　：CollectionId   コレクションの ID。Collection側の確認をする場合に指定。そうじゃない場合は、AppUserModelID になります。
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.notificationsetting
'***************************************************************************************************
Public Property Get RunDll_CheckNotificationSetting(Optional ByVal CollectionID As String) As Long
    '設定値をまとめる
    With ToastConfig
        .AppUserModelID = DecisionAppUserModelID(True)
        '.AllowUse_InternetImageFlag                     '事前に、設定済み
        .XmlSchema = vbNullString

        .Tag = vbNullString
        .Group = vbNullString
        .Schedule_ID = vbNullString
        .CollectionID = CollectionID
        
        .ExpiresOnReboot = False
        .SuppressPopup = False
        
        .Schedule_DeliveryTime = False
        .ExpirationTime = False
    End With
    
    'チェック
    RunDll_CheckNotificationSetting = CheckNotificationSetting(ParamsConvertForDLL(ToastConfig))
End Property



'***************************************************************************************************
'                               ■■■ wpndatabase.db 操作関連 ■■■
'***************************************************************************************************
'* 機能　　：簡単な引数で、SELECT文を作成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：SELECT文
'* 引数　　：※日本語名まんまのため省略します
'---------------------------------------------------------------------------------------------------
'* 注意事項：WHERE句部分は最大、2セットまでとし、ANDとします
'***************************************************************************************************
Private Function SELECT文生成(取得したい列 As String, 対象テーブル As String, 条件列1 As String, 条件値1 As String, Optional 条件列2 As String, Optional 条件値2 As String) As String
    '条件列2 が指定されてるかで、WHERE句に AND を足す
    If StrPtr(条件列2) Then
        SELECT文生成 = WorksheetFunction.TextJoin(" ", False, "SELECT", 取得したい列, "FROM", 対象テーブル, "WHERE", 条件列1 & "='" & 条件値1 & "'", "AND", 条件列2 & "='" & 条件値2 & "';")
    Else
        SELECT文生成 = WorksheetFunction.TextJoin(" ", False, "SELECT", 取得したい列, "FROM", 対象テーブル, "WHERE", 条件列1 & "='" & 条件値1 & "';")
    End If
End Function

'***************************************************************************************************
'* 機能　　：簡単な引数で、UPDATE文を作成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：UPDATE文
'* 引数　　：※日本語名まんまのため省略します
'---------------------------------------------------------------------------------------------------
'* 注意事項：WHERE句部分は最大、2セットまでとし、ANDとします
'***************************************************************************************************
Private Function UPDATE文生成(設定したい列 As String, ByVal 設定したい値 As String, 対象テーブル As String, 条件列1 As String, 条件値1 As String, Optional 条件列2 As String, Optional 条件値2 As String, Optional NullMode As Boolean) As String
    '条件列2 が指定されてるかで、WHERE句に AND を足す
    If StrPtr(条件列2) Then
        UPDATE文生成 = WorksheetFunction.TextJoin(" ", False, "UPDATE", 対象テーブル, "SET", 設定したい列 & "='" & 設定したい値 & "'", "WHERE", 条件列1 & "='" & 条件値1 & "'", "AND", 条件列2 & "='" & 条件値2 & "';")
    Else
        UPDATE文生成 = WorksheetFunction.TextJoin(" ", False, "UPDATE", 対象テーブル, "SET", 設定したい列 & "='" & 設定したい値 & "'", "WHERE", 条件列1 & "='" & 条件値1 & "';")
    End If
    
    'NULLモードなら空文字時、NULLとして扱います。
    If NullMode Then UPDATE文生成 = Replace(UPDATE文生成, "''", "NULL", , 1)
End Function

'***************************************************************************************************
'* 機能　　：簡単な引数で、テーブル：HandlerSettings 用 INSERT文を作成します
'---------------------------------------------------------------------------------------------------
'* 返り値　：INSERT文
'* 引数　　：※日本語名まんまのため省略します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：テーブル：HandlerSettings　は、SettingKey を動的生成するため、新しい設定の適用に使用します。
'* 注意事項：テーブル：HandlerSettings　専用です
'***************************************************************************************************
Private Function INSERT文生成_HandlerSettings(HandlerId As Long, SettingKey As String, flag As Boolean) As String
    INSERT文生成_HandlerSettings = WorksheetFunction.TextJoin(" ", False, "INSERT", "INTO", "HandlerSettings", "VALUES('" & HandlerId & "',", "'" & SettingKey & "',", "'" & flag * -1 & "');")
End Function

'***************************************************************************************************
'* 機能　　：数値指定で、対応の[SettingKey]の名称を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：対応する[SettingKey]の名称
'* 引数　　：列挙型（列挙体）Enum：Wpndatabase_SettingKey
'***************************************************************************************************
Private Function 対応表_Wpndatabase_SettingKey(Parameter As Wpndatabase_SettingKey) As String
    Dim ResultString As String
    Select Case Parameter
        Case skS_tickle: ResultString = "s:tickle"
        Case skC_tile: ResultString = "c:tile"
        Case skR_badge: ResultString = "r:badge"
        Case skS_lock_toast: ResultString = "s:lock:toast"
        Case skS_tile: ResultString = "s:tile"
        Case skS_lock_badge: ResultString = "s:lock:badge"
        Case skS_banner: ResultString = "s:banner"
        Case skS_lock_tile: ResultString = "s:lock:tile"
        Case skR_tile: ResultString = "r:tile"
        Case skS_badge: ResultString = "s:badge"
        Case skS_audio: ResultString = "s:audio"
        Case skS_voip: ResultString = "s:voip"
        Case skS_stopCloud: ResultString = "s:stopCloud"
        Case skS_listenerEnabled: ResultString = "s:listenerEnabled"
        Case skC_cloud: ResultString = "c:cloud"
        Case skC_toast: ResultString = "c:toast"
        Case skC_internet: ResultString = "c:internet"
        Case skC_badge: ResultString = "c:badge"
        Case skC_ringing: ResultString = "c:ringing"
        Case skC_tickle: ResultString = "c:tickle"
        Case skC_voip: ResultString = "c:voip"
        Case skC_storage_toast: ResultString = "c:storage:toast"
        Case skS_toast: ResultString = "s:toast"
        Case skC_storage_tile: ResultString = "c:storage:tile"
        Case skM_SecondaryTileMigrationComplete: ResultString = "m:SecondaryTileMigrationComplete"
        Case skS_cycle_medium: ResultString = "s:cycle:medium"
        Case skS_cycle_wide: ResultString = "s:cycle:wide"
        Case skS_cycle_large: ResultString = "s:cycle:large"
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select

    '結果を返す
    対応表_Wpndatabase_SettingKey = ResultString
End Function

'***************************************************************************************************
'* 機能　　："AppUserModelID"を基に、[RecordId]を取得します。
'---------------------------------------------------------------------------------------------------
'* 返り値　："AppUserModelID"に基づく、[RecordId]
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・同梱のDLLを経由して、"wpndatabase.db"を SQLite を使って取得します
'            ・HandlerSettings の設定に必要です
'***************************************************************************************************
Private Function GetWpndatabase_HandlerID() As Long
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス

    '[RecordId]を取得する SELECT 文を組む
    Dim SQL_SELECT As String: SQL_SELECT = SELECT文生成("RecordId", "NotificationHandler", "PrimaryId", DecisionAppUserModelID(True))

    '実行して、RecordId を取得する(整数を前提)
    GetWpndatabase_HandlerID = CLng(ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_SELECT)))
End Function

'***************************************************************************************************
'* 機能　　："AppUserModelID"を基に、[WNSId]を取得します。
'---------------------------------------------------------------------------------------------------
'* 返り値　："AppUserModelID"に基づく、[WNSID]
'***************************************************************************************************
Public Property Get GetWpndatabase_WNSId() As String
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス

    '[WNSID]を取得する SELECT 文を組む
    Dim SQL_SELECT As String: SQL_SELECT = SELECT文生成("WNSId", "NotificationHandler", "PrimaryId", DecisionAppUserModelID(True))

    '実行して、WNSID を取得する
    GetWpndatabase_WNSId = ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_SELECT))
End Property

'***************************************************************************************************
'* 機能　　："AppUserModelID"を基に、[WNSId]を更新します。
'---------------------------------------------------------------------------------------------------
'* 引数　  ：・Custom_WNSId       任意の WNSId
'          　・Wpndatabase_WNSId  システムで用意されてるパラメーター
'---------------------------------------------------------------------------------------------------
'* 詳細説明：カスタムな WNSId を指定する際は、({任意のWNSId}) = wsCustom と記述します
'***************************************************************************************************
Public Property Let SetWpndatabase_WNSId(Optional Custom_WNSId As String, WNSId As Wpndatabase_WNSId)
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス

    'ID に紐づくParameterを設定
    Dim WNSId_string As String
    Select Case WNSId
        Case wsSystem: WNSId_string = "System"
        Case wsNonImmersivePackage: WNSId_string = "NonImmersivePackage"
        Case Else: WNSId_string = Custom_WNSId
    End Select

    '"WNSID"を設定する UPDATE 文を組む(空文字の場合は、NULL)
    Dim SQL_UPDATE As String: SQL_UPDATE = UPDATE文生成("WNSId", WNSId_string, "NotificationHandler", "PrimaryId", DecisionAppUserModelID(True), , , True)

    '実行して、WNSID を設定する
    LatestExecuteSQLiteNum = ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_UPDATE))
End Property

'***************************************************************************************************
'* 機能　　："AppUserModelID"を基に、[HandlerType]を取得します。
'---------------------------------------------------------------------------------------------------
'* 返り値　："AppUserModelID"に基づく、[HandlerType]
'***************************************************************************************************
Public Property Get GetWpndatabase_HandlerType() As String
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス

    '"HandlerType"を取得する SELECT 文を組む
    Dim SQL_SELECT As String: SQL_SELECT = SELECT文生成("HandlerType", "NotificationHandler", "PrimaryId", DecisionAppUserModelID(True))

    '実行して、HandlerType を取得する
    GetWpndatabase_HandlerType = ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_SELECT))
End Property

'***************************************************************************************************
'* 機能　　："AppUserModelID"を基に、[HandlerType]を更新します。
'---------------------------------------------------------------------------------------------------
'* 引数　  ：Wpndatabase_HandlerType  列挙型（列挙体）Enum：Wpndatabase_HandlerType
'***************************************************************************************************
Public Property Let SetWpndatabase_HandlerType(HandlerTypeMode As Wpndatabase_HandlerType)
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス
    
    '接頭辞
    Const HandlerType接頭辞 As String = "app:"

    'ID に紐づくParameterを設定
    Dim HandlerType_string As String
    Select Case HandlerTypeMode
        Case htDesktop: HandlerType_string = HandlerType接頭辞 & "desktop"
        Case htImmersive: HandlerType_string = HandlerType接頭辞 & "immersive"
        Case htSystem: HandlerType_string = HandlerType接頭辞 & "system"
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select

    '"HandlerType"を設定する UPDATE 文を組む
    Dim SQL_UPDATE As String: SQL_UPDATE = UPDATE文生成("HandlerType", HandlerType_string, "NotificationHandler", "PrimaryId", DecisionAppUserModelID(True))

    '実行して、HandlerType を設定する
    LatestExecuteSQLiteNum = ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_UPDATE))
End Property

'***************************************************************************************************
'* 機能　　：[HandlerId](NotificationHandler-[AppUserModelID]に紐づくやつ)から、[SettingKey]を基に、[Value]を取得します。
'---------------------------------------------------------------------------------------------------
'* 返り値　：論理値
'* 引数　  ：SettingKey     列挙型（列挙体）Enum：Wpndatabase_SettingKey
'---------------------------------------------------------------------------------------------------
'* 注意事項：SettingKey は、動的生成なので、存在しない SettingKey が来たら、False で INSERT文を実行します
'***************************************************************************************************
Public Property Get Wpndatabase_SettingKeyValue(SettingKeyID As Wpndatabase_SettingKey) As Boolean
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス
    
    '現在のこの Class で設定中の AppUserModelID を基に、HandlerId を特定します
    Dim TargetHandlerId As Long: TargetHandlerId = GetWpndatabase_HandlerID

    '[HandlerId]と[SettingKey]を基に、[Value]を取得する SELECT 文を組む
    Dim SettingKeyName As String: SettingKeyName = 対応表_Wpndatabase_SettingKey(SettingKeyID)
    Dim SQL_SELECT As String: SQL_SELECT = SELECT文生成("Value", "HandlerSettings", "HandlerId", CStr(TargetHandlerId), "SettingKey", SettingKeyName)

    '実行して、Value を取得
    Dim SELECT結果 As String: SELECT結果 = ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_SELECT))
    
    'もし何も得られない場合、そもそもdbにないので、INSERTで設定する。デフォルトは、false とします
    If SELECT結果 = "" Then
        LatestExecuteSQLiteNum = ExecuteSQLite(StrPtr(TargetPath), StrPtr(INSERT文生成_HandlerSettings(TargetHandlerId, SettingKeyName, False)))
        Wpndatabase_SettingKeyValue = False
    Else
       Wpndatabase_SettingKeyValue = CBool(SELECT結果)
    End If
End Property

'***************************************************************************************************
'* 機能　　：[HandlerId](NotificationHandler-[AppUserModelID]に紐づくやつ)から、[SettingKey]を基に、[Value]を設定します。
'---------------------------------------------------------------------------------------------------
'* 引数　  ：SettingKey     列挙型（列挙体）Enum：Wpndatabase_SettingKey
'            SetFlag        論理値
'---------------------------------------------------------------------------------------------------
'* 注意事項：SettingKey は、動的生成なので、存在しない SettingKey が来たら、SetFlag で INSERT文を実行します
'***************************************************************************************************
Public Property Let Wpndatabase_SettingKeyValue(SettingKeyID As Wpndatabase_SettingKey, SetFlag As Boolean)
    'wpndatabase.db の絶対パスを取得
    Dim TargetPath As String: TargetPath = Environ("LOCALAPPDATA") & 設定DB保存相対パス
    
    '現在のこの Class で設定中の AppUserModelID を基に、HandlerId を特定します
    Dim TargetHandlerId As Long: TargetHandlerId = GetWpndatabase_HandlerID

    '[HandlerId]と[SettingKey]を基に、[Value]を設定する UPDATE 文を組む
    Dim SettingKeyName As String: SettingKeyName = 対応表_Wpndatabase_SettingKey(SettingKeyID)
    Dim SQL_UPDATE As String: SQL_UPDATE = UPDATE文生成("Value", SetFlag * -1, "HandlerSettings", "HandlerId", CStr(TargetHandlerId), "SettingKey", SettingKeyName)

    '実行して、Value を設定する
    LatestExecuteSQLiteNum = ExecuteSQLite(StrPtr(TargetPath), StrPtr(SQL_UPDATE))
    
    'もし0件の場合、そもそもdbにないので、INSERTで設定する
    If LatestExecuteSQLiteNum = 0 Then
        LatestExecuteSQLiteNum = ExecuteSQLite(StrPtr(TargetPath), StrPtr(INSERT文生成_HandlerSettings(TargetHandlerId, SettingKeyName, SetFlag)))
    End If
End Property



'***************************************************************************************************
'                               ■■■ レジストリ操作関連 ■■■
'***************************************************************************************************
'* 機能　　：レジストリ　コンピューター\HKEY_CLASSES_ROOT\AppUserModelId\[AppUserModelID名称] へパラメーターを登録し、正しく書き込めたか判定します
'---------------------------------------------------------------------------------------------------
'* 引数　  ：aumid      AppUserModelID
'            valueName  値の名称
'            valueType  登録の型 (列挙型（列挙体）Enum：RegistryType)
'            valueData  登録値
'
'* 返り値  ：結果コード
'                0：成功
'               -1：ユーザーアカウント制御にて許可しましたまたは、既に管理者権限状態で実行されたが、登録/削除の確認が取れなかった。
'               -2：ユーザーアカウント制御にて拒否しました
'               他：エラーコード
'---------------------------------------------------------------------------------------------------
'* 注意事項：ユーザーアカウント制御を利用した管理者権限付与状態のRundll32 の実行を行う関係上、パラメーター区切り文字を使った引数を使用してます。
'            そのため、その区切り文字自体を含めると、正しく機能しません
'***************************************************************************************************
Private Function AttemptRegistryValue(ByVal aumid As String, ByVal valueName As String, ByVal valueType As RegistryType, ByVal valueData As String) As Long
    ' 引数をパイプ区切りで結合
    Dim combinedArgs As String: combinedArgs = WorksheetFunction.TextJoin(パラメーター区切り文字, False, aumid, valueName, valueType, valueData)
    
    ' DLL関数を呼び出して、ユーザーアカウント制御を利用した登録処理
    Dim ResultCode As Long: ResultCode = AttemptToWriteRegistry(0, 0, StrPtr(combinedArgs), 0)
    
    '結果コードに応じた分岐
    If ResultCode = -1 Then
        'ユーザーアカウント制御にて許可したので、存在チェックする
        Dim wmiService As Object, regProvider As Object
        Set wmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default")
        Set regProvider = wmiService.Get("StdRegProv")
        
        ' HKEY_CLASSES_ROOT (HKCR) を起点とする
        Const HKEY_CLASSES_ROOT = &H80000000
    
        '確認用パラメーター用意
        Dim strValue As String  'SZ用
        Dim dwordValue As Long  'DWORD用
        Dim lResult As Long     '実行結果
        Dim keyPath As String: keyPath = "AppUserModelId\" & aumid
        
        '設定型に応じた確認処理
        Select Case valueType
            'DWORD 存在確認
            Case RegistryType.REG_DWORD
                'DWORDとして読み取りを試みる
                lResult = regProvider.GetDWORDValue(HKEY_CLASSES_ROOT, keyPath, valueName, dwordValue)
                
                '結果を返す
                If lResult Then AttemptRegistryValue = ResultCode Else AttemptRegistryValue = lResult
                
            'SZ 存在確認
            Case RegistryType.REG_SZ
                'SZとして読み取りを試みる
                lResult = regProvider.GetStringValue(HKEY_CLASSES_ROOT, keyPath, valueName, strValue)
                
                '結果を返す
                If lResult Then AttemptRegistryValue = ResultCode Else AttemptRegistryValue = lResult
            
            '削除確認
            Case RegistryType.REG_Delete
                'DWORDとして読み取りを試みる
                lResult = regProvider.GetDWORDValue(HKEY_CLASSES_ROOT, keyPath, valueName, dwordValue)
                'あったら、ここで終わり
                If lResult = 0 Then AttemptRegistryValue = ResultCode: Exit Function

                'SZとして読み取りを試みる
                lResult = regProvider.GetStringValue(HKEY_CLASSES_ROOT, keyPath, valueName, strValue)
                'あったら、所定の処理
                If lResult = 0 Then AttemptRegistryValue = ResultCode Else AttemptRegistryValue = 0
            
            Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
        End Select
    Else
        'ユーザーアカウント制御にて拒否したか、他エラーコード
        AttemptRegistryValue = ResultCode
    End If
End Function

'***************************************************************************************************
'* 機能　　：レジストリ　コンピューター\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\[AppUserModelID名称] へパラメーターを登録し、正しく書き込めたか判定します
'---------------------------------------------------------------------------------------------------
'* 引数　  ：aumid      AppUserModelID
'            valueName  値の名称
'            valueType  登録の型 (列挙型（列挙体）Enum：RegistryType)
'            valueData  登録値
'
'* 返り値  ：結果コード
'               0   成功
'               2   見つかりません(削除の時に表れることがある)
'              他   失敗
'---------------------------------------------------------------------------------------------------
'* 詳細説明：こちらは、HKEY_CURRENT_USER 領域のため、VBA標準の機能で実装します。
'***************************************************************************************************
Private Function CurrentUserRegistryValue(ByVal AppUserModelID As String, ByVal valueName As String, ByVal valueType As RegistryType, ByVal valueData As String) As Long
    ' レジストリハイブの定数
    Const HKEY_CURRENT_USER As Long = &H80000001

    '必要な変数を用意
    Dim wmiService As Object
    Dim regProvider As Object
    Dim strKeyPath As String
    Dim lResult As Long
    
    ' 1. WMIサービスに接続
    Set wmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default")
    Set regProvider = wmiService.Get("StdRegProv")
    
    ' 2. HKCU配下のフルパスを組み立てる
    strKeyPath = HKCU配下の通知関連フルパス & AppUserModelID
    
    ' 3. キーが存在しなければ作成する(※削除モードはしない)
    If valueType > REG_Delete Then
        lResult = regProvider.CreateKey(HKEY_CURRENT_USER, strKeyPath)
        If lResult <> 0 Then
            Debug.Print "キーの作成/オープンに失敗: " & strKeyPath
    
            CurrentUserRegistryValue = lResult
            Exit Function
        End If
    End If

    ' 4. 値の型に応じて、適切なWMIメソッドを呼び出す
    Select Case valueType
        Case REG_DWORD: lResult = regProvider.SetDWORDValue(HKEY_CURRENT_USER, strKeyPath, valueName, CLng(valueData))
        Case REG_SZ: lResult = regProvider.SetStringValue(HKEY_CURRENT_USER, strKeyPath, valueName, CStr(valueData))
        Case REG_EXPAND_SZ: lResult = regProvider.SetExpandedStringValue(HKEY_CURRENT_USER, strKeyPath, valueName, CStr(valueData))
        Case REG_Delete: lResult = regProvider.DeleteValue(HKEY_CURRENT_USER, strKeyPath, valueName)
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select

    '結果コードを返す
    CurrentUserRegistryValue = lResult
End Function

'***************************************************************************************************
'* 機能　　：レジストリ登録への簡単なプリセットを提供します
'---------------------------------------------------------------------------------------------------
'* 引数　  ：PresetID   列挙型（列挙体）Enum：RegistrySetting   ここで、登録したいIDを指定します。
'            valueData  登録値
'***************************************************************************************************
 Public Property Let PresetRegistry(PresetID As RegistrySetting, ByVal valueData As String)
    Select Case PresetID
        'HKEY_CLASSES_ROOT 用
        Case RegistrySetting.ShowInSettings: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "ShowInSettings", REG_DWORD, CBool(valueData) * -1)
        Case RegistrySetting.DisplayName: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "DisplayName", REG_SZ, valueData)
        Case RegistrySetting.IconUri: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "IconUri", REG_SZ, valueData)
        Case RegistrySetting.IconBackgroundColor: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "IconBackgroundColor", REG_SZ, valueData)
        
        'HKEY_CURRENT_USER 用
        '--wpndatabase.db 連携必須--
        Case RegistrySetting.Enabled: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "Enabled", REG_DWORD, CBool(valueData) * -1)
        Case RegistrySetting.ShowBanner: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "ShowBanner", REG_DWORD, 0)  '値は常に0とします。
        '--wpndatabase.db 連携不要--
        Case RegistrySetting.SoundFile: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "SoundFile", REG_SZ, "")  '値は常に空文字とします。
        Case RegistrySetting.AllowContentAboveLock: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "AllowContentAboveLock", REG_DWORD, CBool(valueData) * -1)
        Case RegistrySetting.Rank: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "Rank", REG_DWORD, valueData)
        Case RegistrySetting.ShowInActionCenter: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "ShowInActionCenter", REG_DWORD, 0)  '値は常に0とします。
        Case RegistrySetting.AllowUrgentNotifications: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "AllowUrgentNotifications", REG_DWORD, CBool(valueData) * -1)

        '未定義
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select
 End Property

'***************************************************************************************************
'* 機能　　：レジストリ削除の簡単なプリセットを提供します
'---------------------------------------------------------------------------------------------------
'* 引数　  ：PresetID   列挙型（列挙体）Enum：RegistrySetting   ここで、削除したいIDを指定します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：この"hoge"に、意味はありません。何も使いません。インデックス不足エラー回避用です。
'***************************************************************************************************
Public Property Let PresetDelRegistry(PresetID As RegistrySetting)
    Select Case PresetID
        'HKEY_CLASSES_ROOT 用
        Case RegistrySetting.ShowInSettings: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "ShowInSettings", REG_Delete, "hoge")
        Case RegistrySetting.DisplayName: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "DisplayName", REG_Delete, "hoge")
        Case RegistrySetting.IconUri: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "IconUri", REG_Delete, "hoge")
        Case RegistrySetting.IconBackgroundColor: 最後に実行したレジストリ登録結果 = AttemptRegistryValue(DecisionAppUserModelID(True), "IconBackgroundColor", REG_Delete, "hoge")
        
        'HKEY_CURRENT_USER 用
        '--wpndatabase.db 連携必須--
        Case RegistrySetting.Enabled: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "Enabled", REG_Delete, "hoge")
        Case RegistrySetting.ShowBanner: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "ShowBanner", REG_Delete, "hoge")
        '--wpndatabase.db 連携不要--
        Case RegistrySetting.SoundFile: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "SoundFile", REG_Delete, "hoge")
        Case RegistrySetting.AllowContentAboveLock: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "AllowContentAboveLock", REG_Delete, "hoge")
        Case RegistrySetting.Rank: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "Rank", REG_Delete, "hoge")
        Case RegistrySetting.ShowInActionCenter: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "ShowInActionCenter", REG_Delete, "hoge")
        Case RegistrySetting.AllowUrgentNotifications: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(DecisionAppUserModelID(True), "AllowUrgentNotifications", REG_Delete, "hoge")

        '未定義
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select
End Property

'***************************************************************************************************
'* 機能　　：最後のレジストリ操作にて、ユーザーアカウント制御の反応を返します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：結果コード
'               -1：ユーザーアカウント制御にて許可しました。または、既に管理者権限状態で実行された
'               -2：ユーザーアカウント制御にて拒否しました
'               他：エラーコード
'***************************************************************************************************
Public Property Get LatestRegistryCode() As Long
    LatestRegistryCode = 最後に実行したレジストリ登録結果
End Property



'***************************************************************************************************
'                       ■■■ インスタンス生成時に自動で設定する内容 ■■■
'***************************************************************************************************
'* 機能　　：・トーストのヘッダー情報を設定します。不要ならコメントアウトしてください。デフォルトで以下にしています
'               Header_ID           このExcelのファイル名
'               Header_Title        このExcelのBook名
'               Header_Arguments    このExcelの保存フォルダパス
'
'            ・タイムゾーンを日本基準にします。
'**************************************************************************************************
Private Sub Class_Initialize()
    '必ず設定
    ElementBinding.template = "ToastGeneric"
    
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
    
    'subgroup要素に関する初期値を設定
    SubgroupTotal = -1
    SubgroupInnerTextsTotal = -1
    SubgroupInnerImagesTotal = -1
    SubgroupInnerTotal = -1
End Sub
