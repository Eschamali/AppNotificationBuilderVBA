VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppNotificationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   VBAから、`DispCallFunc × Vtable`を駆使したトースト通知機能を提供します
'
' MIT License
'
' Copyright (c) 2025 -2026 eschamali
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
'                               ■■■ 各種WindowsAPI宣言 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 基本機能
'---------------------------------------------------------------------------------------------------
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long      'アクティブExcelのプロセスID取得用
Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long    'カスタム音声

'---------------------------------------------------------------------------------------------------
' [ SECTION ] WinRT-API へ直接アクセスするための必要なAPI群
'---------------------------------------------------------------------------------------------------




'---------------------------------------------------------------------------------------------------
' [ SECTION ] レジストリ操作
'---------------------------------------------------------------------------------------------------



'---------------------------------------------------------------------------------------------------
' [ SECTION ] SQLite操作
'---------------------------------------------------------------------------------------------------










'***************************************************************************************************
'                           ■■■ ユーザー定義エラー番号文言 ■■■
'***************************************************************************************************
Private Const vbObjectErrorID_0001 As String = "設定したスケジュールは、現在のシステム日時に対して過去の日時に設定されてます。" & vbCrLf & "設定したスケジュール通知日時が、未来であることを確認してください。 "
Private Const vbObjectErrorID_0002 As String = "設定したパラメーターは、定義されていません。"



'***************************************************************************************************
'                         ■■■ "wpndatabase.db"に関する変数/定数一式 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] 最後のSQL実行結果
'---------------------------------------------------------------------------------------------------
Private LatestExecuteSQLiteNum      As Long

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 設定保存相対パス
'---------------------------------------------------------------------------------------------------
Private Const ToastNotification_SettingPathDB   As String = "\Microsoft\Windows\Notifications\wpndatabase.db"

'---------------------------------------------------------------------------------------------------
' [ SECTION ] WNSId
'---------------------------------------------------------------------------------------------------
'今のところ、システムでは2種類用意されてる
Public Enum Wpndatabase_WNSId
    wsCustom                '好きな文字列指定用
    wsSystem                'システム
    wsNonImmersivePackage   '非パッケージ
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] HandlerType
'---------------------------------------------------------------------------------------------------
'今のところ、3種類用意されてる
Public Enum Wpndatabase_HandlerType
    htDesktop       'Win32アプリ(非パッケージ)
    htImmersive     'パッケージアプリ(UWP)
    htSystem        'システムアプリ
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] SettingKey 一覧
'---------------------------------------------------------------------------------------------------
'今のところ、28種類用意されてる
'※ただし、一度でもアプリ側が対応するパラメーターを使用しないと出てこない仕様のため、まだまだあるかも…?
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
Private 最後に実行したレジストリ登録結果    As Long

'登録型一覧(ひとまず、3 + 1 用意)
Enum RegistryType
    REG_Delete = -1 '削除用
    REG_DWORD
    REG_SZ
    REG_EXPAND_SZ
End Enum

'通知関連のレジストリパスIDを用意
Enum ToastRegistryPathMode
    NotificationsSettings
    ClassesAppUserModelId
End Enum

'レジストリで設定できるパラメーターを用意
Enum RegistrySetting
    'コンピューター\HKEY_CURRENT_USER\Software\Classes\AppUserModelId
    ShowInSettings
    displayName
    iconUri
    IconBackgroundColor
    CustomActivator
    
    'コンピューター\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings
    '--wpndatabase.db 連携必須--
    enabled
    ShowBanner
    '--wpndatabase.db 連携不要--
    SoundFile
    AllowContentAboveLock
    Rank
    ShowInActionCenter
    AllowUrgentNotifications
End Enum



'***************************************************************************************************
'                           ■■■ 各種トースト通知関連変数/定数 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] このClassファイル構築時のプロセスID
'---------------------------------------------------------------------------------------------------
Private ThisExcelPID    As Long

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 既定値系
'---------------------------------------------------------------------------------------------------
Private Const AppUserModelID_Excel      As String = "Microsoft.Office.EXCEL.EXE.15" 'DeskTopアプリ：Excel
Private Const ToastTemplateName         As String = "ToastGeneric"                  '"ToastGeneric" に設定する必要があります
Private Const ToastSchedule             As String = "ExcelSchedule"                 'スケジュール通知として発行する際のID
Private Const パラメーター区切り文字    As String = "|"                             'Group用

'---------------------------------------------------------------------------------------------------
' [ SECTION ] データバインディング - 最上位テキスト要素
'---------------------------------------------------------------------------------------------------
Private Const DataBindingTopText_Title     As String = "{TopTextTitle}"
Private Const DataBindingTopText_Body      As String = "{TopTextBody}"
Private Const DataBindingTopText_Attribute As String = "{TopTextAttribution}"

'---------------------------------------------------------------------------------------------------
' [ SECTION ] データバインディング - Progress要素
'---------------------------------------------------------------------------------------------------
Private Const DataBindingProgress_Title         As String = "{ProgressTitle}"
Private Const DataBindingProgress_Status        As String = "{ProgressStatus}"
Private Const DataBindingProgress_Value         As String = "{ProgressValue}"
Private Const DataBindingProgress_ValueString   As String = "{ProgressValueString}"

'---------------------------------------------------------------------------------------------------
' [ SECTION ] リマインド用の`arguments`属性の予約語
'---------------------------------------------------------------------------------------------------
Const ActionArguments再通知 As String = "snooze"
Const ActionArguments解除   As String = "dismiss"

'---------------------------------------------------------------------------------------------------
' [ SECTION ] トーストが使用されるシナリオ (アラームやアラームなど)
'---------------------------------------------------------------------------------------------------
Public Enum EnumToastScenario
    tsDefault
    tsReminder
    tsAlarm
    tsIncomingCall
    tsUrgent
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 既定のサウンドの代わりに再生するメディア ファイル
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

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ボタンのスタイル
'---------------------------------------------------------------------------------------------------
Public Enum EnumHintButtonStyle
    hbsDefault
    hbsSuccess
    hbsCritical
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ユーザーが特定の操作を行った際に使用されるアクティブ化の種類
'---------------------------------------------------------------------------------------------------
Public Enum EnumToastActivationType
    taForeground = 1  'フォアグラウンド アプリが起動します。
    taBackground      '対応するバックグラウンド タスクがトリガーされ、ユーザーを中断することなくバックグラウンドでコードを実行できます。
    taProtocol        'プロトコルのアクティブ化を使用して別のアプリを起動します。
    taSystem          'Action要素でのリマインダーモード用です
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] フォント サイズ、太さ、不透明度を制御します
'---------------------------------------------------------------------------------------------------
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

'---------------------------------------------------------------------------------------------------
' [ SECTION ] テキストの水平方向の配置を制御
'---------------------------------------------------------------------------------------------------
Public Enum EnumAdaptiveTextAlign
    TextAlign_Default '既定値です。 配置がレンダラーによって自動的に決定されます。
    TextAlign_Auto    '配置が現在の言語とカルチャによって決定されます。
    TextAlign_Left    'テキストを左側に水平方向に配置します。
    TextAlign_Center  'テキストを中央に水平方向に配置します。
    TextAlign_Right   'テキストを右側に水平方向に配置します。
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 画像トリミングを指定
'---------------------------------------------------------------------------------------------------
Public Enum EnumAdaptiveImageCrop
    ImageCrop_Default   '既定値です。トリミングの動作がレンダラーによって決定されます。
    ImageCrop_None      '画像がトリミングされません。
    ImageCrop_Circle    '画像が円形にトリミングされます。
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 画像の水平方向の配置を指定
'---------------------------------------------------------------------------------------------------
Public Enum EnumAdaptiveImageAlign
    ImageAlign_Default  '既定値です。配置の動作がレンダラーによって決定されます。
    ImageAlign_Stretch  '利用可能な幅いっぱいに画像が拡大されます (また同時に、画像が配置される位置に応じて、利用可能な高さいっぱいに拡大されることもあります)。
    ImageAlign_Left     '画像を左側に配置し、ネイディブの解像度で表示します。
    ImageAlign_Center   '画像を中央に水平方向に配置し、ネイティブの解像度で表示します。
    ImageAlign_Right    '画像を右側に配置し、ネイディブの解像度で表示します。
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] コンテンツの垂直方向の配置を指定
'---------------------------------------------------------------------------------------------------
Public Enum AdaptiveSubgroupTextStacking
    TextStacking_Default    '既定値です。 レンダラーが既定の垂直方向の配置を自動的に選択します。
    TextStacking_Top        '上に合わせて垂直に配置されます。
    TextStacking_Center     '中央に合わせて垂直に配置されます。
    TextStacking_Bottom     '下に合わせて垂直に配置されます。
End Enum

'---------------------------------------------------------------------------------------------------
' [ SECTION ] 各種要素の上限
'---------------------------------------------------------------------------------------------------
Private Const MaxElementActionNum       As Integer = 5  'ボタンは最大5つ、入れれます
Private Const MaxElementInputNum        As Integer = 5  '入力要素は最大5つ、入れれます
Private Const MaxElementSelectionNum    As Integer = 5  '選択肢は最大5つ、入れれます



'***************************************************************************************************
'                   ■■■ トーストの挙動を決めるユーザー定義型宣言 ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] プロパティ/メソッド等で、決める設定群
'---------------------------------------------------------------------------------------------------
Private ToastConfig As ToastNotification
Private Type ToastNotification
    '[Windows.UI.Notifications]の挙動を決める設定値
    AppUserModelID              As String   'https://learn.microsoft.com/ja-jp/windows/win32/shell/appids
    Content                     As String   'https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-schema

    '[Windows.UI.Notifications]の主な設定値
    Tag                         As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.tag
    Group                       As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.group
    Schedule_ID                 As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtilenotification.id
    CollectionID                As String   'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationmanagerforuser.gettoastnotifierfortoastcollectionidasync
    ExpiresOnReboot             As Boolean  'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.expiresonreboot
    SuppressPopup               As Boolean  'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.suppresspopup
    Schedule_DeliveryTime       As Date     'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.-ctor
    ExpirationTime              As Date     'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.expirationtime OR https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.expirationtime
    Priority                    As Boolean  'https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationpriority
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] データバイディング関連の設定
'---------------------------------------------------------------------------------------------------
Private ToastDataBinding As NotificationData
Private Type NotificationData
    '最上位レベルのテキスト要素の Text
    TitleText As String
    ContentsText As String
    AttributionText As String

    'ProgressBar 要素
    ProgressTitle As String
    ProgressValueStringOverride As String
    ProgressStatus As String
    progressValue As Double

    '通知データが古いかどうかを判断
    SequenceNumber As Long
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] タイムゾーンの設定　※日付、時刻データを扱う場合は、UTC基準になる項目があるため
'---------------------------------------------------------------------------------------------------
Private TargetTimeZone As TimeZone
Private Type TimeZone
    Hour    As Integer  '-23~23
    Minute  As Integer  '0-59
End Type



'***************************************************************************************************
'                   ■■■ トーストの見た目を決めるユーザー定義型宣言 ■■■
'***************************************************************************************************
'下記に基づいて順番に、宣言します
'→https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/root-elements
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] Toast要素に付けれる属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast
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

'---------------------------------------------------------------------------------------------------
' [ SECTION ] binding要素に付けれる属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-binding
'---------------------------------------------------------------------------------------------------
Private ElementBinding As ToastBindingGeneric
Private Type ToastBindingGeneric
    baseUri     As String   'イメージ ソース属性の相対 URI と組み合わせた既定のベース URI。
    template    As String   'この値は "ToastGeneric" に設定する必要があります
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] binding親要素のimage子要素で使える属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-image
'---------------------------------------------------------------------------------------------------
Private ElementsImage As TopAdaptiveImage

'image要素で使える属性
Private Type AttributeImage
    AddImageQuery As Boolean    'イメージ URI にクエリ文字列を追加できるようにするかの設定値です
    Alt           As String     '支援技術のユーザー向けの画像の説明
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

'---------------------------------------------------------------------------------------------------
' [ SECTION ] binding親要素のtext子要素で使える値と、属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-text
'---------------------------------------------------------------------------------------------------
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
    Body        As ParameterText    '本文
    Attribution As ParameterText    '常に下部に表示するもの
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] subgroup要素に付けれる属性値と値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-subgroup
'---------------------------------------------------------------------------------------------------
Private ElementsSubgroup()      As AdaptiveSubgroup, SubgroupTotal As Long
Private SubgroupInnerTexts()    As SubAdaptiveText, SubgroupInnerTextsTotal As Long
Private SubgroupInnerImages()   As SubAdaptiveImage, SubgroupInnerImagesTotal As Long
Private SubgroupInnerOrder()    As SubgroupInnerOrder, SubgroupInnerTotal As Long

' --- グループ/サブグループの子として配置されるテキスト要素 ---
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

' --- グループ/サブグループの子として配置されるインライン画像 ---
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

' --- 配置順の管理 ---
Private Type SubgroupInnerOrder
    UseText           As Boolean  'テキストか画像かのフラグ値
    TargetElementNo   As Long     '対象の要素No
End Type

' --- subgroupに付けれる属性と値、順序データ ---
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

'---------------------------------------------------------------------------------------------------
' [ SECTION ] binding親要素のprogress子要素で使える属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-progress
'---------------------------------------------------------------------------------------------------
Private ElementProgress As AdaptiveProgressBar
Private Type AdaptiveProgressBar
    Title               As String   '省略可能なタイトル文字列。
    status              As String   '左側の進行状況バーの下に表示されるステータス文字列。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります。
    Value               As Double   'プログレスバーの値。 この値には、0.0 ~ 1.0 の浮動小数点数、または読み込みアニメーションを生成する値 "Indeterminate" を指定します。なおVBAの場合、負の場合で"Indeterminate"扱いとします。
    ValueStringOverride As String   '既定のパーセンテージ文字列の代わりに表示される省略可能な文字列。 これを指定しない場合は、"70%" などの文字が表示されます。
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] Audio要素に付けれる属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio
'---------------------------------------------------------------------------------------------------
Private ElementAudio As ToastAudio
Private Type ToastAudio
    loop   As Boolean   'ループフラグ
    silent As Boolean   '無音フラグ
    src    As String    '通知音名称
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] actions親要素のaction子要素で使える属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-action
'---------------------------------------------------------------------------------------------------
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

'挿入カウンター
Private ElementActionCount  As Integer

'---------------------------------------------------------------------------------------------------
' [ SECTION ] input親要素のselection子要素で使える属性値。
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-selection
'---------------------------------------------------------------------------------------------------
Private ElementSelection(1 To MaxElementSelectionNum) As ToastSelectionBox
Private Type ToastSelectionBox
    SelectionID As String   '識別子名
    ChoseName   As String   '選択肢に出す名称
End Type

'---------------------------------------------------------------------------------------------------
' [ SECTION ] actions親要素のinput子要素で使える属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-input
'---------------------------------------------------------------------------------------------------
Private ElementInput(1 To MaxElementInputNum) As IToastInput
Private Type IToastInput
    '属性一覧
    ID                 As String    '入力に関連付けられている ID
    Type               As String    'この属性には、"text"、"selection" のいずれかの値を指定できます。
    PlaceHolderContent As String    'テキスト入力用に表示されるプレースホルダー
    Title              As String    '入力のラベルとして表示されるテキスト。
    DefaultInput       As String    'デフォルトの入力値(公式にはないが、設定項目としてはあるもよう)

    '要素そのものの配列(Select要素が入ります)
    Value() As ToastSelectionBox
End Type

'挿入カウンター
Private ElementInputCount As Integer

'---------------------------------------------------------------------------------------------------
' [ SECTION ] header要素に付けれる属性値
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-header
'---------------------------------------------------------------------------------------------------
Private ElementHeader As ToastHeader
Private Type ToastHeader
    ID             As String    'このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。
    Title          As String    'ヘッダーのタイトルです。
    Arguments      As String    'ユーザーがこのヘッダーをクリックするとアプリに返されます。 null にすることはできません。
    ActivationType As String    'このヘッダーがクリックされた場合に使用するアクティブ化の種類。
End Type



'***************************************************************************************************
'                       ■■■ Enum → 文字列 変換プロシージャ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] アクティブ化の種類を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringToastActivationType(param As EnumToastActivationType) As String
    Select Case param
        Case EnumToastActivationType.taForeground:  EnumToStringToastActivationType = "foreground"
        Case EnumToastActivationType.taBackground:  EnumToStringToastActivationType = "background"
        Case EnumToastActivationType.taProtocol:    EnumToStringToastActivationType = "protocol"
        Case EnumToastActivationType.taSystem:      EnumToStringToastActivationType = "system"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] トーストが表すシナリオを文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringToastScenario(param As EnumToastScenario) As String
    Select Case param
        Case EnumToastScenario.tsAlarm:         EnumToStringToastScenario = "alarm"        '通知を永遠に表示する。action要素がないと効果は発動しないが、応答不可になっても設定問わず、必ず表示される。通知音はアラーム系(Alarm)か、ミュートのみ設定可能
        Case EnumToastScenario.tsReminder:      EnumToStringToastScenario = "reminder"     '通知を永遠に表示する。action要素がないと効果は発動しないが、任意の通知音に設定可能
        Case EnumToastScenario.tsIncomingCall:  EnumToStringToastScenario = "incomingCall" '通知を永遠に表示する。action要素がなくても効果発動。最後のボタン位置のみ、Windowsのテーマ色に基づく着色が施され、位置が必ず下側になる。通知音は呼び出し系(Call)か、ミュートのみ設定可能
        Case EnumToastScenario.tsUrgent:        EnumToStringToastScenario = "urgent"       '重要な通知として表示。応答不可にしても表示されますが、設定で表示させないようにすることが可能です。とりあえず常に表示させたい場合は、シナリオ：「アラーム」にしましょう
        Case Else   '通常の通知として表示。属性：duration  あるいは、アクセシビリティ→視覚効果→この時間が経過したら通知を破棄する  に基づく表示時間となります
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] メディア ファイルを文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringMsWinSoundEvent(param As EnumMsWinSoundEvent) As String
    Select Case param
        Case EnumMsWinSoundEvent.NotificationDefault:           EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Default"
        Case EnumMsWinSoundEvent.NotificationIM:                EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.IM":
        Case EnumMsWinSoundEvent.NotificationMail:              EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Mail"
        Case EnumMsWinSoundEvent.NotificationReminder:          EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Reminder"
        Case EnumMsWinSoundEvent.NotificationSMS:               EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.SMS"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm01:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm02:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm2"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm03:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm3"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm04:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm4"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm05:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm5"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm06:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm6"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm07:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm7"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm08:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm8"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm09:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm9"
        Case EnumMsWinSoundEvent.NotificationLoopingAlarm10:    EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Alarm10"
        Case EnumMsWinSoundEvent.NotificationLoopingCall01:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call"
        Case EnumMsWinSoundEvent.NotificationLoopingCall02:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call2"
        Case EnumMsWinSoundEvent.NotificationLoopingCall03:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call3"
        Case EnumMsWinSoundEvent.NotificationLoopingCall04:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call4"
        Case EnumMsWinSoundEvent.NotificationLoopingCall05:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call5"
        Case EnumMsWinSoundEvent.NotificationLoopingCall06:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call6"
        Case EnumMsWinSoundEvent.NotificationLoopingCall07:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call7"
        Case EnumMsWinSoundEvent.NotificationLoopingCall08:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call8"
        Case EnumMsWinSoundEvent.NotificationLoopingCall09:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call9"
        Case EnumMsWinSoundEvent.NotificationLoopingCall10:     EnumToStringMsWinSoundEvent = "ms-winsoundevent:Notification.Looping.Call10"
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] ボタンのスタイルを文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringHintButtonStyle(param As EnumHintButtonStyle) As String
    Select Case param
        Case EnumHintButtonStyle.hbsSuccess:    EnumToStringHintButtonStyle = "Success"     'ボタンが緑色です
        Case EnumHintButtonStyle.hbsCritical:   EnumToStringHintButtonStyle = "Critical"    'ボタンが赤です
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] コンテンツの垂直方向の配置を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringAdaptiveSubgroupTextStacking(param As AdaptiveSubgroupTextStacking) As String
    Select Case param
        Case AdaptiveSubgroupTextStacking.TextStacking_Top:     EnumToStringAdaptiveSubgroupTextStacking = "top"
        Case AdaptiveSubgroupTextStacking.TextStacking_Center:  EnumToStringAdaptiveSubgroupTextStacking = "center"
        Case AdaptiveSubgroupTextStacking.TextStacking_Bottom:  EnumToStringAdaptiveSubgroupTextStacking = "bottom"
        Case Else   '既定値
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] イメージの目的のトリミングを文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringAdaptiveImageCrop(param As EnumAdaptiveImageCrop) As String
    Select Case param
        Case EnumAdaptiveImageCrop.ImageCrop_Circle:    EnumToStringAdaptiveImageCrop = "circle"
        Case EnumAdaptiveImageCrop.ImageCrop_None:      EnumToStringAdaptiveImageCrop = "none"
        Case Else   '既定値
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] イメージの水平方向の配置を文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringAdaptiveImageAlign(param As EnumAdaptiveImageAlign) As String
    Select Case param
        Case EnumAdaptiveImageAlign.ImageAlign_Stretch:  EnumToStringAdaptiveImageAlign = "stretch"
        Case EnumAdaptiveImageAlign.ImageAlign_Left:     EnumToStringAdaptiveImageAlign = "left"
        Case EnumAdaptiveImageAlign.ImageAlign_Center:   EnumToStringAdaptiveImageAlign = "center"
        Case EnumAdaptiveImageAlign.ImageAlign_Right:    EnumToStringAdaptiveImageAlign = "right"
        Case Else   '既定値
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] テキスト スタイルは、フォント サイズ、太さ、不透明度を制御する文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringAdaptiveTextStyle(param As EnumAdaptiveTextStyle) As String
    Select Case param
        Case EnumAdaptiveTextStyle.TextStyle_Caption:           EnumToStringAdaptiveTextStyle = "caption"
        Case EnumAdaptiveTextStyle.TextStyle_CaptionSubtle:     EnumToStringAdaptiveTextStyle = "captionsubtle"
        Case EnumAdaptiveTextStyle.TextStyle_Body:              EnumToStringAdaptiveTextStyle = "body"
        Case EnumAdaptiveTextStyle.TextStyle_BodySubtle:        EnumToStringAdaptiveTextStyle = "bodysubtle"
        Case EnumAdaptiveTextStyle.TextStyle_Base:              EnumToStringAdaptiveTextStyle = "base"
        Case EnumAdaptiveTextStyle.TextStyle_BaseSubtle:        EnumToStringAdaptiveTextStyle = "basesubtle"
        Case EnumAdaptiveTextStyle.TextStyle_Subtitle:          EnumToStringAdaptiveTextStyle = "subtitle"
        Case EnumAdaptiveTextStyle.TextStyle_SubtitleSubtle:    EnumToStringAdaptiveTextStyle = "subtitlesubtle"
        Case EnumAdaptiveTextStyle.TextStyle_Title:             EnumToStringAdaptiveTextStyle = "title"
        Case EnumAdaptiveTextStyle.TextStyle_TitleSubtle:       EnumToStringAdaptiveTextStyle = "titlesubtle"
        Case EnumAdaptiveTextStyle.TextStyle_TitleNumeral:      EnumToStringAdaptiveTextStyle = "titlenumeral"
        Case EnumAdaptiveTextStyle.TextStyle_Subheader:         EnumToStringAdaptiveTextStyle = "subheader"
        Case EnumAdaptiveTextStyle.TextStyle_SubheaderSubtle:   EnumToStringAdaptiveTextStyle = "subheadersubtle"
        Case EnumAdaptiveTextStyle.TextStyle_SubheaderNumeral:  EnumToStringAdaptiveTextStyle = "subheadernumeral"
        Case EnumAdaptiveTextStyle.TextStyle_Header:            EnumToStringAdaptiveTextStyle = "header"
        Case EnumAdaptiveTextStyle.TextStyle_HeaderSubtle:      EnumToStringAdaptiveTextStyle = "headersubtle"
        Case EnumAdaptiveTextStyle.TextStyle_HeaderNumeral:     EnumToStringAdaptiveTextStyle = "headernumeral"
        Case Else   '既定値
    End Select
End Function

'---------------------------------------------------------------------------------------------------
' [ SECTION ] テキストの水平方向の配置を制御するを文字列で返します
'---------------------------------------------------------------------------------------------------
Private Function EnumToStringAdaptiveTextAlign(param As EnumAdaptiveTextStyle) As String
    Select Case param
        Case EnumAdaptiveTextAlign.TextAlign_Auto:      EnumToStringAdaptiveTextAlign = "auto"
        Case EnumAdaptiveTextAlign.TextAlign_Left:      EnumToStringAdaptiveTextAlign = "left"
        Case EnumAdaptiveTextAlign.TextAlign_Center:    EnumToStringAdaptiveTextAlign = "center"
        Case EnumAdaptiveTextAlign.TextAlign_Right:     EnumToStringAdaptiveTextAlign = "right"
        Case Else   '既定値
    End Select
End Function



'***************************************************************************************************
'       ■■■ Toastを構成するスキーマ情報を決める基となるメンバ変数へ設定するプロパティ ■■■
'***************************************************************************************************
'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、toast要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast
'---------------------------------------------------------------------------------------------------
'* 機能　　：launch、activationType属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：Schema          起動スキーマ。URLやファイルパスなど
'            ActivationType     このヘッダーがクリックされた場合に使用するアクティブ化の種類
'---------------------------------------------------------------------------------------------------
'* 注意事項：公式サイトでは、`activationType`の項目がありませんが、設定値としては存在してます
'***************************************************************************************************
Public Property Let XmlToast_Launch(Optional ActivationType As EnumToastActivationType = EnumToastActivationType.taProtocol, Schema As String)
    With ElementToast
        .Launch = Schema
        .ActivationType = EnumToStringToastActivationType(ActivationType)
    End With
End Property

'***************************************************************************************************
'* 機能　　：duration属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：UseLong
'               True ：long  として扱います。25秒間、表示されます。
'               False：short として扱います。省略も同等なので、ここでは、省略として扱います
'---------------------------------------------------------------------------------------------------
'* 注意事項：アクセシビリティ→視覚効果→この時間が経過したら通知を破棄する  で表示時間をオーバーライド可能です。(longの場合は、上回る場合にオーバーライド)
'***************************************************************************************************
Public Property Let XmlToast_Duration(UseLong As Boolean)
    If UseLong Then ElementToast.Duration = "long" Else ElementToast.Duration = vbNullString
End Property

'***************************************************************************************************
'* 機能　　：toast要素のdisplayTimestamp属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：DisplayTimestamp    カスタム タイムスタンプで既定のタイムスタンプをオーバーライドします
'---------------------------------------------------------------------------------------------------
'* 注意事項：※ISO 8601で指定する都合上、UTCを設定する引数を設けますが、JSTの場合は気にする必要はありません。
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/custom-timestamps-on-toasts
'***************************************************************************************************
Public Property Let XmlToast_DisplayTimestamp(Optional UTC_Hour As Integer = 9, Optional UTC_Minute As Integer = 0, DisplayTimestamp As Date)
    With WorksheetFunction
        If DisplayTimestamp Then ElementToast.DisplayTimestamp = .Text(DisplayTimestamp, "yyyy-mm-ddThh:mm:ss") & .Text(TargetTimeZone.Hour, "+00:;-00:") & .Text(TargetTimeZone.Minute, "00") Else ElementToast.DisplayTimestamp = vbNullString
    End With
End Property

'***************************************************************************************************
'* 機能　　：synalio属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID    シナリオID。便宜上、数値に応じてシナリオ名を設定します
'***************************************************************************************************
Public Property Let XmlToast_Scenario(ToastScenarioMode As EnumToastScenario)
    ElementToast.Synalio = EnumToStringToastScenario(ToastScenarioMode)
End Property

'***************************************************************************************************
'* 機能　　：useButtonStyle属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArguseButtonStyle
'               True ：スタイル付きボタンを使用します
'               False：スタイル付きボタンを使用しません。
'---------------------------------------------------------------------------------------------------
'* 注意事項：この機能は、Windows 11 以降のみ機能します。
'***************************************************************************************************
Public Property Let XmlToast_UseButtonStyle(UseButtonStyle As Boolean)
    ElementToast.UseButtonStyle = UseButtonStyle
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、binding要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-binding
'---------------------------------------------------------------------------------------------------
'* 機能　　binding要素の各属性へ値をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgBaseImageURI    イメージ ソース属性の相対 URI と組み合わせた既定のベース URI。
'***************************************************************************************************
Public Property Let XmlBinding_baseUri(baseUri As String)
    ElementBinding.baseUri = baseUri
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、image要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-image
'---------------------------------------------------------------------------------------------------
'* 機能　　：インラインに設定する画像のパスと、丸いロゴフラグの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：InlineImageSrc    イメージ ソースの URI。ファイルパスや、http(s)画像など
'            InlineCircle
'               True ：画像は円にトリミングされます。
'               False：画像はトリミングされず、そのまんまのサイズとして表示されます。

'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Alt                画像の説明を記載します。
'***************************************************************************************************
Public Property Let XmlImage_Inline(Optional InlineCircle As Boolean, Optional Flag_addImageQuery As Boolean, Optional Alt As String, InlineImageSrc As String)
    With ElementsImage.Inline
        .src = InlineImageSrc
        .AddImageQuery = Flag_addImageQuery
        .Alt = Alt

        If InlineCircle Then .HintCrop = "circle" Else .HintCrop = vbNullString
    End With
End Property

'***************************************************************************************************
'* 機能　　：ヒーロー画像のパスの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：HeroImageSrc       イメージ ソースの URI。ファイルパスや、http(s)画像など
'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Alt                画像の説明を記載します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：ヒーロー画像の丸み加工はできません
'***************************************************************************************************
Public Property Let XmlImage_Hero(Optional Flag_addImageQuery As Boolean, Optional Alt As String, HeroImageSrc As String)
    With ElementsImage.Hero
        .src = HeroImageSrc
        .AddImageQuery = Flag_addImageQuery
        .Alt = Alt

        .Placement = "hero"
    End With
End Property

'***************************************************************************************************
'* 機能　　：、AppLogo(appLogoOverride)に設定する画像のパスと、丸いロゴフラグの設定を行う
'---------------------------------------------------------------------------------------------------
'* 引数　　：AppLogoSrc    イメージ ソースの URI。ファイルパスや、http(s)画像など
'            LogoCircle
'               True ：画像は円にトリミングされます。
'               False：画像はトリミングされず、正方形として表示されます。

'            Flag_addImageQuery     srcのURIに、クエリ文字列を追加できるようにするには、"true"を指定します。
'            Alt                画像の説明を記載します。
'***************************************************************************************************
Public Property Let XmlImage_AppLogo(Optional LogoCircle As Boolean, Optional Flag_addImageQuery As Boolean, Optional Alt As String, AppLogoSrc As String)
    With ElementsImage.AppLogo
        .src = AppLogoSrc
        .AddImageQuery = Flag_addImageQuery
        .Alt = Alt

        .Placement = "appLogoOverride"

        If LogoCircle Then .HintCrop = "circle" Else .HintCrop = vbNullString
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、text 要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-text
'---------------------------------------------------------------------------------------------------
'* 機能　　：タイトル部分を指定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Title                          タイトル部分の通知内容
'            HintCallScenarioCenterAlign    テキストを中央に配置するフラグ ※Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：通知上部のタイトルに設定する内容です
'* 注意事項：・最大、2行まで表示できます
'            ・クリアは、`vbnullstring`で可能です
'***************************************************************************************************
Public Property Let XmlText_Title(Optional HintCallScenarioCenterAlign As Boolean, Title As String)
    With ElementsText.Title
        .Value = Title
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign
    End With
End Property

'***************************************************************************************************
'* 機能　　：Body部分を指定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Body                           通知内容
'            HintCallScenarioCenterAlign    テキストを中央に配置するフラグ ※Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます
'---------------------------------------------------------------------------------------------------
'* 機能説明：通知の中ほどに設定する内容です
'* 注意事項：・最大、4行まで表示できます
'            ・クリアは、`vbnullstring`で可能です
'***************************************************************************************************
Public Property Let XmlText_Body(Optional HintCallScenarioCenterAlign As Boolean, Body As String)
    'タイトルが空欄の場合は、そこに半角スペースを入れて、このBodyがタイトル化しないようにします
    If ElementsText.Title.Value = "" Then ElementsText.Title.Value = " "

    'Body部分を指定
    With ElementsText.Body
        .Value = Body
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign
    End With
End Property

'***************************************************************************************************
'* 機能　　`attribution`にあたる内容を指定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Attribution                    通知内容
'            HintCallScenarioCenterAlign    テキストを中央に配置するフラグ(Synalio属性値が"incomingCall"のToast要素のみ効果発動。それ以外の場合は無視されます。
'---------------------------------------------------------------------------------------------------
'* 機能説明：通知の下部に設定するテキストを設定します。
'* 注意事項：・最大、2行まで表示できます
'            ・クリアは、`vbnullstring`で可能です
'***************************************************************************************************
Public Property Let XmlText_Attribution(Optional HintCallScenarioCenterAlign As Boolean, Attribution As String)
    With ElementsText.Attribution
        .Value = Attribution
        .Attribute.HintCallScenarioCenterAlign = HintCallScenarioCenterAlign

        .Attribute.Placement = "attribution"
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、Audio 要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio
'---------------------------------------------------------------------------------------------------
'* 機能　　：Audio要素の属性をセット
'---------------------------------------------------------------------------------------------------
'* 引数　　：SrcID          通知音ID。便宜上、数値に応じて通知音名を設定します
'               1-25        Windows側で定められた通知音名を番号で指定
'               上記以外    ミュートにします

'            ArgLoop
'               True  ：トーストが表示されている限りサウンドが繰り返される場合に設定
'               False ：1回だけ再生する場合に設定
'***************************************************************************************************
Public Property Let XmlAudio(Optional ArgLoop As Boolean, SrcID As EnumMsWinSoundEvent)
    With ElementAudio
        'audio要素のLoop属性を設定
        .loop = ArgLoop

        'ArgSrcID に応じて、ミュート判定を決める
        If 0 < SrcID And SrcID <= DefNoticeSoundNum Then
            '存在する番号ならミュートしない
            .silent = False

            '音源を設定
            '※toast要素のsynalio属性値によっては、この設定が効かないことがあります
            .src = EnumToStringMsWinSoundEvent(SrcID)
        Else
            '存在しない番号ならミュート
            .silent = True
        End If
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、actions  要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-action
'---------------------------------------------------------------------------------------------------
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
'* 注意事項：・公式Document上には、type属性があるようですがここでは一旦無視します
'            ・`Content`に`vbnullstring`を適用すると、そのIndexでは未使用にできます
'            ・文字列のパラメータークリアにも`vbnullstring`で対応可能です
'***************************************************************************************************
Public Property Let XmlAction(Content As String, Arguments As String, Optional ActivationType As EnumToastActivationType = EnumToastActivationType.taProtocol, Optional PendingUpdate As Boolean, Optional ContextMenu As Boolean, Optional IconPath As String, Optional HintInputId As String, Optional HintButtonStyle As EnumHintButtonStyle, Optional HintToolTip As String, Index As Integer)
    With ElementAction(Index)
        '1. 挿入カウントを行う
        '有効な入力値かつ、Index先が`null`
        If StrPtr(Content) > 0 And StrPtr(.Content) = 0 Then
            'カウントアップ
            ElementActionCount = ElementActionCount + 1

        '`null`入力かつ、Index先が有効な入力値
        ElseIf StrPtr(Content) = 0 And StrPtr(.Content) > 0 Then
            'カウントダウン
            ElementActionCount = ElementActionCount - 1
        End If

        '2. 各種パラメーター適用
        .Content = Content
        .Arguments = Arguments
        .ActivationType = EnumToStringToastActivationType(ActivationType)
        '.Type = ArgType

        If PendingUpdate Then .afterActivationBehavior = "pendingUpdate" Else .afterActivationBehavior = vbNullString
        If ContextMenu Then .Placement = "contextMenu" Else .Placement = vbNullString

        .ImageUri = IconPath
        .HintInputId = HintInputId
        .HintButtonStyle = EnumToStringHintButtonStyle(HintButtonStyle)
        .HintToolTip = HintToolTip
    End With
End Property

'***************************************************************************************************
'* 機能　　：input要素の各属性にセットしたものをaction要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ID                     入力に関連付けられている ID
'            ChoseFlag              Trueで"SelectionBox"、Falseで、"TextBox"扱いとします
'            PlaceHolderContent  テキスト入力用に表示されるプレースホルダー。
'            title               入力のラベルとして表示されるテキスト
'            DefaultInput        デフォルトの入力値
'            Index                  追加位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：・「ChoseFlag = True」でこのプロパティを使用すると「XmlSelection」で設定したプロパティをリセットします
'            ・`ID`に`vbnullstring`を適用すると、そのIndexでは未使用にできます
'            ・文字列のパラメータークリアにも`vbnullstring`で対応可能です
'***************************************************************************************************
Public Property Let XmlInput(ByVal ID As String, Optional ByVal ChoseFlag As Boolean, Optional ByVal PlaceHolderContent As String, Optional ByVal Title As String, Optional ByVal DefaultInput As String, ByVal Index As Integer)
    With ElementInput(Index)
        '1. 挿入カウントを行う
        '有効な入力値かつ、Index先が`null`
        If StrPtr(ID) > 0 And StrPtr(.ID) = 0 Then
            'カウントアップ
            ElementInputCount = ElementInputCount + 1

        '`null`入力かつ、Index先が有効な入力値
        ElseIf StrPtr(ID) = 0 And StrPtr(.ID) > 0 Then
            'カウントダウン
            ElementInputCount = ElementInputCount - 1
        End If

        '2. 各種パラメーター適用
        .ID = ID

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

        .PlaceHolderContent = PlaceHolderContent
        .Title = Title
        .DefaultInput = DefaultInput
    End With
End Property

'***************************************************************************************************
'* 機能　　：input要素のtype属性が"selection" の時に挿入する"selection"要素を定義します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgSelectionID    識別子。※`ToastNotification.Activated`を利用する場合、これが`UserInput`の値として返されます。
'                              なお、数値のみの指定にすると、リマインダー用途としても使用可能です(単位：分)
'
'            ArgChoseName      選択肢名称
'            Index             挿入位置
'---------------------------------------------------------------------------------------------------
'* 注意事項：前述の`XmlInput`にて、`ChoseFlag:= True`を使用する前に、このプロパティで設定してください
'***************************************************************************************************
Public Property Let XmlSelection(SelectionID As String, ChoseName As String, Index As Integer)
    With ElementSelection(Index)
        .SelectionID = SelectionID
        .ChoseName = ChoseName
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、group 要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-group
'---------------------------------------------------------------------------------------------------
'* 機能　　：subgroup要素に属性と、事前に定義したtext,image要素をそれぞれ追加します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Weight         別のサブグループを基準として太さを指定することで、このサブグループの列の幅を制御します。
'            TextStacking   このサブグループのコンテンツの垂直方向の配置を制御します。
'---------------------------------------------------------------------------------------------------
'* 注意事項：このプロパティを使用すると以前、「AddAdaptiveSubgroupImage」、「AddAdaptiveSubgroupText」で呼び出して設定した配列、
'             SubgroupInnerImages,SubgroupInnerTexts の中身をリセットします。
'***************************************************************************************************
Public Sub AddAdaptiveSubgroup(Optional Weight As Integer, Optional TextStacking As AdaptiveSubgroupTextStacking)
    '1. 加算処理(0基準)
    SubgroupTotal = SubgroupTotal + 1

    '2. 配列を増やします
    ReDim Preserve ElementsSubgroup(SubgroupTotal)

    '3. パラメーターを設定します
    With ElementsSubgroup(SubgroupTotal)
        '各属性
        .HintWeight = Weight
        .HintTextStacking = EnumToStringAdaptiveSubgroupTextStacking(TextStacking)

        '順序データ
        .QueueOrder = SubgroupInnerOrder

        '各値
        .QueueText = SubgroupInnerTexts
        .QueueImage = SubgroupInnerImages
    End With

    '4. 以前定義したsubgroup要素内のtext,image要素をリセットする
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
'* 機能　　：subgroup要素に設定した中身の要素を全てクリーンします
'***************************************************************************************************
Public Sub CleanAdaptiveSubgroup()
    'カウントリセット
    SubgroupTotal = -1

    '配列リセット
    ReDim ElementsSubgroup(SubgroupTotal + 1)
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
Public Property Let AddAdaptiveSubgroupImage(Optional Crop As EnumAdaptiveImageCrop, Optional RemoveMargin As Boolean, Optional Align As EnumAdaptiveImageAlign, Optional AltText As String, Optional AddImageQuery As Boolean, ImageSrc As String)
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
        .src = ImageSrc
        .HintCrop = EnumToStringAdaptiveImageCrop(Crop)
        .HintAlign = EnumToStringAdaptiveImageAlign(Align)
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
'               Language XML ペイロードの対象ロケールです。
'               TextContents 表示テキスト
'---------------------------------------------------------------------------------------------------
'* 機能説明：アダプティブを使用したテキストをレンダリング出来ます。
'***************************************************************************************************
Public Property Let AddAdaptiveSubgroupText(Optional Style As EnumAdaptiveTextStyle, Optional Wrap As Boolean, Optional MaxLines As Integer, Optional MinLines As Integer, Optional Align As EnumAdaptiveTextAlign, Optional Language As String, TextContents As String)
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
        .HintStyle = EnumToStringAdaptiveTextStyle(Style)
        .HintWrap = Wrap
        .HintMaxLines = MaxLines
        .HintMinLines = MinLines
        .HintAlign = EnumToStringAdaptiveTextAlign(Align)
        .Language = Language

        'テキスト値
        .Value = TextContents
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、progress 要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-progress
'---------------------------------------------------------------------------------------------------
'* 機能　　：binding親要素のProgress子要素に必要な属性を一括設定します
'---------------------------------------------------------------------------------------------------
'* 引数　　：Title                  タイトルの文字列 (オプション) を取得または設定します。 データ バインディングをサポートしています。
'            Arg_Value              進行状況バーの値を取得または設定します。 データ バインディングをサポートしています。 既定値は 0 です。 0.0 ～ 1.0 の double 型で、ゲージ式。負の値にすると、ドットアニメーションの不確定式になります。
'            ValueStringOverride    割合を示す既定の文字列に代わって表示される文字列 (オプション) を取得または設定します。 これを指定しない場合は、"70%" などの文字が表示されます。
'            Status                 進行状況バーの下の左側に表示されるステータス文字列 (必須) を取得または設定します。 この文字列は、"ダウンロード中..." や "インストール中..." などのように、操作の状態を反映する必要があります。
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`Status`に`vbnullstring`をセットすることで、この要素を削除できます
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-progress-bar
'***************************************************************************************************
Public Property Let SetAdaptiveProgressBar(Optional Arg_Value As Double, Optional Title As String, Optional ValueStringOverride As String, status As String)
    With ElementProgress
        .Title = Title
        .status = status
        .Value = Arg_Value
        .ValueStringOverride = ValueStringOverride
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] XMLを構成するプロパティのうち、header 要素を指します
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-header
'---------------------------------------------------------------------------------------------------
'* 機能　　：header要素の各属性にセットしたものをtoast要素に挿入します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ArgID            ：このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。(省略時：このExcelのファイルフルパス)
'            ArgArguments     ：起動スキーマ。URLやファイルパスなど(省略時：このExcelの保存フォルダパス)
'            ArgActivationType：このヘッダーがクリックされた場合に使用するアクティブ化の種類
'            ArgTitle         ：ヘッダーのタイトルです。(推奨：このExcelのBook名　ThisWorkbook.Name)
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ArgTitle を vbnullstring で渡すとこの header要素自体を除去できます。
'* 注意事項：Microsoft の思想設計上、ActivationType は、Protocol が推奨されてます。
'***************************************************************************************************
Public Property Let XmlHeader(Optional ID As String, Optional Arguments As String, Optional ActivationType As EnumToastActivationType = EnumToastActivationType.taProtocol, ByVal Title As String)
    With ElementHeader
        '省略時：このExcelのファイルフルパス
        If StrPtr(ID) Then .ID = ID Else .ID = ThisWorkbook.FullName

        '省略時：このExcelの保存フォルダパス
        If StrPtr(Arguments) Then .Arguments = Arguments Else .Arguments = ThisWorkbook.Path

        '引数の中身そのまんま
        .ActivationType = EnumToStringToastActivationType(ActivationType)
        .Title = Title
    End With
End Property

'---------------------------------------------------------------------------------------------------
' [ SECTION ] トースト スキーマ生成
' [ document] https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/schema-root
'---------------------------------------------------------------------------------------------------
'* 機能　　 ：上記のLETプロパティで設定した「トースト スキーマ」のXMLを取り出す
'---------------------------------------------------------------------------------------------------
'* 返り値　 ：トースト スキーマのxmlContentsが返る
'* 引数　　 ：DataBindingMode   True ：トーストの更新ができるような値の指定を行います(例：{UpdateParameter})
'*                              False：値を直接記入します。スケジュール通知や、リマインド通知といった更新処理が使えないシチュエーションに指定します
'---------------------------------------------------------------------------------------------------
'* 注意事項 ：「Microsoft XML. v6.0」の参照設定が必要です
'***************************************************************************************************
Public Property Get ToastNotificationContent(DataBindingMode As Boolean) As String
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
        If StrPtr(ElementBinding.baseUri) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "baseUri", "")).NodeValue = ElementBinding.baseUri
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
                If StrPtr(ElementsImage.AppLogo.Alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.AppLogo.Alt

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
                If StrPtr(ElementsImage.Inline.Alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.Inline.Alt

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
                If StrPtr(ElementsImage.Hero.Alt) Then .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "alt", "")).NodeValue = ElementsImage.Hero.Alt

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
                If DataBindingMode Then .Text = DataBindingTopText_Body Else .Text = ElementsText.Body.Value

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
    If StrPtr(ElementProgress.status) And DataBindingMode Then
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
    'action要素、input要素、いずれかが設定されてあったら、ノード：actions  を作成してルートノードの子ノードとして追加します。
    If ElementActionCount > 0 Or ElementInputCount > 0 Then
        Set Node03__actions = RootNode_toast.appendChild(ToastContents.createNode(NODE_ELEMENT, NameNode03__actions, ""))


        '-------------------------------------------------------input要素を設定-------------------------------------------------------
        '各属性ノードを作成して値を設定し、「input」ノードの属性に追加します。
        '5つまで、繰り返す
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
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "id", "")).NodeValue = ElementHeader.ID
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "title", "")).NodeValue = ElementHeader.Title
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "arguments", "")).NodeValue = ElementHeader.Arguments
            .Attributes.setNamedItem(ToastContents.createNode(NODE_ATTRIBUTE, "activationType", "")).NodeValue = ElementHeader.ActivationType
        End With
    End If


    '上記の処理で生成されたxmlコンテンツを返します(末尾に改行が入らないように抽出文字数を調整します)
    ToastNotificationContent = Mid(ToastContents.xml, 1, Len(ToastContents.xml) - Len(vbCrLf))
End Property



'***************************************************************************************************
'                                  ■■■ Toast全般の設定 ■■■
'***************************************************************************************************
'* 機能    ：この通知をどのアプリで出すかを設定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：設定すると、ExcelのAppUserModelIDよりも優先して設定されます。batファイル等、VBAと直接関係なくなり、アイコンが噛み合わない時に使います。
'* 注意事項：存在しない(未インストール)AppID、無効な文字列を指定すると、Toastが発行されないのでご注意ください。
'---------------------------------------------------------------------------------------------------
Public Property Let SetAppUserModelID(appId As String)
    ToastConfig.AppUserModelID = appId
End Property

'***************************************************************************************************
'* 機能    ：タイムゾーンを変更します。デフォルトは、日本基準のため日本在宅者は基本、使用する必要はありません
'---------------------------------------------------------------------------------------------------
'* 注意事項：Class_Initialize　で、日本時間にセットされます
'---------------------------------------------------------------------------------------------------
Public Sub SetTimeZone(ArgUTC_Hour As Integer, ArgUTC_Minute As Integer)
    With TargetTimeZone
        .Hour = ArgUTC_Hour
        .Minute = ArgUTC_Minute
    End With
End Sub

'***************************************************************************************************
'* 機能　　：トーストからマクロ実行するためのブック情報を提供します
'---------------------------------------------------------------------------------------------------
'* 返り値　：ブック名|プロセスID　という形式
'---------------------------------------------------------------------------------------------------
'* 機能説明：ToastNotification.Group プロパティの設定に使い、使用中のブックをユニークに特定します
'            これは、コールバックでの特定に使います
'***************************************************************************************************
Private Function GetUniqueBookInfo() As String
    GetUniqueBookInfo = ThisWorkbook.Name & パラメーター区切り文字 & ThisExcelPID
End Function



'***************************************************************************************************
'                                    ■■■ 変換用 ■■■
'***************************************************************************************************
'* 機能    ：WinRT 直呼び出し用の設定構造体へ変換します
'---------------------------------------------------------------------------------------------------
'* 機能説明：基本は、As Is ですが日時系は、UTC基準になるため、一部変換が入ります
'***************************************************************************************************
Private Function ParamsConvertForWinRT(ToastSetting As ToastNotification) As WinRT_ToastConfig
    With ParamsConvertForWinRT
        .AppUserModelID = ToastSetting.AppUserModelID
        .XmlTemplate = ToastSetting.Content
        .Tag = ToastSetting.Tag
        .Group = ToastSetting.Group
        .Schedule_ID = ToastSetting.Schedule_ID
        .CollectionID = ToastSetting.CollectionID
        .ExpiresOnReboot = ToastSetting.ExpiresOnReboot
        .SuppressPopup = ToastSetting.SuppressPopup

        .Schedule_DeliveryTimeLocal = ToastSetting.Schedule_DeliveryTime
        .Schedule_DeliveryTime = DateAdd("h", TargetTimeZone.Hour * -1, ToastSetting.Schedule_DeliveryTime)
        .Schedule_DeliveryTime = DateAdd("n", TargetTimeZone.Minute * -1, .Schedule_DeliveryTime)

        .ExpirationTime = DateAdd("h", TargetTimeZone.Hour * -1, ToastSetting.ExpirationTime)
        .ExpirationTime = DateAdd("n", TargetTimeZone.Minute * -1, .ExpirationTime)
    End With
End Function

'***************************************************************************************************
'* 機能    ：WinRT 直呼び出し用のデータバインディング構造体へ変換します
'***************************************************************************************************
Private Function DataBindingConvertForWinRT() As WinRT_DataBinding
    With DataBindingConvertForWinRT
        '最上位のテキスト要素部分
        .TitleText = ElementsText.Title.Value
        .ContentsText = ElementsText.Body.Value
        .AttributionText = ElementsText.Attribution.Value

        'プログレスバー
        '※必須項目が入ってたら、他の項目も適用させます
        If StrPtr(ElementProgress.status) Then
            .ProgressTitle = ElementProgress.Title
            .ProgressStatus = ElementProgress.status
            .progressValue = ElementProgress.Value
            .ProgressValueStringOverride = ElementProgress.ValueStringOverride
        End If
    End With
End Function



'***************************************************************************************************
'                                   ■■■ 各種便利機能 ■■■
'***************************************************************************************************
'* 機能　　：Windows API:mciSendString を利用して、疑似的にトーストの通知音を自分好みの音源ファイルに指定して鳴らします
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
'* 処理名　：PresetReminder
'* 機能　　：リマインダー機能をプリセットします。システム側で用意されているパラメーターをセットすることで、システムの再通知と無視アクションを利用するアラーム通知を作成できます。
'---------------------------------------------------------------------------------------------------
'* 引数　　：RemindMinute1-5    再通知する時間を分で指定
'            Message            タイトル

'* 注意事項：このプロシージャを呼び出す前に定義した下記の要素は、上書きされますのでご注意ください
'           ・1,2つ目のaction要素
'           ・1つ目のinput要素

'* URL 　　：https://learn.microsoft.com/ja-jp/windows/apps/develop/notifications/app-notifications/app-notifications-content?tabs=appsdk#snoozedismiss
'***************************************************************************************************
Public Sub PresetReminder(RemindMinute1 As Long, Optional RemindMinute2 As Long, Optional RemindMinute3 As Long, Optional RemindMinute4 As Long, Optional RemindMinute5 As Long, Optional Message As String)
    '紐付け用識別子(なんでもOK)
    Const ReminderID As String = "snoozeTime"


    '----------------------------1.各Selection要素を用意----------------------------
    '選択肢を用意
    XmlSelection(CStr(RemindMinute1), RemindMinute1 & " 分後") = 1
    XmlSelection(CStr(RemindMinute2), RemindMinute2 & " 分後") = 2
    XmlSelection(CStr(RemindMinute3), RemindMinute3 & " 分後") = 3
    XmlSelection(CStr(RemindMinute4), RemindMinute4 & " 分後") = 4
    XmlSelection(CStr(RemindMinute5), RemindMinute5 & " 分後") = 5


    '----------------------------2.1 で用意したSelection要素をinput要素を作成する際のIndex値を指定し、Type=True でset----------------------------
    '選択肢を1つ目のinput要素に格納
    XmlInput(ReminderID, True, , Message, RemindMinute1) = 1


    '----------------------------3.1,2つ目のaction要素に、それぞれプリセット----------------------------
    XmlAction("", ActionArguments再通知, taSystem, , , , ReminderID) = 1
    XmlAction("", ActionArguments解除, taSystem) = 2
End Sub



'***************************************************************************************************
'               ■■■ Toast通知の機能を呼び出すメインディッシュラッパーメソッド ■■■
'***************************************************************************************************
'* 機能　　：引数に渡された値で、単純なトースト通知を表示します。指定日時に通知するスケジュール機能も対応します
'---------------------------------------------------------------------------------------------------
'* 引数　　：ToastTag           ToastTagを指定します。複数通知を提供する際に必要です
'            CollectionID       トーストコレクションIDがあるなら、それを指定
'            ScheduleDate       指定日時に通知したい場合は、「Date型」として指定することでスケジュール通知として登録されます。省略時、即時通知とします
'            ExpirationDate     アクションセンターから自動削除される日時を指定。最大30sの遅延バグが発生します。 ※上限は3日後までです
'            Suppress           True で、ポップアップ メッセージを表示せずに、トースト通知がアクション センターに通知されます。
'            ExpiresOnReboot    True で、再起動すると同時に、トースト通知が消えます
'**************************************************************************************************
Public Sub RunToastNotifierShow(ToastTag As String, Optional CollectionID As String, Optional ScheduleDate As Date, Optional ExpirationDate As Date, Optional Suppress As Boolean, Optional ExpiresOnReboot As Boolean)
    '--- 1. データバインディングを使用するか判定する ---
    Dim UseDataBinding As Boolean, i As Long
    If ScheduleDate > 0 Then
        'スケジュール設定：あり
        'データバインディングを、無効化する
        UseDataBinding = False

        '過去になっている場合、VBAエラーを起こして、処理を停止させます
        If Now() >= ScheduleDate Then Err.Raise Number:=vbObjectError + 1, Description:=vbObjectErrorID_0001
    Else
        'スケジュール設定：なし
        'データバインディングを、有効化する
        UseDataBinding = True

        'ただし、action要素に特定の組み合わせを検知したら、データバインディングを、無効化する
        For i = 1 To MaxElementActionNum
            If ElementAction(i).ActivationType = EnumToStringToastActivationType(taSystem) And ElementAction(i).Arguments = ActionArguments再通知 Then UseDataBinding = False: Exit For
        Next
    End If

    '--- 2. 設定値をまとめる ---
    With ToastConfig
        .Content = ToastNotificationContent(UseDataBinding)

        .Tag = ToastTag
        .Group = GetUniqueBookInfo
        .CollectionID = CollectionID

        .ExpiresOnReboot = ExpiresOnReboot
        .SuppressPopup = Suppress

        .Schedule_DeliveryTime = ScheduleDate
        .ExpirationTime = ExpirationDate
    End With

    '--- 3. データバインディングを設定しつつ、WinRT 直呼び出しで通知を表示します ---
    Dim WinRTConfig As WinRT_ToastConfig
    WinRTConfig = ParamsConvertForWinRT(ToastConfig)
    WinRT_ShowToastNotification WinRTConfig, DataBindingConvertForWinRT()
End Sub

'***************************************************************************************************
'* 機能　　：事前に「XmlText系」と「SetAdaptiveProgressBar」で変更した値を、表示中のトーストタグに即反映させ、更新させます
'---------------------------------------------------------------------------------------------------
'* 返り値　：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.notificationupdateresult
'            0(Succeeded)               通知が更新されました。
'            1(Failed)                  通知の更新に失敗しました。
'            2(NotificationNotFound)    指定した通知が見つかりませんでした。
'            ※上記以外                  予期せぬエラー

'* 引数　　：ToastTag                   ToastTagを指定します。以後の進捗状況を更新する際に必要です
'            CollectionID               `CollectionID`がある場合は指定します
'---------------------------------------------------------------------------------------------------
'* 機能説明：既にトーストの進行状況バーを表示しているとき、その時に指定したタグを指定することで、更新が可能です
'***************************************************************************************************
Public Function RunToastNotifierUpdate(ToastTag As String, Optional CollectionID As String) As Long
    '--- 1. 必要な設定値をまとめる ---
    With ToastConfig
        .Tag = ToastTag
        .Group = GetUniqueBookInfo
        .CollectionID = CollectionID
    End With

    '--- 2. WinRT 直呼び出しで通知を更新します ---
    RunToastNotifierUpdate = WinRT_UpdateToastNotification(ParamsConvertForWinRT(ToastConfig), DataBindingConvertForWinRT())
End Function

'**************************************************************************************************
'* 機能　　：タグから、トースト通知を削除します
'---------------------------------------------------------------------------------------------------
'* 引数　　：CollectionID       コレクション通知IDの方を削除したい時に指定
'            ToastTag           削除するToastTagを指定します。指定しない場合、アクティブExcelで出したトースト通知を全て削除します。
'            AllDel             True で、AppUserModelID に紐づく全てのトースト通知を削除します
'---------------------------------------------------------------------------------------------------
'* 注意事項：AllDel が True になると、ToastTag を無視します。
'**************************************************************************************************
Public Sub RunToastNotificationHistoryRemove(Optional ToastTag As String, Optional CollectionID As String, Optional AllDel As Boolean)
    '--- 1. 必要な設定値をまとめる ---
    With ToastConfig
        'AllDelモードなら、NULL 指定
        If AllDel Then
            .Tag = vbNullString
            .Group = vbNullString
        Else
            .Tag = ToastTag
            .Group = GetUniqueBookInfo
        End If

        .CollectionID = CollectionID
    End With


    '--- 2. WinRT 直呼び出しで通知を削除します ---
    WinRT_RemoveToastNotification ParamsConvertForWinRT(ToastConfig)
End Sub

'***************************************************************************************************
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
Public Function RunToastCollectionManagerSaveToastCollectionAsync(CollectionID As String, displayName As String, launchArgs As String, iconUri As String) As Long
    '1. 必要な設定値をまとめる
    ToastConfig.CollectionID = CollectionID

    '2. ToastCollectionを作成
    RunToastCollectionManagerSaveToastCollectionAsync = WinRT_SaveToastCollection(ParamsConvertForWinRT(ToastConfig), displayName, launchArgs, iconUri)
End Function

'***************************************************************************************************
'* 機能　　：引数で渡された値で、コレクションを使用したトースト通知のグループ化を削除します。なお、省略時、全削除となります
'---------------------------------------------------------------------------------------------------
'* 返り値　：0で、成功扱いです。
'* 引数　　：CollectionId   コレクションの ID。 ID を使用して、ToastCollectionManager を使用してグループを操作できます
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/toast-collections
'***************************************************************************************************
Public Function RunToastCollectionManagerRemoveToastCollectionAsync(Optional CollectionID As String) As Long
    '1. 必要な設定値をまとめる
    ToastConfig.CollectionID = CollectionID

    '2. ToastCollectionを削除
    RunToastCollectionManagerRemoveToastCollectionAsync = WinRT_RemoveToastCollection(ParamsConvertForWinRT(ToastConfig))
End Function

'***************************************************************************************************
'* 機能　　：トースト通知の表示に関する制限を確認します
'---------------------------------------------------------------------------------------------------
'* 返り値　：0~4 の値。意味は公式サイトを御覧ください。
'* 引数　　：CollectionId   コレクションの ID。Collection側の確認をする場合に指定。そうじゃない場合は、AppUserModelID になります。
'---------------------------------------------------------------------------------------------------
'* URL     ：https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.notificationsetting
'***************************************************************************************************
Public Property Get CheckNotificationSetting(Optional CollectionID As String) As Long
    '1. 必要な設定値をまとめる
    ToastConfig.CollectionID = CollectionID

    '2. チェック
    CheckNotificationSetting = WinRT_CheckNotificationSetting(ParamsConvertForWinRT(ToastConfig))
End Property



'***************************************************************************************************
'                     ■■■ WinRT-API へ直接アクセスするためのメソッド群 ■■■
'***************************************************************************************************






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
'* 機能　　：一般権限領域のレジストリパスへパラメーターを登録します
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
Private Function CurrentUserRegistryValue(ByVal 書き込み先 As ToastRegistryPathMode, ByVal AppUserModelID As String, ByVal valueName As String, ByVal valueType As RegistryType, ByVal valueData As String) As Long
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
    
    ' 2. 一般権限で扱えるHKCU配下のフルパスを組み立てる
    Select Case 書き込み先
        Case ToastRegistryPathMode.NotificationsSettings: strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\" & AppUserModelID
        Case ToastRegistryPathMode.ClassesAppUserModelId: strKeyPath = "Software\Classes\AppUserModelId\" & AppUserModelID
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select
    
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
        Case REG_SZ: lResult = regProvider.SetStringValue(HKEY_CURRENT_USER, strKeyPath, valueName, valueData)
        Case REG_EXPAND_SZ: lResult = regProvider.SetExpandedStringValue(HKEY_CURRENT_USER, strKeyPath, valueName, valueData)
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
        'コンピューター\HKEY_CURRENT_USER\Software\Classes\AppUserModelId 用
        Case RegistrySetting.ShowInSettings: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "ShowInSettings", REG_DWORD, valueData)
        Case RegistrySetting.DisplayName: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "DisplayName", REG_SZ, valueData)
        Case RegistrySetting.IconUri: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "IconUri", REG_SZ, valueData)
        Case RegistrySetting.IconBackgroundColor: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "IconBackgroundColor", REG_SZ, valueData)
        Case RegistrySetting.CustomActivator: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "CustomActivator", REG_SZ, valueData)
        
        'コンピューター\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings 用
        '--wpndatabase.db 連携必須--
        Case RegistrySetting.Enabled: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "Enabled", REG_DWORD, valueData)
        Case RegistrySetting.ShowBanner: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "ShowBanner", REG_DWORD, valueData) '値0で、効果発揮
        '--wpndatabase.db 連携不要--
        Case RegistrySetting.SoundFile: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "SoundFile", REG_SZ, valueData) '値は空文字で、効果発揮
        Case RegistrySetting.AllowContentAboveLock: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "AllowContentAboveLock", REG_DWORD, valueData)
        Case RegistrySetting.Rank: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "Rank", REG_DWORD, valueData)
        Case RegistrySetting.ShowInActionCenter: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "ShowInActionCenter", REG_DWORD, valueData) '値0で、効果発揮
        Case RegistrySetting.AllowUrgentNotifications: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "AllowUrgentNotifications", REG_DWORD, valueData)

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
    'ダミー値
    Const DummyValue As String = "hoge"

    Select Case PresetID
        'コンピューター\HKEY_CURRENT_USER\Software\Classes\AppUserModelId 用
        Case RegistrySetting.ShowInSettings: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "ShowInSettings", REG_Delete, DummyValue)
        Case RegistrySetting.DisplayName: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "DisplayName", REG_Delete, DummyValue)
        Case RegistrySetting.IconUri: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "IconUri", REG_Delete, DummyValue)
        Case RegistrySetting.IconBackgroundColor: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "IconBackgroundColor", REG_Delete, DummyValue)
        Case RegistrySetting.CustomActivator: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.ClassesAppUserModelId, DecisionAppUserModelID(True), "CustomActivator", REG_Delete, DummyValue)
        
        'コンピューター\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings 用
        '--wpndatabase.db 連携必須--
        Case RegistrySetting.Enabled: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "Enabled", REG_Delete, DummyValue)
        Case RegistrySetting.ShowBanner: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "ShowBanner", REG_Delete, DummyValue)
        '--wpndatabase.db 連携不要--
        Case RegistrySetting.SoundFile: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "SoundFile", REG_Delete, DummyValue)
        Case RegistrySetting.AllowContentAboveLock: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "AllowContentAboveLock", REG_Delete, DummyValue)
        Case RegistrySetting.Rank: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "Rank", REG_Delete, "hoge")
        Case RegistrySetting.ShowInActionCenter: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "ShowInActionCenter", REG_Delete, DummyValue)
        Case RegistrySetting.AllowUrgentNotifications: 最後に実行したレジストリ登録結果 = CurrentUserRegistryValue(ToastRegistryPathMode.NotificationsSettings, DecisionAppUserModelID(True), "AllowUrgentNotifications", REG_Delete, DummyValue)

        '未定義
        Case Else: Err.Raise Number:=vbObjectError + 2, Description:=vbObjectErrorID_0002
    End Select
End Property

'***************************************************************************************************
'* 機能　　：最後のレジストリ操作にて、ユーザーアカウント制御の反応を返します
'---------------------------------------------------------------------------------------------------
'* 返り値  ：結果コード
'                2：キーが見つかりません
'                0：成功
'               -1：ユーザーアカウント制御にて許可しました。または、既に管理者権限状態で実行されたがうまく行ってない
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

    'クラスオブジェクト構築時の、ExcelのPIDを取得する
    ThisExcelPID = GetCurrentProcessId()
    
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
