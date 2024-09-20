# AppNotificationBuilderVBA
VBAから、[アプリ通知(トースト通知)](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts)を表示する機能を提供します。

# DEMO

| シチュエーション例            | 動作イメージ | 
| ---------------------------- |   ------------ | 
| 処理終了 その1                    | ![alt text](doc/Demo1.png)       | 
| 処理終了 その2          | ![alt text](doc/Demo2.png)       | 
| リマインド通知          | ![alt text](doc/Demo3.png)       | 
| プログレスバー付き通知          | ![alt text](doc/Demo4.gif)       | 

他にも様々なアプリ通知の外観を設定できます。設定方法等は後述します。

# Features
- DLLインポートと専用に用意されたクラスファイルをインポートすることにより、数行で手軽にアプリ通知の表示が可能です。
- DLLインポートが使用できない環境でも、WindowsPowerShellを経由したアプリ通知の表示が可能です。
- [「自動的に閉じるMsgBox」](http://officetanaka.net/excel/vba/tips/tips21.htm)の代わりに使用することが可能です。
- 昔ながらの通知手法：[Shell_NotifyIconA](https://learn.microsoft.com/ja-jp/windows/win32/api/shellapi/nf-shellapi-shell_notifyicona)ではなく、WinRT APIを使った通知手法なので、カスタマイズ性が高いです。

# Development History
「自動的に閉じるMsgBox」という機能に一定のニーズを感じ、その代替として作ってみました。<br>
この機能は、vbsで実現しておりそのvbsがもうすぐで最新OSでは、搭載しなくなるとのことで色々と試行錯誤して作成してみました。<br>
モダンなWindows との親和性を高めるためにも、使ってみてはいかがでしょうか？

# Requirement

以下で検証済みです。

- Microsoft Office Excel 2019 以上 64bit
- Windows 10 , 11 64bit

少なくともサポート中のOffice , OSの組み合わせで、動作確認済みです。

# Load DLL

WindowsAPIの「LoadLibrary関数」を使って、読み込みます。

```bas
hDll = LoadLibrary("AppNotificationBuilderVBA.dll")
```

実際に使う場合は、"Excelファイル(.xlsm)の存在するディレクトリ"というような[動的な場所を設定する仕組み](https://liclog.net/vba-dll-create-5/)で読み込むことをおすすめします。

```bas
'動的にDLLを取得するためのWinAPI
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr

Private Sub Workbook_Open()

    Dim hDll As LongPtr
    Dim sFolderPath As String
    
    'DLLファイルを保存するフォルダパスを設定
    sFolderPath = ThisWorkbook.Path
    
    'DLLﾌｧｲﾙを読み込む
    hDll = LoadLibrary(sFolderPath & "\" & "AppNotificationBuilderVBA.dll")　'DLLファイルフルパス

    debug.print hDll
End Sub
```

hDll の中身が、0 以外であれば読み込み、成功です。

# Usage

1. [このクラスファイル](doc/SampleCode/cls_AppNotificationBuilder.cls)をVisual Basic Editorのプロジェクトにインポートして下さい。<br>

2. Visual Basic Editorを開きメニューバーの「ツール」→「参照設定」→「Microsoft XML v6.0」にCheckをいれOKを押下して下さい。<br>
![alt text](doc/Usage1.png)<br>
![alt text](doc/Usage2.png)<br>
これは、クラスファイルで[トースト コンテンツ スキーマ](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/schema-root)の作成に使います。

3. 標準モジュールを作成し、下記のように簡単なコードを記述してみましょう
```bas
Sub ShowToastTest()
    With New cls_AppNotificationBuilder
        '1. プロパティ設定
       .SetToastContent_TextTitle = "Hello World"
       .SetToastContent_TextBody = "Test message"

        '2. メソッド実行
       .RunDll_ToastNotifierShow "Hello World"
    End With
End Sub
```
実行結果は、下記のとおりです<br>
![alt text](doc/Usage3.png)<br>
この「Book1」は、Excelのブック名と連動しています。

# プロパティ説明
## AppUserModelID 関連
### AllowUse_InternetImage
HTTP上の画像ソースを使うか決めます。<br>
実際は、プリセットのOffice系AppUserModelIDを切り替えます。<br>
これは、[マニフェストにインターネット機能があるパッケージ アプリ](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/send-local-toast?tabs=uwp#:~:text=%E3%81%A8%E3%81%97%E3%81%BE%E3%81%99%E3%80%82-,%E9%87%8D%E8%A6%81,-HTTP%20%E3%82%A4%E3%83%A1%E3%83%BC%E3%82%B8%E3%81%AF)(主にStore アプリ)でないと、HTTP上の画像ソースが使えない制限があるためです。

| 値            | 設定するAppUserModelID                                                                          | 
| ------------- | ------------------------------------------------------------------------------- | 
| True          | ![alt text](doc/Ex_AppUserModelID1-1.png)<br>Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe!Microsoft.MicrosoftOfficeHub | 
| False(規定値) | ![Excel app](doc/Ex_AppUserModelID1-2.png)<br>Microsoft.Office.EXCEL.EXE.15                                           | 

True の方を使用する場合は、[こちらから](https://www.microsoft.com/store/productId/9WZDNCRD29V9?ocid=pdpshare)インストールを行って下さい。<br>
既定値、記述なしは、Falseです。
```bas
Sub ShowToastTest()
    With New cls_AppNotificationBuilder
        '切り替え
        .AllowUse_InternetImage = True



        .SetToastContent_TextTitle = "Microsoft 365から"
        Shell .GenerateCmd_ToastNotifierShow("Run"), vbHide
    End With
End Sub
```

### SetToastContent_AppUserModelID
この通知をどのAppUserModelIDで出すかを設定します。<br>
存在しない(未インストール)AppUserModelID、無効な文字列を指定すると、Toastが発行されないのでご注意ください。<br>
このプロパティが設定されてると、AllowUse_InternetImageの設定より優先されます。
```bas
Sub ShowToastTest()
    With New cls_AppNotificationBuilder
        '任意のAppUserModelID
        .SetToastContent_AppUserModelID = "Microsoft.WindowsTerminal_8wekyb3d8bbwe!App"



        .SetToastContent_TextTitle = "By Terminal"
        Shell .GenerateCmd_ToastNotifierShow("Run"), vbHide
    End With
End Sub
```
![alt text](doc/Ex_AppUserModelID2.png)<br>
上記の例では、[Windows Terminal](https://apps.microsoft.com/detail/9n0dx20hk701) のAppUserModelIDを設定します。<br>
既定値、記述なしは、vbnullstringです。

## [toast要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast)
### SetToastContent_Duration
トーストが[表示される時間](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast#:~:text=%E6%97%A2%E5%AE%9A%E5%80%A4-,duration,-%E3%83%88%E3%83%BC%E3%82%B9%E3%83%88%E3%81%8C%E8%A1%A8%E7%A4%BA)を設定します。
| 値            | 説明                            | 
| ------------- | ------------------------------- | 
| False(既定値) | shortと同等                     | 
| True          | longと同等<br>25s、表示できます | 

```bas
Sub 長く表示される通知()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        '25秒表示されます
        .SetToastContent_Duration = True



        .SetToastContent_TextTitle = "25秒間、表示"

        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```
![alt text](doc/Ex_Element-Toast1.png)

### SetToastContent_Launch
[トースト通知自体のクリック](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast#:~:text=%E3%81%AA%E3%81%97-,launch,-%E3%83%88%E3%83%BC%E3%82%B9%E3%83%88%E9%80%9A%E7%9F%A5%E3%81%AB%E3%82%88%E3%81%A3%E3%81%A6)によって、アプリケーションがアクティブ化されるときにアプリケーションに渡される文字列です。
VBAでは、起動スキーマ(https:// , ms-excel:// など)を設定するぐらいの役目です。

#### 設定可能な引数
| 引数名            | 解説                                                                                                                                                                                                                                                                                                                                                             | 既定値   | 
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------- | 
| ArgActivationType | ユーザーが特定の操作を行った際に使用されるアクティブ化の種類を決定します。<br>・"foreground" - 既定値。 フォアグラウンド アプリが起動します。<br>・"background" - 対応するバックグラウンド タスクがトリガーされ、ユーザーを中断することなくバックグラウンドでコードを実行できます。<br>・"protocol" - プロトコルのアクティブ化を使用して別のアプリを起動します。 | protocol | 

```bas
Sub リンクを開く()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        'URL等を指定
        .SetToastContent_Launch = "https://www.google.com/"



        .SetToastContent_TextTitle = "このトーストをクリックすると、指定リンクに対応するアプリが起動"
        
        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```
![alt text](doc/Ex_Element-Toast2.png)

### SetToastContent_DisplayTimestamp
Windows プラットフォームによって通知が受信された時刻ではなく、通知コンテンツが実際に配信された日時を表すカスタム タイムスタンプで既定の[タイムスタンプをオーバーライド](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/custom-timestamps-on-toasts?tabs=xml)します。

```bas
Sub アプリ通知のカスタムタイムスタンプ()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        'シリアル値で設定。基本は過去に設定
        .SetToastContent_DisplayTimestamp = Now() - 0.1



        .SetToastContent_TextTitle = "Hello World"
        .SetToastContent_TextBody = "このメッセージは、以前から通知されてました。"
        .SetToastContent_TextAttribute = "カスタムタイムスタンプテスト"

        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
        Debug.Print ActionCmd
    End With
End Sub
```
![alt text](doc/Ex_Element-Toast3.png)

### SetToastScenario
トーストが使用される[シナリオ](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#scenarios)を設定します。列挙型に対応します。

| シナリオ名   | 主な特徴                                                                                                                                                                                  | 
| ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | 
| Default(既定値)      | 一般的な挙動通知                                                                                                                                                                      | 
| [Reminder](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#reminders)     | ・通知を永遠に表示する。<br>・action要素がないと効果は発動しない<br>・任意の通知音に設定可能<br>![reminder,alarm](doc/Ex_Element-Toast4-2.png)                                                                                          | 
| [Alarm](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#alarms)        | ・通知を永遠に表示する。<br>・action要素がないと効果は発動しない<br>・通知音は、アラーム系(Alarm)のみ<br>・応答不可モードでも必ず表示<br>![reminder,alarm](doc/Ex_Element-Toast4-2.png)                                                 | 
| [IncomingCall](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#incoming-calls) | ・通知を永遠に表示する。<br>・action要素がなくても効果発動<br>・通知音は呼び出し系(Call)のみ<br>・最後のボタン位置のみ、Windowsのテーマ色に基づく着色が施され、位置が必ず下側になる。<br>![reminder,alarm](doc/Ex_Element-Toast4-3.png) | 
| [Urgent](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#important-notifications)       | ・通知に感嘆符が付与<br>・応答不可モードでの表示/非表示の、切り替え可能<br>・Build 22546 以降のOS で有効<br>![reminder,alarm](doc/Ex_Element-Toast4-4.png)                                                                                                               | 
```bas
Sub シナリオテスト()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        'トーストのシナリオを設定(ctrl + Space で候補を表示できます)
        .SetToastScenario = Urgent



        '紐付け用識別子(解説は後述)
        Const ReminderID As String = "アラーム"

        'select要素を準備(解説は後述)
        .SetToastSelectionBox(1, "1 分後") = 1
        .SetToastSelectionBox(5, "5 分後") = 2
        .SetToastSelectionBox(10, "10 分後") = 3
        .SetToastSelectionBox(30, "30 分後") = 4
        .SetToastSelectionBox(60, "1 時間後") = 5

        'input要素を作成し、上記で準備したselect要素を挿入(解説は後述)
        .SetIToastInput(ReminderID, True, , "選択肢から、再通知する時間を選択", 10) = 1

        '再通知用と、解除用を用意(解説は後述)
        .SetIToastActions("", "snooze", "system", , , , ReminderID) = 1
        .SetIToastActions("", "dismiss", "system") = 2

        'テキスト要素を用意
        .SetToastContent_TextTitle = "Hello World"


        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")
        
        .RunDll_ToastNotifierShow "sample025"
        'Shell ActionCmd, vbHide

    End With
End Sub
```

### AllowToastContent_UseButtonStyle
toast要素の[useButtonStyle](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-toast#:~:text=%E3%81%AA%E3%81%97-,useButtonStyle,-%E3%82%B9%E3%82%BF%E3%82%A4%E3%83%AB%E4%BB%98%E3%81%8D%E3%83%9C%E3%82%BF%E3%83%B3)属性の設定を行います。<br>
Trueで、スタイル付きボタンを使用します。後述の[action 要素](#SetIToastActions)の 「hint-buttonStyle」 属性に影響します。<br>
```bas
Sub UseButtonStyle()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        'スタイル付きボタンを有効化
        .AllowToastContent_UseButtonStyle = True



        .SetToastContent_TextTitle = "緑と赤のボタン"
        
        '設定方法は後述
        .SetIToastActions("Green", "", , , , , , Success) = 1
        .SetIToastActions("Red", "", , , , , , Critical) = 2
        
        
        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```
![alt text](doc/Ex_Element-Toast5.png)

## [image要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-image)
### SetToastContent_ImageAppLogo
image要素のうち、AppLogo(appLogoOverride)に設定する画像のパスと、丸いロゴフラグの設定を行います。<br>
ロゴ画像のパス指定は、ローカルパス(C:\\)、HTTPソースに対応してます


#### 設定可能な引数
| 引数名             | 説明                                                                                                           | 既定値       | 
| ------------------ | -------------------------------------------------------------------------------------------------------------- | ------------ | 
| Arg_LogoCircle     | True：画像は円にトリミングされます。<br>False：画像はトリミングされず、正方形として表示されます。              | False        | 
| Flag_addImageQuery | Windows がトースト通知で指定されたイメージ URI にクエリ文字列を追加できるようにするには、"true" に設定します。 | False        | 
| Arg_Alt            | 支援技術のユーザー向けの画像の説明。                                                                      | vbnullstring | 

```bas
Sub 丸いロゴ画像()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        'ロゴ画像のパスを指定します。(Arg_LogoCircle=False)
        .SetToastContent_ImageAppLogo = "C:\Windows\SystemApps\Microsoft.XboxGameCallableUI_cw5n1h2txyewy\Assets\Logo.scale-100.png"

        'ロゴ画像のパスを指定し、円にトリミング。(Arg_LogoCircle=True)
        '.SetToastContent_ImageAppLogo(True) = "C:\Windows\SystemApps\Microsoft.XboxGameCallableUI_cw5n1h2txyewy\Assets\Logo.scale-100.png"



        .SetToastContent_TextTitle = "ロゴ画像テスト"
        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")
        .RunDll_ToastNotifierShow "sample"

        'Shell ActionCmd, vbHide
    End With
End Sub
```

| Arg_LogoCircle=False | Arg_LogoCircle=True | 
| -------------------- | ------------------- | 
| ![正方形](doc/Ex_Element-Image1-1.png)                   | ![円にトリミング](doc/Ex_Element-Image1-2.png)                  | 

### SetToastContent_ImageInline
image要素のうち、テキスト要素の後に表示する画像パスと、丸いロゴフラグの設定を行います。br>
先ほどと同様、インライン画像のパス指定も、ローカルパス(C:\\)、HTTPソースに対応してます。<br>
引数の内容も同様です。

```bas
Sub インライン画像()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        '通常
        .SetToastContent_ImageInline(False, , "win の壁紙") = "C:\Windows\Web\Screen\img100.jpg"
        '円にトリミング
        ''.SetToastContent_ImageInline(True, , "win の壁紙") = "C:\Windows\Web\Screen\img100.jpg"



        .SetToastContent_TextTitle = "インライン画像テスト"

        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```

| Arg_LogoCircle=False | Arg_LogoCircle=True | 
| -------------------- | ------------------- | 
| ![正方形](doc/Ex_Element-Image2-1.png)                   | ![円にトリミング](doc/Ex_Element-Image2-2.png)                  | 

### SetToastContent_ImageHero
ヒーローイメージとして表示させる画像を設定します。<br>
先ほどと同様、ヒーロー画像のパス指定も、ローカルパス(C:\\)、HTTPソースに対応してます。

#### 設定可能な引数
| 引数名             | 説明                                                                                                           | 既定値       | 
| ------------------ | -------------------------------------------------------------------------------------------------------------- | ------------ | 
| Flag_addImageQuery | Windows がトースト通知で指定されたイメージ URI にクエリ文字列を追加できるようにするには、"true" に設定します。 | False        | 
| Arg_Alt            | 支援技術のユーザー向けの画像の説明。                                                                      | vbnullstring | 

```bas
Sub 上部に画像()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        '上部に画像を配置します
        .SetToastContent_ImageHero(, "win11壁紙") = "C:\Windows\Web\Screen\img100.jpg"



        .SetToastContent_TextTitle = "上部に画像を配置"
        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```

![alt text](doc/Ex_Element-Image3.png)

## [text要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-text)
アプリ通知で使用されるテキストを指定します。

| プロパティ名                  | 配置位置 | 最大行数 | 
| ----------------------------- | -------- | -------- | 
| SetToastContent_TextTitle     | タイトル | 2行      | 
| SetToastContent_TextBody      | 内容     | 4行      | 
| SetToastContent_TextAttribute | 下部     | 2行      | 

### 設定可能な引数
| 引数名             | 説明                                                                                                           | 既定値       | 
| ------------------ | -------------------------------------------------------------------------------------------------------------- | ------------ | 
| HintCallScenarioCenterAlign | 横中央揃えの配置にする設定です。trueにしつつ、シナリオモードを「IncomingCall」にしないと効果ありません。 | False        | 

```bas
Sub 最大行数テキスト()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String

    With AppNotification
        'テキスト設定
        .SetToastContent_TextTitle(False) = "タイトル 1Line" & vbCrLf & "タイトル 2Line" & vbCrLf & "タイトル 3Line"
        .SetToastContent_TextBody(False) = "コンテンツ 1Line" & vbCrLf & "コンテンツ 2Line" & vbCrLf & "コンテンツ 3Line" & vbCrLf & "コンテンツ 4Line" & vbCrLf & "コンテンツ 5Line"
        .SetToastContent_TextAttribute(False) = "コンテンツソース 1Line" & vbCrLf & "コンテンツソース 2Line" & vbCrLf & "コンテンツソース 3Line"



        '中央揃えにするとき
        '.SetToastScenario = IncomingCall
        
        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"

        Debug.Print ActionCmd
    End With
```

| HintCallScenarioCenterAlign = False             | HintCallScenarioCenterAlign = True かつ、SetToastScenario = IncomingCall |
| ------------------------------------------------- | --------------------------------------- |
| ![alt text](doc/Ex_Element-text1-1.png)           | ![alt text](doc/Ex_Element-text1-2.png) |

## [audio要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio)
### SetToastAudio
アプリ通知を表示するときに再生するサウンドを指定します。 ミュートも対応してます。<br>
ただし、ファイルシステム上の音声ファイルのパスや URLの指定は使えません。システムで決められた通知音のみ設定可能です。

#### 設定可能な引数
| 引数名             | 説明                                                                                                           | 既定値       | 
| ------------------ | -------------------------------------------------------------------------------------------------------------- | ------------ | 
| ArgLoop            | トーストが表示されている限り、サウンドを繰り返す場合は true に設定します。<br> 1 回だけ再生する場合は false。  | False        | 

```bas
Sub 通知音変更テスト()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String

    With AppNotification
        '通知音設定(ctrl + Space　で候補が出ます)
        .SetToastAudio = NotificationLoopingAlarm01
        
        
        
        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")
        .SetToastContent_TextTitle = "通知音変更"
        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```
設定可能な通知音は、[こちら](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-audio#:~:text=false-,src,-%E6%97%A2%E5%AE%9A%E3%81%AE%E3%82%B5%E3%82%A6%E3%83%B3%E3%83%89)をどうぞ。<br>
また、False で指定すると、ミュート扱いになります。

## [action要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-action)
### SetIToastActions
トーストに表示されるボタンを指定します。

#### 設定可能な引数
| 引数名             | 説明                                                                                                                                                                                                                                                                                                                                                             | 既定値                       | 
| ------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------- | 
| ArgContent         | ボタンに表示されるコンテンツ                                                                                                                                                                                                                                                                                                                                     | ※必須項目                   | 
| ArgArguments       | ユーザーがこのボタンをクリックした場合にアプリが後から受け取る、アプリで定義された引数の文字列です。                                                                                                                                                                                                                                                             | ※必須項目だが、空文字でもOK | 
| ArgActivationType  | ユーザーが特定の操作を行った際に使用されるアクティブ化の種類を決定します。<br>・"foreground"：既定値。 フォアグラウンド アプリが起動します。<br>・"background"：対応するバックグラウンド タスクがトリガーされ、ユーザーを中断することなくバックグラウンドでコードを実行できます。<br>・"protocol"：プロトコルのアクティブ化を使用して別のアプリを起動します。 <br>・"system"：ArgArgumentsに特定の文字列を入れると、リマインダー機能が使えます。(後述)| protocol                     | 
| ArgPendingUpdate   | ・TRUE：ユーザーがトースト上のボタンをクリックすると、通知は "保留中の更新" 表示状態のままです。 この "更新の保留中" の表示状態が長時間続くことを避けるため、バックグラウンド タスクから即座にトーストを更新する必要があります。<br>・FALSE：ユーザーがトーストに対して操作を行うと、トーストが無視されます。                                                    | FALSE                        | 
| ArgContextMenu     | ・TRUE：トースト ボタンではなく、トースト通知のコンテキスト メニューに追加されたコンテキスト メニュー アクションになります。<br>・FALSE：従来通り、トースト ボタンに配置                                                                                                                                                                                         | FALSE                        | 
| ArgIcon            | トースト ボタン アイコンのイメージ ソースの URI。<br>ローカルパス、HTTPソースに対応します。                                                                                                                                                                                                                                                                      | vbnullstring                 | 
| ArgHintInputId     | 入力の横にある [位置への 入力 ] ボタンの ID に設定します。                                                                                                                                                                                                                                                                                                       | vbnullstring                 | 
| ArgHintButtonStyle | ボタンのスタイル。<br>事前に[toast要素のuseButtonStyle属性](#AllowToastContent_UseButtonStyle)にtrue を設定する必要があります。<br><br>・Success：緑<br>・Critical：赤<br>・NoStyle：無色                                                                                                                                                                                                             | NoStyle                      | 
| ArgHintToolTip     | ボタンに空のコンテンツ文字列がある場合のボタンのヒント。                                                                                                                                                                                                                                                                                                         | vbnullstring                 | 
```bas
Sub MakeActionTest()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String

    With AppNotification
        'ボタン作成
        .SetIToastActions("Green", "", , , , , , Success) = 1
        
        'コンテキストメニュー側に移す
        .SetIToastActions("コンテキストメニューにあります", "", , , True) = 2
        
        'ボタンにカーソルをあてるとToolTip表示し、アイコンセット
        .SetIToastActions("", "ms-search://Search", , , , "C:\Windows\IdentityCRL\WLive48x48.png", , , "クリックで、検索を開く") = 3
        
        'このボタンを押下すると、Youtubeにアクセスします
        .SetIToastActions("YouTube開く", "https://www.youtube.com/", , , , , , Critical, "ツールチップ") = 4


        
        

        .AllowToastContent_UseButtonStyle = True
        .SetToastContent_TextTitle = "ActionTest"
        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```

最後の数字は、ボタンの配置順を表します。1~5まで有効です。<br>
![alt text](doc/Ex_Element-Action1-1.png) ![alt text](doc/Ex_Element-Action1-2.png)

## [subgroup要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-subgroup)
現時点では、作成アシストには非対応です。

## [header要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-header)
### SetToastHeader
アクション センター内で複数の通知をグループにまとめるカスタム ヘッダーを指定します。<br>
なお、Classファイルを定義する際(Class_Initialize)、予め既定値を入れるように仕込んでいるため基本、呼び出す必要はありません。
また、ヘッダー名に表示される既定値名称は、ThisWorkbook.Nameですが、拡張子は省略します。

#### 設定可能な引数
| 引数名            | 説明                                                                                                                           | 既定値            | 
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------ | ----------------- | 
| ArgID             | このヘッダーを一意に識別します。 2 つの通知が同じヘッダー ID を持つ場合、アクション センターで同じヘッダーの下に表示されます。 | ThisWorkbook.Name | 
| ArgArguments      | ユーザーがこのヘッダーをクリックするとアプリに返されます。 null にすることはできません。                                       | ThisWorkbook.Path | 
| ArgActivationType | このヘッダーがクリックされた場合に使用するアクティブ化の種類。                                                                 | protocol          | 

```bas
Sub ヘッダーテスト()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String

    With AppNotification
        'ヘッダー名を変更
        .SetToastHeader = "えくせる"



        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")
        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
    End With
End Sub
```
![alt text](doc/Ex_Element-header.png)<br>


## [input要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-input)
### SetIToastInput
トースト通知に表示される入力 (テキスト ボックスまたは選択メニュー) を指定します。<br>
VBAでは、リマインダー用途でしか使い所がないと思います。

#### 設定可能な引数
| 引数名                | 説明                                                                    | 既定値       | 
| --------------------- | ----------------------------------------------------------------------- | ------------ | 
| ArgID                 | 入力に関連付けられている ID                                             | ※必須項目     | 
| ChoseFlag             | ・True："selection"<br>・False："text"                                  | False        | 
| ArgPlaceHolderContent | テキスト入力用に表示されるプレースホルダー。<br>ChoseFlag=False時、有効 | vbnullstring | 
| ArgTitle              | 入力のラベルとして表示されるテキスト                                    | vbnullstring | 
| ArgDefaultInput       | デフォルトの入力値                                                      | vbnullstring | 

```bas
Sub メッセージ()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        'テキスト入力要素を作成
        .SetIToastInput("textBox",, "reply","はみがきなう！") = 1



        'InputのIDと、Actionのhint-inputIdを同じ値にして、同じIndex値に対応するInput要素の横にボタンを配置できます
        .SetIToastActions("Send", "", , , , , "textBox") = 1

        'ネット上の画像を使用する
        .AllowUse_InternetImage = True
        .SetToastContent_ImageAppLogo(True) = "https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEhkdhNl7CCgOAZmjObZRRINCu9udW_Mum-_FSCEvLFULZRP_wEuO_Y1grKy28zSCw2fyBN5jM2RS1PYmE9prAk5uSI8ImDn1wcyZPy8ctGJ-vFaY36ULy_rjvrilHjXjuN0_O-p6sQi3Hc3/s857/ha_hamigaki_suit_woman.png"
        .SetToastContent_ImageHero = "https://unsplash.it/360/180?image=1043"

        .SetToastContent_TextTitle = "メッセージ返信デザイン"

        ActionCmd = .GenerateCmd_ToastNotifierShow("sample")

        'Shell ActionCmd, vbHide
        .RunDll_ToastNotifierShow "sample"
        
        Debug.Print ActionCmd
    End With
End Sub
```
最後の数字は、Input要素の配置順を表します。1~5まで有効です。<br>
![alt text](doc/Ex_Element-Input1-1.png) ![alt text](doc/Ex_Element-Input1-2.png)<br>


## [selection要素](https://learn.microsoft.com/ja-jp/uwp/schemas/tiles/toastschema/element-selection)
### SetToastSelectionBox
選択項目の id とテキストを指定します。全て必須項目です。
基本、リマインダー用途のみとなります。

#### 設定可能な引数
| 引数名         | 説明                                               | 備考                       | 
| -------------- | -------------------------------------------------- | -------------------------- | 
| ReminderMinute | 何分後にリマインダー通知させるか、値で指定します。 | 現状、数値以外は扱いません。<br>0で、未定義扱いとします。 | 
| ArgChoseName   | 選択項目の内容                                     |                            | 

#### [リマインダーの設定方法](https://learn.microsoft.com/ja-jp/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml#snoozedismiss)
Input要素と、selection要素を使ったリマインダー方法を紹介します。<br>
コード内コメントにある手順を参考にどうぞ。
```bas
Sub リマインドテスト()
    Dim AppNotification As New cls_AppNotificationBuilder
    Dim ActionCmd As String
    
    With AppNotification
        '1. トーストシナリオをリマインダーか、アラームにする
        .SetToastScenario = Reminder

        '2. 紐付け用識別子を設定
        Const ReminderID As String = "リマインダー"

        '3. select要素を準備し、リマインドする"分"と名称をセット(最大、5つ)
        .SetToastSelectionBox(1, "1 分後") = 1
        .SetToastSelectionBox(5, "5 分後") = 2
        .SetToastSelectionBox(10, "10 分後") = 3
        .SetToastSelectionBox(30, "30 分後") = 4
        .SetToastSelectionBox(60, "1 時間後") = 5

        '4. input要素を作成し、上記で準備したselect要素を挿入し、先ほど作成した紐付け用識別子をInput-IDにセット
        .SetIToastInput(ReminderID, True, , "選択肢から、リマインドする時間を選択", 10) = 1

        '5. 再通知用と、解除用を用意("snooze", "system",ReminderID にセットされてる引数位置は、必ずこの値にする)
        .SetIToastActions("", "snooze", "system", , , , ReminderID) = 1
        .SetIToastActions("", "dismiss", "system") = 2

        '6. テキスト要素を用意(任意)
        .SetToastContent_TextTitle = "リマインダーテスト"
        .SetToastContent_TextBody = "「再通知」で、選択した時間で、再通知" & vbcrlf & "解除で、何もしない"

        '7. コマンド文字列を生成(Windows PowerShell経由で実行する場合)
        ActionCmd = .GenerateCmd_ToastNotifierShow("リマインド")

        '8. コマンド実行
        .RunDll_ToastNotifierShow "リマインド"
        'Shell ActionCmd,vbHide
    End With
End Sub
```
![alt text](doc/Ex_Element-Selection1-1.png) ![alt text](doc/Ex_Element-Selection1-2.png)

# メソッド説明
## GenerateCmd_ToastNotifierShow(ToastTag , Optional CollectionID , Optional ScheduleDate , Optional ExpirationDate , Optional Suppress)
引数に渡された値で、単純なトースト通知を表示するコマンド文字列を返します。指定日時に通知するスケジュール機能も対応します<br>
Shell関数と併用して使用して下さい。Windows PowerShell環境があれば、どのWindows マシンでも動作が可能です。

| 引数                                                                                                                                                         | 意味                                                                                 | 型         | 既定値       | 
| ------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------ | ---------- | ------------ | 
| [ToastTag](https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.tag)                                                         | 通知 グループ内のこの通知の一意識別子を取得または設定します。                        | 文字列     | ※必須項目   | 
| [CollectionID](https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotificationmanagerforuser.gettoastnotifierfortoastcollectionidasync) | 送信する通知グループの ID。                                                          | 文字列     | vbnullstring | 
| [ScheduleDate](https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.-ctor)                                          | Windows でトースト通知を表示する日付と時刻。                                         | シリアル値 | 0            | 
| [ExpirationDate](https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.scheduledtoastnotification.expirationtime)                               | 通知の有効期限。                                                                     | シリアル値 | 0            | 
| [Suppress](https://learn.microsoft.com/ja-jp/uwp/api/windows.ui.notifications.toastnotification.suppresspopup)                                               | トーストのポップアップ UI をユーザーの画面に表示するかどうかを取得または設定します。 | フラグ値   | False        | 


```bas
With New cls_AppNotificationBuilder
    Shell .GenerateCmd_ToastNotifierShow("NoSetting"), vbHide
End With
```
なお、上記のように特に事前設定なくとも動作は可能です。この場合は次のように表示されます<br>
![alt text](doc/ExampleMethod1.png)

## RunDll_ToastNotifierShow
GenerateCmd_ToastNotifierShow と同様の機能です。
こちらは、DLLファイルを読み込んだときに使う専用メソッドです。パフォーマンスが向上するので使える環境であればこちらがおすすめです。<br>
引数等は、GenerateCmd_ToastNotifierShow と同じなので省略します。
