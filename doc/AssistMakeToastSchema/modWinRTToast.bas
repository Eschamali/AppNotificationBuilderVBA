Attribute VB_Name = "modWinRTToast"
Option Explicit

' modWinRTToast
' WinRT Interfaces x64.txt / ToastCompat_WinRT.bas に基づく VTable 直呼び出し。
' AppNotificationBuilderVBA.dll の Show / Update / Remove / CheckNotificationSetting を置き換える。
'
' スケジュール通知（no-TLB）:
'   TLB 版 Calendar SetToNow + AddMinutes と同じく、Now からの分差で UTC tick を生成する。
'   絶対日時のフォールバックは TzSpecificLocalTimeToSystemTime → SystemTimeToFileTime。
' TLB 早期バインディングに切り替える場合は WRT_USE_TLB_SCHEDULE = 1（参照設定 → WinRT Interfaces x64）。
#Const WRT_USE_TLB_SCHEDULE = 0

Private Declare PtrSafe Function RoInitialize Lib "combase.dll" (ByVal initType As Long) As Long
Private Declare PtrSafe Function RoUninitialize Lib "combase.dll" () As Long
Private Declare PtrSafe Function WindowsCreateString Lib "combase.dll" (ByVal sourceString As LongPtr, ByVal length As Long, ByRef hstring As LongPtr) As Long
Private Declare PtrSafe Function WindowsDeleteString Lib "combase.dll" (ByVal hstring As LongPtr) As Long
Private Declare PtrSafe Function RoGetActivationFactory Lib "combase.dll" (ByVal activatableClassId As LongPtr, ByRef iid As Any, ByRef factory As LongPtr) As Long
Private Declare PtrSafe Function RoActivateInstance Lib "combase.dll" (ByVal activatableClassId As LongPtr, ByRef instance As LongPtr) As Long
#If WRT_USE_TLB_SCHEDULE Then
Private Declare PtrSafe Function RoActivateInstanceObj Lib "combase.dll" Alias "RoActivateInstance" (ByVal activatableClassId As LongPtr, ByRef instance As Any) As Long
#End If
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Integer, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Sub VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vtime As Double, ByRef lpSystemTime As WinRT_SYSTEMTIME)
Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (ByRef lpSystemTime As WinRT_SYSTEMTIME, ByRef lpFileTime As WinRT_FILETIME) As Long
Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (ByVal lpTimeZoneInformation As LongPtr, ByRef lpLocalTime As WinRT_SYSTEMTIME, ByRef lpUniversalTime As WinRT_SYSTEMTIME) As Long
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As WinRT_SYSTEMTIME)
Private Declare PtrSafe Function WindowsGetStringRawBuffer Lib "combase.dll" (ByVal hstring As LongPtr, ByRef length As Long) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function WinRT_GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long

Private Type WinRT_SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WinRT_FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WinRT_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type WRT_DateTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

#If Win64 Then
    Private Const WinRT_vbPtr As Integer = 20
#Else
    Private Const WinRT_vbPtr As Integer = 3
#End If

Private Const VT_QI As Long = 0
Private Const VT_RELEASE As Long = 2
Private Const VT_IToastNotifier_Show As Long = 6
Private Const VT_IToastNotifier_GetSetting As Long = 8
Private Const VT_IToastNotifier_AddToSchedule As Long = 9
Private Const VT_IScheduledToastNotificationFactory_Create As Long = 6
Private Const VT_IScheduledToastNotification_SetId As Long = 10
Private Const VT_IToastNotification2_SetTag As Long = 6
Private Const VT_IToastNotification2_SetGroup As Long = 8
Private Const VT_IToastNotification2_SetSuppressPopup As Long = 10
Private Const VT_IToastNotification_SetExpirationTime As Long = 7
Private Const VT_IToastNotification4_SetData As Long = 7
Private Const VT_IScheduledToastNotification4_SetExpirationTime As Long = 7
Private Const VT_IToastNotification6_SetExpiresOnReboot As Long = 7
Private Const VT_IToastNotificationHistory_RemoveGroupedTagWithId As Long = 8
Private Const VT_IToastNotificationHistory_RemoveGroupWithId As Long = 7
Private Const VT_IToastNotificationHistory_ClearWithId As Long = 12
Private Const VT_IToastNotificationManagerStatics2_GetHistory As Long = 6
Private Const VT_IToastNotificationManagerStatics5_GetDefault As Long = 6
Private Const VT_IToastNotificationManagerForUser2_GetToastCollectionManagerWithAppId As Long = 9
Private Const VT_IUriFactory_CreateUri As Long = 6
Private Const VT_IToastCollectionFactory_CreateInstance As Long = 6
Private Const VT_IToastCollectionManager_SaveToastCollectionAsync As Long = 6
Private Const VT_IToastCollectionManager_RemoveToastCollectionAsync As Long = 9
Private Const VT_IToastCollectionManager_RemoveAllToastCollectionsAsync As Long = 10
Private Const VT_IAsyncOperation_GetResults As Long = 8
Private Const VT_IAsyncInfo_GetStatus As Long = 7
Private Const VT_IAsyncInfo_GetErrorCode As Long = 8
Private Const WinRT_AsyncStatus_Started As Long = 0
Private Const WinRT_AsyncStatus_Completed As Long = 1
Private Const WinRT_AsyncStatus_Canceled As Long = 2
Private Const WinRT_AsyncStatus_Error As Long = 3
Private Const WinRT_IID_IAsyncInfo As String = "{00000036-0000-0000-C000-000000000046}"
Private Const VT_IToastNotifier2_UpdateWithTagAndGroup As Long = 6
Private Const VT_IToastNotifier2_UpdateWithTag As Long = 7
Private Const VT_INotificationData_SetSequenceNumber As Long = 8
Private Const VT_IPropertyValueStatics_CreateDateTime As Long = 21

' --- トーストアクティブ化（イベントコールバック）用 vtable インデックス ---
' IToastNotification (IInspectable 派生, 固有メソッドは index 6 起点)
Private Const VT_IToastNotification_AddActivated As Long = 11
Private Const VT_IToastNotification_AddDismissed As Long = 9
Private Const VT_IToastNotification_AddFailed As Long = 13
' IToastNotification2: 6 put_Tag / 7 get_Tag / 8 put_Group / 9 get_Group
Private Const VT_IToastNotification2_GetTag As Long = 7
Private Const VT_IToastNotification2_GetGroup As Long = 9
' イベント引数（いずれも IInspectable 派生で固有メソッドは index 6）
Private Const VT_IToastActivatedEventArgs_GetArguments As Long = 6
Private Const VT_IToastActivatedEventArgs2_GetUserInput As Long = 6
Private Const VT_IToastDismissedEventArgs_GetReason As Long = 6
Private Const VT_IToastFailedEventArgs_GetErrorCode As Long = 6
' コレクション反復（ValueSet → IIterable<IKeyValuePair<HSTRING,IInspectable>>）
Private Const VT_IIterable_First As Long = 6
Private Const VT_IIterator_GetCurrent As Long = 6
Private Const VT_IIterator_GetHasCurrent As Long = 7
Private Const VT_IIterator_MoveNext As Long = 8
Private Const VT_IKeyValuePair_GetKey As Long = 6
Private Const VT_IKeyValuePair_GetValue As Long = 7
Private Const VT_IPropertyValue_GetString As Long = 19
Private Const VT_IStringable_ToString As Long = 6

' --- トーストアクティブ化用 IID ---
Private Const WinRT_IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
Private Const WinRT_IID_IToastNotification2 As String = "{9DFB9FD1-143A-490E-90BF-B9FBA7132DE7}"
Private Const WinRT_IID_IToastActivatedEventArgs As String = "{E3BF92F3-C197-436F-8265-0625824F8DAC}"
Private Const WinRT_IID_IToastActivatedEventArgs2 As String = "{AB7DA512-CC61-568E-81BE-304AC31038FA}"
Private Const WinRT_IID_IToastDismissedEventArgs As String = "{3F89D935-D9CB-4538-A0F0-FFE7659938F8}"
Private Const WinRT_IID_IToastFailedEventArgs As String = "{35176862-CFD4-44F8-AD64-F500FD896C3B}"
Private Const WinRT_IID_TypedEventHandler_Activated As String = "{AB54DE2D-97D9-5528-B6AD-105AFE156530}"
Private Const WinRT_IID_TypedEventHandler_Dismissed As String = "{61C2402F-0ED0-5A18-AB69-59F4AA99A368}"
Private Const WinRT_IID_TypedEventHandler_Failed As String = "{95E3E803-C969-5E3A-9753-EA2AD22A9A33}"
Private Const WinRT_IID_IIterable_KVP_IInspectable As String = "{FE2F3D47-5D47-5499-8374-430C7CDA0204}"
Private Const WinRT_IID_IPropertyValue As String = "{4BD682DD-7554-40E9-9A9B-82654EDE7E62}"
Private Const WinRT_IID_IStringable As String = "{96369F54-8EB6-48F0-ABCE-C1B211E627C3}"

' デリゲートの種類（KeepAlive / vtable 配列のインデックス）
Private Const WRT_DELEGATE_ACTIVATED As Long = 0
Private Const WRT_DELEGATE_DISMISSED As Long = 1
Private Const WRT_DELEGATE_FAILED As Long = 2
Private Const WRT_DELEGATE_COUNT As Long = 3
Private Const WRT_GroupDelimiter As String = "|"
Private Const WRT_GPTR As Long = &H40
Private Const WRT_S_OK As Long = 0
Private Const WRT_E_NOINTERFACE As Long = &H80004002
' Dismissed / Failed 受信時に呼ぶ既定マクロ名（DLL 版 EventNotice.cpp と同名）
Private Const WRT_MacroName_Dismissed As String = "ExcelToast_Dismissed"
Private Const WRT_MacroName_Failed As String = "ExcelToast_Failed"

Private Const WRT_FileTimeTicksPerMinute As LongLong = 600000000
Private Const WinRT_vtCy As Integer = 6
Private Const WinRT_MaxScheduleIdLen As Long = 16
Private Const WinRT_RO_INIT_SINGLETHREADED As Long = 0
Private Const WinRT_RPC_E_CHANGED_MODE As Long = &H80010106

' Excel セッション中は WinRT を維持し、Show→Update 間で RoUninitialize しない
Private WinRT_RoInitialized As Boolean
Private WinRT_ToastDataSequenceNumber As Long

' --- ネイティブ COM デリゲート（IUnknown + Invoke の 4 エントリ vtable）---
' Activated / Dismissed / Failed の 3 種類を自前 vtable で構築し add_* に渡す。
' QueryInterface は IUnknown と各デリゲート IID のみ受理し、それ以外は E_NOINTERFACE。
' これにより別アパートメントからのコールバックが COM 標準マーシャリングで
' STA(メインスレッド)へ転送され、Application.Run を安全に実行できる。
Private WRT_Act_pVTable(0 To 2) As LongPtr
Private WRT_Act_pObject(0 To 2) As LongPtr
Private WRT_Act_RefCount(0 To 2) As Long
Private WRT_Act_iidDelegate(0 To 2) As WinRT_GUID
Private WRT_Act_iidUnknown As WinRT_GUID
Private WRT_Act_DelegatesReady As Boolean
' Activated 受信まで生かしておく必要があるオブジェクト（トースト本体）
Private WRT_Act_KeepAlive As Collection

' WinRT Interfaces x64.tlb 参照時は WinRT_* 型名が TLB の WinRT 名前空間と衝突するため ANB 接頭辞を使う。
Public Type WinRT_ToastConfig
    AppUserModelID As String
    XmlTemplate As String
    Tag As String
    Group As String
    Schedule_ID As String
    CollectionID As String
    ExpiresOnReboot As Boolean
    SuppressPopup As Boolean
    Schedule_DeliveryTime As Date
    Schedule_DeliveryTimeLocal As Date
    ExpirationTime As Date
End Type

Public Type WinRT_DataBinding
    TitleText As String
    ContentsText As String
    AttributionText As String
    ProgressTitle As String
    ProgressValueStringOverride As String
    ProgressStatus As String
    progressValue As Double
    SequenceNumber As Long
End Type

Public Sub WinRT_ShowToastNotification(ByRef Config As WinRT_ToastConfig, ByRef Binding As WinRT_DataBinding)
    Dim initialized As Boolean
    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized
    WinRT_ValidateToastIdentity Config, "WinRT_ShowToastNotification"

    If WinRT_DataBindingHasContent(Binding) Then
        Binding.SequenceNumber = WinRT_NextToastDataSequence()
    End If

    If Config.Schedule_DeliveryTime > 0 Or Config.Schedule_DeliveryTimeLocal > 0 Then
        WinRT_ShowScheduledToast Config
    Else
        WinRT_ShowImmediateToast Config, Binding
    End If

Cleanup:
    If Err.Number <> 0 Then Err.Raise Err.Number, "WinRT_ShowToastNotification", Err.Description
    WinRT_FinalizeRo initialized
End Sub

Public Function WinRT_UpdateToastNotification(ByRef Config As WinRT_ToastConfig, ByRef Binding As WinRT_DataBinding) As Long
    Dim initialized As Boolean
    Dim pNotifier As LongPtr
    Dim pNotifier2 As LongPtr
    Dim pData As LongPtr
    Dim hStrTag As LongPtr
    Dim hStrGroup As LongPtr
    Dim iidNotifier2 As WinRT_GUID
    Dim updateResult As Long

    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized
    WinRT_ValidateToastIdentity Config, "WinRT_UpdateToastNotification"

    If WinRT_DataBindingHasContent(Binding) Then
        Binding.SequenceNumber = WinRT_NextToastDataSequence()
    End If

    pNotifier = WinRT_CreateToastNotifier(Config.AppUserModelID, Config.CollectionID)
    If pNotifier = 0 Then Err.Raise 513, , "CreateToastNotifier failed."

    IIDFromString StrPtr("{354389C6-7C01-4BD5-9C20-604340CD2B74}"), iidNotifier2
    WinRT_CallComMethod pNotifier, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidNotifier2), WinRT_vbPtr, VarPtr(pNotifier2)
    If pNotifier2 = 0 Then Err.Raise 513, , "QueryInterface IToastNotifier2 failed."

    pData = WinRT_CreateNotificationData(Binding)
    If pData = 0 Then Err.Raise 513, , "CreateNotificationData failed."

    hStrTag = WinRT_CreateHString(Config.Tag)
    hStrGroup = WinRT_CreateHString(Config.Group)
    updateResult = 0
    WinRT_CallComMethod pNotifier2, VT_IToastNotifier2_UpdateWithTagAndGroup, vbLong, WinRT_vbPtr, pData, WinRT_vbPtr, hStrTag, WinRT_vbPtr, hStrGroup, WinRT_vbPtr, VarPtr(updateResult)

    WinRT_UpdateToastNotification = updateResult

Cleanup:
    If Err.Number <> 0 Then
        If Err.Source = "" Then Err.Raise Err.Number, "WinRT_UpdateToastNotification", Err.Description
        WinRT_UpdateToastNotification = Err.Number
    End If
    On Error Resume Next
    If hStrGroup <> 0 Then WindowsDeleteString hStrGroup
    If hStrTag <> 0 Then WindowsDeleteString hStrTag
    If pData <> 0 Then WinRT_CallComMethod pData, VT_RELEASE, vbLong
    If pNotifier2 <> 0 Then WinRT_CallComMethod pNotifier2, VT_RELEASE, vbLong
    If pNotifier <> 0 Then WinRT_CallComMethod pNotifier, VT_RELEASE, vbLong
    WinRT_FinalizeRo initialized
    On Error GoTo 0
End Function

Public Sub WinRT_Shutdown()
    If WinRT_RoInitialized Then
        RoUninitialize
        WinRT_RoInitialized = False
    End If
    WinRT_ToastDataSequenceNumber = 0
End Sub

Public Sub WinRT_ResetToastDataSequence()
    WinRT_ToastDataSequenceNumber = 0
End Sub

Private Function WinRT_NextToastDataSequence() As Long
    WinRT_ToastDataSequenceNumber = WinRT_ToastDataSequenceNumber + 1
    WinRT_NextToastDataSequence = WinRT_ToastDataSequenceNumber
End Function

Public Sub WinRT_RemoveToastNotification(ByRef Config As WinRT_ToastConfig)
    Dim initialized As Boolean
    Dim pHistory As LongPtr
    Dim hStrTag As LongPtr
    Dim hStrGroup As LongPtr
    Dim hStrAppId As LongPtr

    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized

    If Len(Config.AppUserModelID) = 0 Then Err.Raise 513, "WinRT_RemoveToastNotification", "AppUserModelID is required."
    If Len(Config.Tag) > 0 And Len(Config.Group) = 0 Then Err.Raise 513, "WinRT_RemoveToastNotification", "Group is required when Tag is specified."

    pHistory = WinRT_GetToastHistory(Config.CollectionID)
    If pHistory = 0 Then Err.Raise 513, , "GetToastHistory failed."

    hStrTag = WinRT_CreateHString(Config.Tag)
    hStrGroup = WinRT_CreateHString(Config.Group)
    hStrAppId = WinRT_CreateHString(Config.AppUserModelID)

    If Len(Config.Tag) > 0 Then
        WinRT_CallComMethod pHistory, VT_IToastNotificationHistory_RemoveGroupedTagWithId, vbLong, WinRT_vbPtr, hStrTag, WinRT_vbPtr, hStrGroup, WinRT_vbPtr, hStrAppId
    ElseIf Len(Config.Group) > 0 Then
        WinRT_CallComMethod pHistory, VT_IToastNotificationHistory_RemoveGroupWithId, vbLong, WinRT_vbPtr, hStrGroup, WinRT_vbPtr, hStrAppId
    Else
        WinRT_CallComMethod pHistory, VT_IToastNotificationHistory_ClearWithId, vbLong, WinRT_vbPtr, hStrAppId
    End If

Cleanup:
    If Err.Number <> 0 Then Err.Raise Err.Number, "WinRT_RemoveToastNotification", Err.Description
    On Error Resume Next
    If hStrAppId <> 0 Then WindowsDeleteString hStrAppId
    If hStrGroup <> 0 Then WindowsDeleteString hStrGroup
    If hStrTag <> 0 Then WindowsDeleteString hStrTag
    If pHistory <> 0 Then WinRT_CallComMethod pHistory, VT_RELEASE, vbLong
    WinRT_FinalizeRo initialized
    On Error GoTo 0
End Sub

Public Function WinRT_CheckNotificationSetting(ByRef Config As WinRT_ToastConfig) As Long
    Dim initialized As Boolean
    Dim pNotifier As LongPtr
    Dim settingValue As Long

    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized

    pNotifier = WinRT_CreateToastNotifier(Config.AppUserModelID, Config.CollectionID)
    If pNotifier = 0 Then Err.Raise 513, , "CreateToastNotifier failed."

    ' IToastNotifier.Setting は HRESULT get_Setting(NotificationSetting* out)。
    ' 出力ポインタを渡さないと out 引数がゴミ値になりアクセス違反でクラッシュする。
    settingValue = 0
    WinRT_CallComMethod pNotifier, VT_IToastNotifier_GetSetting, vbLong, WinRT_vbPtr, VarPtr(settingValue)
    WinRT_CheckNotificationSetting = settingValue

Cleanup:
    If Err.Number <> 0 Then WinRT_CheckNotificationSetting = Err.Number
    On Error Resume Next
    If pNotifier <> 0 Then WinRT_CallComMethod pNotifier, VT_RELEASE, vbLong
    WinRT_FinalizeRo initialized
    On Error GoTo 0
End Function

' CollectionNotice.cpp CreateToastCollection と同等（WinRT 直呼び）。0=成功
Public Function WinRT_SaveToastCollection(ByRef Config As WinRT_ToastConfig, ByVal DisplayName As String, ByVal LaunchArgs As String, ByVal IconUri As String) As Long
    Dim initialized As Boolean
    Dim pManager As LongPtr
    Dim pCollection As LongPtr
    Dim pAsync As LongPtr

    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized

    If Len(Config.AppUserModelID) = 0 Then Err.Raise 513, "WinRT_SaveToastCollection", "AppUserModelID is required."
    If Len(Config.CollectionID) = 0 Then Err.Raise 513, "WinRT_SaveToastCollection", "CollectionID is required."

    pManager = WinRT_GetToastCollectionManager(Config.AppUserModelID)
    If pManager = 0 Then Err.Raise 513, , "GetToastCollectionManager failed."

    pCollection = WinRT_CreateToastCollection(Config.CollectionID, DisplayName, LaunchArgs, IconUri)
    If pCollection = 0 Then Err.Raise 513, , "CreateToastCollection failed."

    pAsync = 0
    WinRT_CallComMethod pManager, VT_IToastCollectionManager_SaveToastCollectionAsync, vbLong, WinRT_vbPtr, pCollection, WinRT_vbPtr, VarPtr(pAsync)
    WinRT_SaveToastCollection = 0

Cleanup:
    If Err.Number <> 0 Then WinRT_SaveToastCollection = Err.Number
    On Error Resume Next
    If pAsync <> 0 Then WinRT_CallComMethod pAsync, VT_RELEASE, vbLong
    If pCollection <> 0 Then WinRT_CallComMethod pCollection, VT_RELEASE, vbLong
    If pManager <> 0 Then WinRT_CallComMethod pManager, VT_RELEASE, vbLong
    WinRT_FinalizeRo initialized
    On Error GoTo 0
End Function

' CollectionNotice.cpp DeleteToastCollection と同等。CollectionID 省略時は全削除。0=成功
Public Function WinRT_RemoveToastCollection(ByRef Config As WinRT_ToastConfig) As Long
    Dim initialized As Boolean
    Dim pManager As LongPtr
    Dim pAsync As LongPtr
    Dim hStrCollectionId As LongPtr

    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized

    If Len(Config.AppUserModelID) = 0 Then Err.Raise 513, "WinRT_RemoveToastCollection", "AppUserModelID is required."

    pManager = WinRT_GetToastCollectionManager(Config.AppUserModelID)
    If pManager = 0 Then Err.Raise 513, , "GetToastCollectionManager failed."

    pAsync = 0
    If Len(Config.CollectionID) > 0 Then
        hStrCollectionId = WinRT_CreateHString(Config.CollectionID)
        WinRT_CallComMethod pManager, VT_IToastCollectionManager_RemoveToastCollectionAsync, vbLong, WinRT_vbPtr, hStrCollectionId, WinRT_vbPtr, VarPtr(pAsync)
    Else
        WinRT_CallComMethod pManager, VT_IToastCollectionManager_RemoveAllToastCollectionsAsync, vbLong, WinRT_vbPtr, VarPtr(pAsync)
    End If
    WinRT_RemoveToastCollection = 0

Cleanup:
    If Err.Number <> 0 Then WinRT_RemoveToastCollection = Err.Number
    On Error Resume Next
    If hStrCollectionId <> 0 Then WindowsDeleteString hStrCollectionId
    If pAsync <> 0 Then WinRT_CallComMethod pAsync, VT_RELEASE, vbLong
    If pManager <> 0 Then WinRT_CallComMethod pManager, VT_RELEASE, vbLong
    WinRT_FinalizeRo initialized
    On Error GoTo 0
End Function

' ToastNotificationManager.GetDefault → IToastNotificationManagerForUser2.GetToastCollectionManager(appId)
Private Function WinRT_GetToastCollectionManager(ByVal appId As String) As LongPtr
    Dim pManagerStatics5 As LongPtr
    Dim pManagerForUser As LongPtr
    Dim pManagerForUser2 As LongPtr
    Dim pCollectionManager As LongPtr
    Dim hStrAppId As LongPtr
    Dim iidManagerStatics5 As WinRT_GUID
    Dim iidManagerForUser2 As WinRT_GUID

    IIDFromString StrPtr("{D6F5F569-D40D-407C-8989-88CAB42CFD14}"), iidManagerStatics5
    IIDFromString StrPtr("{679C64B7-81AB-42C2-8819-C958767753F4}"), iidManagerForUser2

    WinRT_GetActivationFactory "Windows.UI.Notifications.ToastNotificationManager", iidManagerStatics5, pManagerStatics5
    If pManagerStatics5 = 0 Then Exit Function

    WinRT_CallComMethod pManagerStatics5, VT_IToastNotificationManagerStatics5_GetDefault, vbLong, WinRT_vbPtr, VarPtr(pManagerForUser)
    WinRT_CallComMethod pManagerStatics5, VT_RELEASE, vbLong
    If pManagerForUser = 0 Then Exit Function

    WinRT_CallComMethod pManagerForUser, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidManagerForUser2), WinRT_vbPtr, VarPtr(pManagerForUser2)
    WinRT_CallComMethod pManagerForUser, VT_RELEASE, vbLong
    If pManagerForUser2 = 0 Then Exit Function

    hStrAppId = WinRT_CreateHString(appId)
    WinRT_CallComMethod pManagerForUser2, VT_IToastNotificationManagerForUser2_GetToastCollectionManagerWithAppId, vbLong, WinRT_vbPtr, hStrAppId, WinRT_vbPtr, VarPtr(pCollectionManager)
    If hStrAppId <> 0 Then WindowsDeleteString hStrAppId
    WinRT_CallComMethod pManagerForUser2, VT_RELEASE, vbLong

    WinRT_GetToastCollectionManager = pCollectionManager
End Function

' ToastCollection(collectionId, displayName, launchArgs, iconUri)
Private Function WinRT_CreateToastCollection(ByVal collectionId As String, ByVal displayName As String, ByVal launchArgs As String, ByVal iconUri As String) As LongPtr
    Dim pFactory As LongPtr
    Dim pUri As LongPtr
    Dim pCollection As LongPtr
    Dim hStrCollectionId As LongPtr
    Dim hStrDisplayName As LongPtr
    Dim hStrLaunchArgs As LongPtr
    Dim iidFactory As WinRT_GUID

    pUri = WinRT_CreateUri(iconUri)

    IIDFromString StrPtr("{164DD3D7-73C4-44F7-B4FF-FB6D4BF1F4C6}"), iidFactory
    WinRT_GetActivationFactory "Windows.UI.Notifications.ToastCollection", iidFactory, pFactory
    If pFactory <> 0 Then
        hStrCollectionId = WinRT_CreateHString(collectionId)
        hStrDisplayName = WinRT_CreateHString(displayName)
        hStrLaunchArgs = WinRT_CreateHString(launchArgs)

        WinRT_CallComMethod pFactory, VT_IToastCollectionFactory_CreateInstance, vbLong, _
            WinRT_vbPtr, hStrCollectionId, WinRT_vbPtr, hStrDisplayName, WinRT_vbPtr, hStrLaunchArgs, WinRT_vbPtr, pUri, WinRT_vbPtr, VarPtr(pCollection)
    End If

    If hStrCollectionId <> 0 Then WindowsDeleteString hStrCollectionId
    If hStrDisplayName <> 0 Then WindowsDeleteString hStrDisplayName
    If hStrLaunchArgs <> 0 Then WindowsDeleteString hStrLaunchArgs
    If pUri <> 0 Then WinRT_CallComMethod pUri, VT_RELEASE, vbLong
    If pFactory <> 0 Then WinRT_CallComMethod pFactory, VT_RELEASE, vbLong
    WinRT_CreateToastCollection = pCollection
End Function

' Windows.Foundation.Uri を生成（空文字なら 0）
Private Function WinRT_CreateUri(ByVal uri As String) As LongPtr
    Dim pFactory As LongPtr
    Dim pUri As LongPtr
    Dim hStrUri As LongPtr
    Dim iidFactory As WinRT_GUID

    If Len(uri) = 0 Then Exit Function

    IIDFromString StrPtr("{44A9796F-723E-4FDF-A218-033E75B0C084}"), iidFactory
    WinRT_GetActivationFactory "Windows.Foundation.Uri", iidFactory, pFactory
    If pFactory = 0 Then Exit Function

    hStrUri = WinRT_CreateHString(uri)
    WinRT_CallComMethod pFactory, VT_IUriFactory_CreateUri, vbLong, WinRT_vbPtr, hStrUri, WinRT_vbPtr, VarPtr(pUri)
    If hStrUri <> 0 Then WindowsDeleteString hStrUri
    WinRT_CallComMethod pFactory, VT_RELEASE, vbLong
    WinRT_CreateUri = pUri
End Function

Private Sub WinRT_ShowImmediateToast(ByRef Config As WinRT_ToastConfig, ByRef Binding As WinRT_DataBinding)
    Dim pXmlDoc As LongPtr
    Dim pToast As LongPtr
    Dim pToast2 As LongPtr
    Dim pToast4 As LongPtr
    Dim pToast6 As LongPtr
    Dim pNotifier As LongPtr
    Dim pData As LongPtr
    Dim hStrTag As LongPtr
    Dim hStrGroup As LongPtr
    Dim iidToast2 As WinRT_GUID
    Dim iidToast4 As WinRT_GUID
    Dim iidToast6 As WinRT_GUID

    pXmlDoc = WinRT_LoadXmlDocument(Config.XmlTemplate)
    If pXmlDoc = 0 Then Err.Raise 513, , "LoadXmlDocument failed."

    pToast = WinRT_CreateToastNotification(pXmlDoc)
    If pToast = 0 Then Err.Raise 513, , "CreateToastNotification failed."
    WinRT_CallComMethod pXmlDoc, VT_RELEASE, vbLong
    pXmlDoc = 0

    pData = WinRT_CreateNotificationData(Binding)
    If pData <> 0 Then
        IIDFromString StrPtr("{15154935-28EA-4727-88E9-C58680E2D118}"), iidToast4
        WinRT_CallComMethod pToast, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidToast4), WinRT_vbPtr, VarPtr(pToast4)
        If pToast4 <> 0 Then
            WinRT_CallComMethod pToast4, VT_IToastNotification4_SetData, vbLong, WinRT_vbPtr, pData
            WinRT_CallComMethod pToast4, VT_RELEASE, vbLong
            pToast4 = 0
        End If
        WinRT_CallComMethod pData, VT_RELEASE, vbLong
        pData = 0
    End If

    IIDFromString StrPtr("{9DFB9FD1-143A-490E-90BF-B9FBA7132DE7}"), iidToast2
    WinRT_CallComMethod pToast, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidToast2), WinRT_vbPtr, VarPtr(pToast2)
    If pToast2 <> 0 Then
        hStrTag = WinRT_CreateHString(Config.Tag)
        hStrGroup = WinRT_CreateHString(Config.Group)
        If hStrTag <> 0 Then WinRT_CallComMethod pToast2, VT_IToastNotification2_SetTag, vbLong, WinRT_vbPtr, hStrTag
        If hStrGroup <> 0 Then WinRT_CallComMethod pToast2, VT_IToastNotification2_SetGroup, vbLong, WinRT_vbPtr, hStrGroup
        If Config.SuppressPopup Then WinRT_CallComMethod pToast2, VT_IToastNotification2_SetSuppressPopup, vbLong, vbByte, CByte(1)
        WinRT_CallComMethod pToast2, VT_RELEASE, vbLong
        pToast2 = 0
    End If

    If Config.ExpiresOnReboot Then
        IIDFromString StrPtr("{43EBFE53-89AE-5C1E-A279-3AECFE9B6F54}"), iidToast6
        WinRT_CallComMethod pToast, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidToast6), WinRT_vbPtr, VarPtr(pToast6)
        If pToast6 <> 0 Then
            WinRT_CallComMethod pToast6, VT_IToastNotification6_SetExpiresOnReboot, vbLong, vbByte, CByte(1)
            WinRT_CallComMethod pToast6, VT_RELEASE, vbLong
            pToast6 = 0
        End If
    End If

    WinRT_ApplyExpirationTime pToast, VT_IToastNotification_SetExpirationTime, Config

    ' Activated / Dismissed / Failed を登録（DLL 版 GeneralNotice.cpp と同等）。
    ' トーストクリック・ボタン押下で launch / action arguments のマクロを実行する。
    WinRT_RegisterToastEventHandlers pToast

    pNotifier = WinRT_CreateToastNotifier(Config.AppUserModelID, Config.CollectionID)
    If pNotifier = 0 Then Err.Raise 513, , "CreateToastNotifier failed."
    WinRT_CallComMethod pNotifier, VT_IToastNotifier_Show, vbLong, WinRT_vbPtr, pToast

    If hStrTag <> 0 Then WindowsDeleteString hStrTag
    If hStrGroup <> 0 Then WindowsDeleteString hStrGroup
    If pNotifier <> 0 Then WinRT_CallComMethod pNotifier, VT_RELEASE, vbLong
    If pToast <> 0 Then WinRT_CallComMethod pToast, VT_RELEASE, vbLong
End Sub

Private Sub WinRT_ValidateToastIdentity(ByRef Config As WinRT_ToastConfig, ByVal procName As String)
    If Len(Config.Tag) = 0 Then Err.Raise 513, procName, "Tag is required."
    If Len(Config.Group) = 0 Then Err.Raise 513, procName, "Group is required."
    If Len(Config.AppUserModelID) = 0 Then Err.Raise 513, procName, "AppUserModelID is required."
End Sub

Private Sub WinRT_ShowScheduledToast(ByRef Config As WinRT_ToastConfig)
    Dim pXmlDoc As LongPtr
    Dim pScheduledFactory As LongPtr
    Dim pScheduled As LongPtr
    Dim pScheduled2 As LongPtr
    Dim pScheduled4 As LongPtr
    Dim pNotifier As LongPtr
    Dim hStrScheduleId As LongPtr
    Dim hStrTag As LongPtr
    Dim hStrGroup As LongPtr
    Dim deliveryTime As WRT_DateTime
    Dim iidScheduledFactory As WinRT_GUID
    Dim iidScheduled2 As WinRT_GUID
    Dim iidScheduled4 As WinRT_GUID

    pXmlDoc = WinRT_LoadXmlDocument(Config.XmlTemplate)
    If pXmlDoc = 0 Then Err.Raise 513, , "LoadXmlDocument failed."

    IIDFromString StrPtr("{E7BED191-0BB9-4189-8394-31761B476FD7}"), iidScheduledFactory
    WinRT_GetActivationFactory "Windows.UI.Notifications.ScheduledToastNotification", iidScheduledFactory, pScheduledFactory
    If pScheduledFactory = 0 Then Err.Raise 513, , "ScheduledToastNotification factory failed."

    WinRT_ScheduleDateToWinRTDateTimeFillFromConfig Config, deliveryTime
    pScheduled = 0
    WinRT_FactoryCreateScheduledToast pScheduledFactory, pXmlDoc, deliveryTime, pScheduled
    WinRT_CallComMethod pScheduledFactory, VT_RELEASE, vbLong
    WinRT_CallComMethod pXmlDoc, VT_RELEASE, vbLong
    pScheduledFactory = 0
    pXmlDoc = 0
    If pScheduled = 0 Then Err.Raise 513, , "CreateScheduledToastNotification failed."

    IIDFromString StrPtr("{A66EA09C-31B4-43B0-B5DD-7A40E85363B1}"), iidScheduled2
    WinRT_CallComMethod pScheduled, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidScheduled2), WinRT_vbPtr, VarPtr(pScheduled2)
    If pScheduled2 <> 0 Then
        If Len(Config.Schedule_ID) > 0 Then
            hStrScheduleId = WinRT_CreateHString(WinRT_NormalizeScheduleId(Config.Schedule_ID))
        End If
        hStrTag = WinRT_CreateHString(Config.Tag)
        hStrGroup = WinRT_CreateHString(Config.Group)
        If hStrScheduleId <> 0 Then WinRT_CallComMethod pScheduled, VT_IScheduledToastNotification_SetId, vbLong, WinRT_vbPtr, hStrScheduleId
        If hStrTag <> 0 Then WinRT_CallComMethod pScheduled2, VT_IToastNotification2_SetTag, vbLong, WinRT_vbPtr, hStrTag
        If hStrGroup <> 0 Then WinRT_CallComMethod pScheduled2, VT_IToastNotification2_SetGroup, vbLong, WinRT_vbPtr, hStrGroup
        If Config.SuppressPopup Then WinRT_CallComMethod pScheduled2, VT_IToastNotification2_SetSuppressPopup, vbLong, vbByte, CByte(1)
        WinRT_CallComMethod pScheduled2, VT_RELEASE, vbLong
        pScheduled2 = 0
    End If

    If Config.ExpirationTime > 0 Then
        ' IScheduledToastNotification4（1D4761FD...）。98429E8B... は v3 で ExpirationTime は存在しない
        IIDFromString StrPtr("{1D4761FD-BDEF-4E4A-96BE-0101369B58D2}"), iidScheduled4
        WinRT_CallComMethod pScheduled, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidScheduled4), WinRT_vbPtr, VarPtr(pScheduled4)
        If pScheduled4 <> 0 Then
            WinRT_ApplyExpirationTime pScheduled4, VT_IScheduledToastNotification4_SetExpirationTime, Config
            WinRT_CallComMethod pScheduled4, VT_RELEASE, vbLong
            pScheduled4 = 0
        End If
    End If

    pNotifier = WinRT_CreateToastNotifier(Config.AppUserModelID, Config.CollectionID)
    If pNotifier = 0 Then Err.Raise 513, , "CreateToastNotifier failed."
    WinRT_CallComMethod pNotifier, VT_IToastNotifier_AddToSchedule, vbLong, WinRT_vbPtr, pScheduled

    If hStrScheduleId <> 0 Then WindowsDeleteString hStrScheduleId
    If hStrTag <> 0 Then WindowsDeleteString hStrTag
    If hStrGroup <> 0 Then WindowsDeleteString hStrGroup
    If pNotifier <> 0 Then WinRT_CallComMethod pNotifier, VT_RELEASE, vbLong
    If pScheduled <> 0 Then WinRT_CallComMethod pScheduled, VT_RELEASE, vbLong
End Sub

Private Function WinRT_CreateToastNotifier(ByVal AppUserModelID As String, ByVal CollectionID As String) As LongPtr
    If Len(CollectionID) > 0 Then
        WinRT_CreateToastNotifier = WinRT_CreateToastNotifierForCollection(CollectionID)
    Else
        WinRT_CreateToastNotifier = WinRT_CreateToastNotifierWithAppId(AppUserModelID)
    End If
End Function

Private Function WinRT_CreateToastNotifierWithAppId(ByVal AppUserModelID As String) As LongPtr
    Dim hStrManagerClass As LongPtr
    Dim hStrAppId As LongPtr
    Dim pManagerStatics As LongPtr
    Dim pNotifier As LongPtr
    Dim iidManagerStatics As WinRT_GUID

    IIDFromString StrPtr("{50AC103F-D235-4598-BBEF-98FE4D1A3AD4}"), iidManagerStatics
    hStrManagerClass = WinRT_CreateHString("Windows.UI.Notifications.ToastNotificationManager")
    RoGetActivationFactory hStrManagerClass, iidManagerStatics, pManagerStatics
    If hStrManagerClass <> 0 Then WindowsDeleteString hStrManagerClass

    If pManagerStatics = 0 Then Exit Function
    hStrAppId = WinRT_CreateHString(AppUserModelID)
    WinRT_CallComMethod pManagerStatics, 7, vbLong, WinRT_vbPtr, hStrAppId, WinRT_vbPtr, VarPtr(pNotifier)
    If hStrAppId <> 0 Then WindowsDeleteString hStrAppId
    WinRT_CallComMethod pManagerStatics, VT_RELEASE, vbLong
    WinRT_CreateToastNotifierWithAppId = pNotifier
End Function

Private Function WinRT_CreateToastNotifierForCollection(ByVal CollectionID As String) As LongPtr
    Dim hStrCollectionId As LongPtr
    Dim pManagerStatics5 As LongPtr
    Dim pManagerForUser As LongPtr
    Dim pManagerForUser2 As LongPtr
    Dim pAsync As LongPtr
    Dim pNotifier As LongPtr
    Dim hr As Long
    Dim iidManagerStatics5 As WinRT_GUID
    Dim iidManagerForUser2 As WinRT_GUID

    IIDFromString StrPtr("{D6F5F569-D40D-407C-8989-88CAB42CFD14}"), iidManagerStatics5
    IIDFromString StrPtr("{679C64B7-81AB-42C2-8819-C958767753F4}"), iidManagerForUser2

    WinRT_GetActivationFactory "Windows.UI.Notifications.ToastNotificationManager", iidManagerStatics5, pManagerStatics5
    If pManagerStatics5 = 0 Then Exit Function

    WinRT_CallComMethod pManagerStatics5, 6, vbLong, WinRT_vbPtr, VarPtr(pManagerForUser)
    WinRT_CallComMethod pManagerStatics5, VT_RELEASE, vbLong
    If pManagerForUser = 0 Then Exit Function

    WinRT_CallComMethod pManagerForUser, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidManagerForUser2), WinRT_vbPtr, VarPtr(pManagerForUser2)
    WinRT_CallComMethod pManagerForUser, VT_RELEASE, vbLong
    If pManagerForUser2 = 0 Then Exit Function

    hStrCollectionId = WinRT_CreateHString(CollectionID)
    WinRT_CallComMethod pManagerForUser2, 6, vbLong, WinRT_vbPtr, hStrCollectionId, WinRT_vbPtr, VarPtr(pAsync)
    If hStrCollectionId <> 0 Then WindowsDeleteString hStrCollectionId
    WinRT_CallComMethod pManagerForUser2, VT_RELEASE, vbLong

    If pAsync <> 0 Then
        hr = WinRT_WaitAsync(pAsync)
        If hr = 0 Then
            pNotifier = 0
            WinRT_CallComMethod pAsync, VT_IAsyncOperation_GetResults, vbLong, WinRT_vbPtr, VarPtr(pNotifier)
        End If
        WinRT_CallComMethod pAsync, VT_RELEASE, vbLong
    End If
    WinRT_CreateToastNotifierForCollection = pNotifier
End Function

Private Function WinRT_GetToastHistory(ByVal CollectionID As String) As LongPtr
    If Len(CollectionID) > 0 Then
        WinRT_GetToastHistory = WinRT_GetHistoryForCollection(CollectionID)
    Else
        WinRT_GetToastHistory = WinRT_GetDefaultHistory()
    End If
End Function

Private Function WinRT_GetDefaultHistory() As LongPtr
    Dim pManagerStatics As LongPtr
    Dim pManagerStatics2 As LongPtr
    Dim pHistory As LongPtr
    Dim iidManagerStatics As WinRT_GUID
    Dim iidManagerStatics2 As WinRT_GUID

    IIDFromString StrPtr("{50AC103F-D235-4598-BBEF-98FE4D1A3AD4}"), iidManagerStatics
    IIDFromString StrPtr("{7AB93C52-0E48-4750-BA9D-1A4113981847}"), iidManagerStatics2

    WinRT_GetActivationFactory "Windows.UI.Notifications.ToastNotificationManager", iidManagerStatics, pManagerStatics
    If pManagerStatics = 0 Then Exit Function

    WinRT_CallComMethod pManagerStatics, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidManagerStatics2), WinRT_vbPtr, VarPtr(pManagerStatics2)
    WinRT_CallComMethod pManagerStatics, VT_RELEASE, vbLong
    If pManagerStatics2 = 0 Then Exit Function

    pHistory = 0
    WinRT_CallComMethod pManagerStatics2, VT_IToastNotificationManagerStatics2_GetHistory, vbLong, WinRT_vbPtr, VarPtr(pHistory)
    WinRT_CallComMethod pManagerStatics2, VT_RELEASE, vbLong
    WinRT_GetDefaultHistory = pHistory
End Function

Private Function WinRT_GetHistoryForCollection(ByVal CollectionID As String) As LongPtr
    Dim hStrCollectionId As LongPtr
    Dim pManagerStatics5 As LongPtr
    Dim pManagerForUser As LongPtr
    Dim pManagerForUser2 As LongPtr
    Dim pAsync As LongPtr
    Dim pHistory As LongPtr
    Dim hr As Long
    Dim iidManagerStatics5 As WinRT_GUID
    Dim iidManagerForUser2 As WinRT_GUID

    IIDFromString StrPtr("{D6F5F569-D40D-407C-8989-88CAB42CFD14}"), iidManagerStatics5
    IIDFromString StrPtr("{679C64B7-81AB-42C2-8819-C958767753F4}"), iidManagerForUser2

    WinRT_GetActivationFactory "Windows.UI.Notifications.ToastNotificationManager", iidManagerStatics5, pManagerStatics5
    If pManagerStatics5 = 0 Then Exit Function

    WinRT_CallComMethod pManagerStatics5, 6, vbLong, WinRT_vbPtr, VarPtr(pManagerForUser)
    WinRT_CallComMethod pManagerStatics5, VT_RELEASE, vbLong
    If pManagerForUser = 0 Then Exit Function

    WinRT_CallComMethod pManagerForUser, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidManagerForUser2), WinRT_vbPtr, VarPtr(pManagerForUser2)
    WinRT_CallComMethod pManagerForUser, VT_RELEASE, vbLong
    If pManagerForUser2 = 0 Then Exit Function

    hStrCollectionId = WinRT_CreateHString(CollectionID)
    WinRT_CallComMethod pManagerForUser2, 7, vbLong, WinRT_vbPtr, hStrCollectionId, WinRT_vbPtr, VarPtr(pAsync)
    If hStrCollectionId <> 0 Then WindowsDeleteString hStrCollectionId
    WinRT_CallComMethod pManagerForUser2, VT_RELEASE, vbLong

    If pAsync <> 0 Then
        hr = WinRT_WaitAsync(pAsync)
        If hr = 0 Then
            pHistory = 0
            WinRT_CallComMethod pAsync, VT_IAsyncOperation_GetResults, vbLong, WinRT_vbPtr, VarPtr(pHistory)
        End If
        WinRT_CallComMethod pAsync, VT_RELEASE, vbLong
    End If
    WinRT_GetHistoryForCollection = pHistory
End Function

Private Function WinRT_LoadXmlDocument(ByVal xml As String) As LongPtr
    Dim hStrXmlClass As LongPtr
    Dim hStrXml As LongPtr
    Dim pInspectable As LongPtr
    Dim pXmlDocIO As LongPtr
    Dim pXmlDoc As LongPtr
    Dim iidXmlDocIO As WinRT_GUID
    Dim iidXmlDoc As WinRT_GUID

    IIDFromString StrPtr("{6CD0E74E-EE65-4489-9EBF-CA43E87BA637}"), iidXmlDocIO
    IIDFromString StrPtr("{F7F3A506-1E87-42D6-BCFB-B8C809FA5494}"), iidXmlDoc

    hStrXmlClass = WinRT_CreateHString("Windows.Data.Xml.Dom.XmlDocument")
    RoActivateInstance hStrXmlClass, pInspectable
    If hStrXmlClass <> 0 Then WindowsDeleteString hStrXmlClass
    If pInspectable = 0 Then Exit Function

    WinRT_CallComMethod pInspectable, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidXmlDocIO), WinRT_vbPtr, VarPtr(pXmlDocIO)
    hStrXml = WinRT_CreateHString(xml)
    WinRT_CallComMethod pXmlDocIO, 6, vbLong, WinRT_vbPtr, hStrXml
    If hStrXml <> 0 Then WindowsDeleteString hStrXml
    WinRT_CallComMethod pXmlDocIO, VT_RELEASE, vbLong
    pXmlDocIO = 0

    WinRT_CallComMethod pInspectable, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidXmlDoc), WinRT_vbPtr, VarPtr(pXmlDoc)
    WinRT_CallComMethod pInspectable, VT_RELEASE, vbLong
    WinRT_LoadXmlDocument = pXmlDoc
End Function

Private Function WinRT_CreateToastNotification(ByVal pXmlDoc As LongPtr) As LongPtr
    Dim hStrToastClass As LongPtr
    Dim pToastFactory As LongPtr
    Dim pToast As LongPtr
    Dim iidToastFactory As WinRT_GUID

    IIDFromString StrPtr("{04124B20-82C6-4229-B109-FD9ED4662B53}"), iidToastFactory
    hStrToastClass = WinRT_CreateHString("Windows.UI.Notifications.ToastNotification")
    RoGetActivationFactory hStrToastClass, iidToastFactory, pToastFactory
    If hStrToastClass <> 0 Then WindowsDeleteString hStrToastClass
    If pToastFactory = 0 Then Exit Function

    WinRT_CallComMethod pToastFactory, 6, vbLong, WinRT_vbPtr, pXmlDoc, WinRT_vbPtr, VarPtr(pToast)
    WinRT_CallComMethod pToastFactory, VT_RELEASE, vbLong
    WinRT_CreateToastNotification = pToast
End Function

Private Function WinRT_CreateNotificationData(ByRef Binding As WinRT_DataBinding) As LongPtr
    Dim hStrClass As LongPtr
    Dim pInspectable As LongPtr
    Dim pData As LongPtr
    Dim pMap As LongPtr
    Dim iidData As WinRT_GUID
    Dim hasProgress As Boolean

    If Not WinRT_DataBindingHasContent(Binding) Then Exit Function

    IIDFromString StrPtr("{9FFD2312-9D6A-4AAF-B6AC-FF17F0C1F280}"), iidData

    hStrClass = WinRT_CreateHString("Windows.UI.Notifications.NotificationData")
    RoActivateInstance hStrClass, pInspectable
    If hStrClass <> 0 Then WindowsDeleteString hStrClass
    If pInspectable = 0 Then Exit Function

    WinRT_CallComMethod pInspectable, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidData), WinRT_vbPtr, VarPtr(pData)
    WinRT_CallComMethod pInspectable, VT_RELEASE, vbLong
    If pData = 0 Then Exit Function

    If Binding.SequenceNumber > 0 Then
        WinRT_CallComMethod pData, VT_INotificationData_SetSequenceNumber, vbLong, vbLong, Binding.SequenceNumber
    End If

    hasProgress = (Len(Binding.ProgressStatus) > 0)
    WinRT_CallComMethod pData, 6, vbLong, WinRT_vbPtr, VarPtr(pMap)
    If pMap = 0 Then
        WinRT_CallComMethod pData, VT_RELEASE, vbLong
        Exit Function
    End If

    If Len(Binding.TitleText) > 0 Then WinRT_InsertMapValue pMap, "TopTextTitle", Binding.TitleText
    If Len(Binding.ContentsText) > 0 Then WinRT_InsertMapValue pMap, "TopTextContents", Binding.ContentsText
    If Len(Binding.AttributionText) > 0 Then WinRT_InsertMapValue pMap, "TopTextAttribution", Binding.AttributionText
    If hasProgress Then
        If Len(Binding.ProgressTitle) > 0 Then WinRT_InsertMapValue pMap, "ProgressTitle", Binding.ProgressTitle
        WinRT_InsertMapValue pMap, "ProgressStatus", Binding.ProgressStatus
        If Binding.progressValue < 0 Then
            WinRT_InsertMapValue pMap, "ProgressValue", "Indeterminate"
        Else
            WinRT_InsertMapValue pMap, "ProgressValue", WinRT_FormatProgressValue(Binding.progressValue)
        End If
        If Len(Binding.ProgressValueStringOverride) > 0 Then
            WinRT_InsertMapValue pMap, "ProgressValueString", Binding.ProgressValueStringOverride
        Else
            WinRT_InsertMapValue pMap, "ProgressValueString", WinRT_FormatProgressValueString(Binding.progressValue)
        End If
    End If
    WinRT_CallComMethod pMap, VT_RELEASE, vbLong

    WinRT_CreateNotificationData = pData
End Function

Private Function WinRT_FormatProgressValue(ByVal progressValue As Double) As String
    WinRT_FormatProgressValue = Replace(CStr(progressValue), ",", ".")
End Function

Private Function WinRT_FormatProgressValueString(ByVal progressValue As Double) As String
    If progressValue < 0 Then
        WinRT_FormatProgressValueString = "処理中"
    Else
        WinRT_FormatProgressValueString = Format$(progressValue, "0%")
    End If
End Function

Private Function WinRT_CoerceLongResult(ByVal vResult As Variant) As Long
    Select Case VarType(vResult)
        Case vbLong, vbInteger, vbByte
            WinRT_CoerceLongResult = CLng(vResult)
        Case vbLongLong
            WinRT_CoerceLongResult = CLng(vResult)
        Case Else
            WinRT_CoerceLongResult = 0
    End Select
End Function

Private Function WinRT_DataBindingHasContent(ByRef Binding As WinRT_DataBinding) As Boolean
    If Len(Binding.TitleText) > 0 Then WinRT_DataBindingHasContent = True: Exit Function
    If Len(Binding.ContentsText) > 0 Then WinRT_DataBindingHasContent = True: Exit Function
    If Len(Binding.AttributionText) > 0 Then WinRT_DataBindingHasContent = True: Exit Function
    If Len(Binding.ProgressStatus) > 0 Then WinRT_DataBindingHasContent = True
End Function

Private Function WinRT_CreateDateTimeReferenceFromConfig(ByRef Config As WinRT_ToastConfig) As LongPtr
    Dim dateTimeValue As WRT_DateTime

    ' DLL (GeneralNotice.cpp) と同一: UTC 補正済みシリアル → VariantTimeToSystemTime → SystemTimeToFileTime。
    If Config.ExpirationTime <= 0 Then Exit Function
    WinRT_DateSerialToWinRTDateTimeFill Config.ExpirationTime, dateTimeValue
    If WRT_DateTimeIsZero(dateTimeValue) Then Exit Function

    WinRT_CreateDateTimeReferenceFromConfig = WinRT_CreateDateTimeReferenceFromDateTime(dateTimeValue)
End Function

Private Sub WinRT_ApplyExpirationTime(ByVal pNotification As LongPtr, ByVal vTableIndex As Long, ByRef Config As WinRT_ToastConfig)
    Dim pExpirationRef As LongPtr

    If pNotification = 0 Then Exit Sub
    If Config.ExpirationTime <= 0 Then Exit Sub

    pExpirationRef = WinRT_CreateDateTimeReferenceFromConfig(Config)
    If pExpirationRef = 0 Then
        Err.Raise 513, "WinRT_ApplyExpirationTime", "PropertyValue.CreateDateTime failed for ExpirationTime."
    End If

    WinRT_CallComMethod pNotification, vTableIndex, vbLong, WinRT_vbPtr, pExpirationRef
    WinRT_CallComMethod pExpirationRef, VT_RELEASE, vbLong
End Sub

Private Function WinRT_CreateDateTimeReferenceFromDateTime(ByRef dateTimeValue As WRT_DateTime) As LongPtr
    Dim hStrClass As LongPtr
    Dim pPropertyValueStatics As LongPtr
    Dim pPropertyValue As LongPtr
    Dim pReference As LongPtr
    Dim iidPropertyValueStatics As WinRT_GUID
    Dim iidReferenceDateTime As WinRT_GUID

    IIDFromString StrPtr("{629BDBC8-D932-4FF4-96B9-8D96C5C1E858}"), iidPropertyValueStatics

    hStrClass = WinRT_CreateHString("Windows.Foundation.PropertyValue")
    RoGetActivationFactory hStrClass, iidPropertyValueStatics, pPropertyValueStatics
    If hStrClass <> 0 Then WindowsDeleteString hStrClass
    If pPropertyValueStatics = 0 Then Exit Function

    WinRT_PropertyValueCreateDateTime pPropertyValueStatics, dateTimeValue, pPropertyValue
    WinRT_CallComMethod pPropertyValueStatics, VT_RELEASE, vbLong
    If pPropertyValue = 0 Then Exit Function

    ' CreateDateTime は IPropertyValue を返すが、ExpirationTime は IReference<DateTime> を要求する。
    ' 別 vtable なので必ず QueryInterface してから渡す（直接渡すと値が壊れる）。
    IIDFromString StrPtr("{5541D8A7-497C-5AA4-86FC-7713ADBF2A2C}"), iidReferenceDateTime
    WinRT_CallComMethod pPropertyValue, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidReferenceDateTime), WinRT_vbPtr, VarPtr(pReference)
    WinRT_CallComMethod pPropertyValue, VT_RELEASE, vbLong
    If pReference = 0 Then Exit Function

    WinRT_CreateDateTimeReferenceFromDateTime = pReference
End Function

Private Sub WinRT_InsertMapValue(ByVal pMap As LongPtr, ByVal key As String, ByVal val As String)
    Dim hKey As LongPtr
    Dim hVal As LongPtr
    Dim replaced As Byte
    hKey = WinRT_CreateHString(key)
    hVal = WinRT_CreateHString(val)
    WinRT_CallComMethod pMap, 10, vbLong, WinRT_vbPtr, hKey, WinRT_vbPtr, hVal, WinRT_vbPtr, VarPtr(replaced)
    WindowsDeleteString hKey
    WindowsDeleteString hVal
End Sub

Private Sub WinRT_GetActivationFactory(ByVal className As String, ByRef iid As WinRT_GUID, ByRef factory As LongPtr)
    Dim hStrClass As LongPtr
    hStrClass = WinRT_CreateHString(className)
    RoGetActivationFactory hStrClass, iid, factory
    If hStrClass <> 0 Then WindowsDeleteString hStrClass
End Sub

Private Function WinRT_CreateHString(ByVal s As String) As LongPtr
    Dim hStr As LongPtr
    Dim hr As Long
    Dim length As Long

    length = Len(s)
    If length > 0 Then
        hr = WindowsCreateString(StrPtr(s), length, hStr)
    Else
        hr = WindowsCreateString(StrPtr(""), 0, hStr)
    End If
    If hr < 0 Or hStr = 0 Then Err.Raise 513, "WinRT_CreateHString", "WindowsCreateString failed: 0x" & Hex$(hr)
    WinRT_CreateHString = hStr
End Function


Private Function WRT_DateTimeIsZero(ByRef dt As WRT_DateTime) As Boolean
    WRT_DateTimeIsZero = (dt.dwLowDateTime = 0 And dt.dwHighDateTime = 0)
End Function

Private Sub WRT_CopyFileTimeToDateTime(ByRef ft As WinRT_FILETIME, ByRef outDt As WRT_DateTime)
    outDt.dwLowDateTime = ft.dwLowDateTime
    outDt.dwHighDateTime = ft.dwHighDateTime
End Sub

Private Function WRT_DateTimeToLongLong(ByRef dt As WRT_DateTime) As LongLong
    Dim v As LongLong
    RtlMoveMemory v, dt.dwLowDateTime, 8
    WRT_DateTimeToLongLong = v
End Function

Private Function WRT_DateTimeToCurrency(ByRef dt As WRT_DateTime) As Currency
    Dim v As Currency
    RtlMoveMemory v, dt.dwLowDateTime, 8
    WRT_DateTimeToCurrency = v
End Function

Private Function WinRT_ResolveScheduleDeliveryLocal(ByRef Config As WinRT_ToastConfig) As Date
    If Config.Schedule_DeliveryTimeLocal > 0 Then
        WinRT_ResolveScheduleDeliveryLocal = Config.Schedule_DeliveryTimeLocal
    ElseIf Config.Schedule_DeliveryTime > 0 Then
        WinRT_ResolveScheduleDeliveryLocal = Config.Schedule_DeliveryTime
    End If
End Function

Private Sub WinRT_ScheduleDateToWinRTDateTimeFillFromConfig(ByRef Config As WinRT_ToastConfig, ByRef outDt As WRT_DateTime)
    Dim deliveryLocal As Date
    Dim minutes As Long

    outDt.dwLowDateTime = 0
    outDt.dwHighDateTime = 0
    deliveryLocal = WinRT_ResolveScheduleDeliveryLocal(Config)

    If deliveryLocal <= 0 Then
        Err.Raise 5, "WinRT_ScheduleDateToWinRTDateTimeFillFromConfig", _
            "Schedule delivery time is empty. Re-import AppNotificationBuilder.cls and modWinRTToast.bas."
    End If

#If WRT_USE_TLB_SCHEDULE Then
    WinRT_ScheduleDateToWinRTDateTimeFillViaTlb deliveryLocal, outDt
#Else
    ' 指定された絶対時刻（ローカル wall clock）を秒精度でそのまま UTC FILETIME へ変換する（誤差なし）
    WinRT_LocalSerialToWinRTDateTimeFill deliveryLocal, outDt
    ' TZ 変換に失敗した場合のみ「現在 + 分差」でフォールバック（秒は丸められる）
    If WRT_DateTimeIsZero(outDt) Then
        minutes = DateDiff("n", Now, deliveryLocal)
        If minutes < 1 Then
            Err.Raise 5, "WinRT_ScheduleDateToWinRTDateTimeFillFromConfig", "Schedule delivery time must be in the future."
        End If
        WinRT_DateTimeFromMinutesFromNow minutes, outDt
    End If
#End If

    If WRT_DateTimeIsZero(outDt) Then
        Err.Raise 5, "WinRT_ScheduleDateToWinRTDateTimeFillFromConfig", _
            "DeliveryTime is zero. local=" & Format$(deliveryLocal, "yyyy/mm/dd hh:nn:ss") & _
            " UTC serial=" & CStr(CDbl(Config.Schedule_DeliveryTime)) & _
            " local serial=" & CStr(CDbl(Config.Schedule_DeliveryTimeLocal)) & _
            " ft=0x" & Hex$(outDt.dwHighDateTime) & Right$("00000000" & Hex$(outDt.dwLowDateTime), 8)
    End If
End Sub

' TLB の SetToNow + AddMinutes と同じ考え方（UTC FILETIME + 分差）
Private Sub WinRT_DateTimeFromMinutesFromNow(ByVal minutesFromNow As Long, ByRef outDt As WRT_DateTime)
    Dim stUtc As WinRT_SYSTEMTIME
    Dim ftUtc As WinRT_FILETIME
    Dim ftAdd As WinRT_FILETIME
    Dim ticks As LongLong
    Dim addTicks As LongLong

    If minutesFromNow < 1 Then Exit Sub

    GetSystemTime stUtc
    If SystemTimeToFileTime(stUtc, ftUtc) = 0 Then Exit Sub

    RtlMoveMemory ticks, ftUtc.dwLowDateTime, 8
    addTicks = CLngLng(minutesFromNow) * WRT_FileTimeTicksPerMinute
    ticks = ticks + addTicks
    RtlMoveMemory ftAdd.dwLowDateTime, ticks, 8
    WRT_CopyFileTimeToDateTime ftAdd, outDt
End Sub

Private Sub WinRT_ScheduleDateToWinRTDateTimeFill(ByVal deliverySerial As Date, ByRef outDt As WRT_DateTime)
    outDt.dwLowDateTime = 0
    outDt.dwHighDateTime = 0

    If deliverySerial <= 0 Then Err.Raise 5, "WinRT_ScheduleDateToWinRTDateTimeFill", "deliverySerial is empty."

#If WRT_USE_TLB_SCHEDULE Then
    WinRT_ScheduleDateToWinRTDateTimeFillViaTlb deliverySerial, outDt
#Else
    WinRT_DateTimeFromMinutesFromNow DateDiff("n", Now, deliverySerial), outDt
    If WRT_DateTimeIsZero(outDt) Then WinRT_DateSerialToWinRTDateTimeFill deliverySerial, outDt
#End If

    If WRT_DateTimeIsZero(outDt) Then
        Err.Raise 5, "WinRT_ScheduleDateToWinRTDateTimeFill", _
            "DeliveryTime is zero. deliverySerial=" & CStr(CDbl(deliverySerial))
    End If
End Sub

#If WRT_USE_TLB_SCHEDULE Then
Private Sub WinRT_ScheduleDateToWinRTDateTimeFillViaTlb(ByVal deliverySerial As Date, ByRef outDt As WRT_DateTime)
    Dim cal As WinRT.ICalendar
    Dim dt As WinRT.DateTime
    Dim hClass As LongPtr
    Dim hr As Long
    Dim minutes As Long

    minutes = DateDiff("n", Now, deliverySerial)
    If minutes < 1 Then Err.Raise 5, "WinRT_ScheduleDateToWinRTDateTimeFillViaTlb", "Schedule delivery time must be in the future."

    hClass = WinRT_CreateHString("Windows.Globalization.Calendar")
    hr = RoActivateInstanceObj(hClass, cal)
    If hClass <> 0 Then WindowsDeleteString hClass
    If hr < 0 Then Err.Raise hr, "WinRT_ScheduleDateToWinRTDateTimeFillViaTlb", "RoActivateInstance Calendar failed: 0x" & Hex$(hr)

    hr = cal.SetToNow()
    If hr < 0 Then Err.Raise hr, "WinRT_ScheduleDateToWinRTDateTimeFillViaTlb", "ICalendar.SetToNow failed: 0x" & Hex$(hr)

    hr = cal.AddMinutes(minutes)
    If hr < 0 Then Err.Raise hr, "WinRT_ScheduleDateToWinRTDateTimeFillViaTlb", "ICalendar.AddMinutes failed: 0x" & Hex$(hr)

    dt = cal.GetDateTime()
    RtlMoveMemory outDt.dwLowDateTime, dt, 8
    Set cal = Nothing
End Sub
#End If

' GeneralNotice.cpp SystemTimeToDateTime と同じ（VariantTimeToSystemTime → SystemTimeToFileTime）
' ExpirationTime など UTC 補正済み serial 向け。スケジュール配信時刻は使わない。
Private Sub WinRT_DateSerialToWinRTDateTimeFill(ByVal deliverySerial As Date, ByRef outDt As WRT_DateTime)
    Dim st As WinRT_SYSTEMTIME
    Dim ft As WinRT_FILETIME

    If deliverySerial <= 0 Then Exit Sub

    VariantTimeToSystemTime CDbl(deliverySerial), st
    If SystemTimeToFileTime(st, ft) = 0 Then Exit Sub
    WRT_CopyFileTimeToDateTime ft, outDt
End Sub

' ローカル wall clock → UTC FILETIME（Calendar.Set 系の絶対日時向け）
Private Sub WinRT_LocalSerialToWinRTDateTimeFill(ByVal deliveryLocal As Date, ByRef outDt As WRT_DateTime)
    Dim stLocal As WinRT_SYSTEMTIME
    Dim stUtc As WinRT_SYSTEMTIME
    Dim ft As WinRT_FILETIME

    If deliveryLocal <= 0 Then Exit Sub

    VariantTimeToSystemTime CDbl(deliveryLocal), stLocal
    stLocal.wMilliseconds = 0
    If TzSpecificLocalTimeToSystemTime(0&, stLocal, stUtc) = 0 Then Exit Sub
    If SystemTimeToFileTime(stUtc, ft) = 0 Then Exit Sub
    WRT_CopyFileTimeToDateTime ft, outDt
End Sub

' CreateToastNotification と同じ WinRT_CallComMethod 経路（DispCallFunc 直呼びは使わない）
Private Sub WinRT_FactoryCreateScheduledToast( _
    ByVal pFactory As LongPtr, _
    ByVal pXmlDoc As LongPtr, _
    ByRef deliveryTime As WRT_DateTime, _
    ByRef pScheduled As LongPtr)

    Dim deliveryCy As Currency
    Dim deliveryTicks As LongLong
    Dim errNum As Long
    Dim errDesc As String
    Dim rawHex As String

    If pFactory = 0 Or pXmlDoc = 0 Then Err.Raise 513, "WinRT_FactoryCreateScheduledToast", "Factory or XmlDocument is null."
    If WRT_DateTimeIsZero(deliveryTime) Then Err.Raise 5, "WinRT_FactoryCreateScheduledToast", "DeliveryTime is zero."

    ' ヘルパーに依存せず、この場で UDT の 8 バイトを LongLong / Currency へ直接コピーする
    RtlMoveMemory deliveryTicks, deliveryTime.dwLowDateTime, 8
    RtlMoveMemory deliveryCy, deliveryTime.dwLowDateTime, 8
    rawHex = "raw=0x" & Right$("00000000" & Hex$(deliveryTime.dwHighDateTime), 8) & _
             Right$("00000000" & Hex$(deliveryTime.dwLowDateTime), 8) & _
             " ticks=0x" & Hex$(deliveryTicks)
    pScheduled = 0

    On Error Resume Next
    Err.Clear
    WinRT_CallComMethod pFactory, VT_IScheduledToastNotificationFactory_Create, vbLong, _
        WinRT_vbPtr, pXmlDoc, WinRT_vtCy, deliveryCy, WinRT_vbPtr, VarPtr(pScheduled)
    If Err.Number = 0 And pScheduled <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    pScheduled = 0
    WinRT_CallComMethod pFactory, VT_IScheduledToastNotificationFactory_Create, vbLong, _
        WinRT_vbPtr, pXmlDoc, WinRT_vbPtr, deliveryTicks, WinRT_vbPtr, VarPtr(pScheduled)
    If Err.Number = 0 And pScheduled <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    If errNum = 0 Then errNum = Err.Number
    If Len(errDesc) = 0 Then errDesc = Err.Description
    On Error GoTo 0

    If errNum = 0 Then errNum = &H80070057
    If Len(errDesc) = 0 Then errDesc = "CreateScheduledToastNotification failed: 0x" & Hex$(errNum)
    Err.Raise errNum, "WinRT_FactoryCreateScheduledToast", errDesc & " " & rawHex
End Sub

Private Sub WinRT_PropertyValueCreateDateTime( _
    ByVal pPropertyValueStatics As LongPtr, _
    ByRef dateTimeValue As WRT_DateTime, _
    ByRef pPropertyValue As LongPtr)

    Dim deliveryCy As Currency
    Dim deliveryTicks As LongLong
    Dim errNum As Long

    If pPropertyValueStatics = 0 Then Err.Raise 513, "WinRT_PropertyValueCreateDateTime", "PropertyValue statics is null."

    pPropertyValue = 0
    RtlMoveMemory deliveryCy, dateTimeValue.dwLowDateTime, 8

    On Error Resume Next
    Err.Clear
    WinRT_CallComMethod pPropertyValueStatics, VT_IPropertyValueStatics_CreateDateTime, vbLong, _
        WinRT_vtCy, deliveryCy, WinRT_vbPtr, VarPtr(pPropertyValue)
    If Err.Number = 0 And pPropertyValue <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    errNum = Err.Number
    Err.Clear
    pPropertyValue = 0
    RtlMoveMemory deliveryTicks, dateTimeValue.dwLowDateTime, 8
    WinRT_CallComMethod pPropertyValueStatics, VT_IPropertyValueStatics_CreateDateTime, vbLong, _
        WinRT_vbPtr, deliveryTicks, WinRT_vbPtr, VarPtr(pPropertyValue)
    If Err.Number = 0 And pPropertyValue <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo 0
    If errNum = 0 Then errNum = Err.Number
    If errNum = 0 Then errNum = &H80070057
    Err.Raise errNum, "WinRT_PropertyValueCreateDateTime", "CreateDateTime failed: 0x" & Hex$(errNum)
End Sub

Private Function WinRT_NormalizeScheduleId(ByVal scheduleId As String) As String
    If Len(scheduleId) = 0 Then
        WinRT_NormalizeScheduleId = "ExcelSchedule"
    ElseIf Len(scheduleId) > WinRT_MaxScheduleIdLen Then
        Err.Raise 5, "WinRT_NormalizeScheduleId", _
            "Schedule_ID must be <= " & WinRT_MaxScheduleIdLen & " chars (WPN_E_DEV_ID_SIZE). Got " & Len(scheduleId) & ": """ & scheduleId & """"
    Else
        WinRT_NormalizeScheduleId = scheduleId
    End If
End Function

Private Sub WinRT_FinalizeRo(ByVal initialized As Boolean)
    ' 通知表示→更新の連続呼び出しで WinRT ランタイムを破棄しない
End Sub

Private Sub WinRT_EnsureRoInitialized(ByRef initialized As Boolean)
    Dim hr As Long

    If Not WinRT_RoInitialized Then
        hr = RoInitialize(WinRT_RO_INIT_SINGLETHREADED)
        If hr <> 0 And hr <> 1 And hr <> WinRT_RPC_E_CHANGED_MODE Then
            Err.Raise hr, "RoInitialize", "RoInitialize failed: 0x" & Hex$(hr)
        End If
        WinRT_RoInitialized = True
    End If
    initialized = True
End Sub

'==================================================================================
' トーストアクティブ化（イベントコールバック）— DLL 不要・純 VBA + DispCallFunc
'   標準モジュール関数の関数ポインタで IUnknown ベースのネイティブ COM デリゲート
'   (vtable 4 エントリ: QI / AddRef / Release / Invoke) を自前構築し、
'   IToastNotification.add_Activated / add_Dismissed / add_Failed に登録する。
'==================================================================================

' トーストに Activated / Dismissed / Failed の 3 ハンドラを登録する
Private Sub WinRT_RegisterToastEventHandlers(ByVal pToast As LongPtr)
    Dim token As LongLong

    If pToast = 0 Then Exit Sub

    ' イベント登録の失敗で通知表示自体を止めないよう、ここでは握りつぶす
    On Error Resume Next
    WinRT_EnsureDelegates

    token = 0
    WinRT_CallComMethod pToast, VT_IToastNotification_AddActivated, vbLong, _
        WinRT_vbPtr, WRT_Act_pObject(WRT_DELEGATE_ACTIVATED), WinRT_vbPtr, VarPtr(token)

    token = 0
    WinRT_CallComMethod pToast, VT_IToastNotification_AddDismissed, vbLong, _
        WinRT_vbPtr, WRT_Act_pObject(WRT_DELEGATE_DISMISSED), WinRT_vbPtr, VarPtr(token)

    token = 0
    WinRT_CallComMethod pToast, VT_IToastNotification_AddFailed, vbLong, _
        WinRT_vbPtr, WRT_Act_pObject(WRT_DELEGATE_FAILED), WinRT_vbPtr, VarPtr(token)
    On Error GoTo 0
End Sub

' 3 種のデリゲート（vtable + オブジェクト実体）を一度だけ構築する
Private Sub WinRT_EnsureDelegates()
    Dim i As Long
    Dim entrySize As LongPtr
    Dim vt(0 To 3) As LongPtr

    If WRT_Act_KeepAlive Is Nothing Then Set WRT_Act_KeepAlive = New Collection
    If WRT_Act_DelegatesReady Then Exit Sub

    IIDFromString StrPtr(WinRT_IID_IUnknown), WRT_Act_iidUnknown
    IIDFromString StrPtr(WinRT_IID_TypedEventHandler_Activated), WRT_Act_iidDelegate(WRT_DELEGATE_ACTIVATED)
    IIDFromString StrPtr(WinRT_IID_TypedEventHandler_Dismissed), WRT_Act_iidDelegate(WRT_DELEGATE_DISMISSED)
    IIDFromString StrPtr(WinRT_IID_TypedEventHandler_Failed), WRT_Act_iidDelegate(WRT_DELEGATE_FAILED)

    entrySize = LenB(WRT_Act_pObject(0))

    For i = 0 To WRT_DELEGATE_COUNT - 1
        vt(0) = WinRT_DelegateAddr(AddressOf WinRT_Act_QueryInterface)
        vt(1) = WinRT_DelegateAddr(AddressOf WinRT_Act_AddRef)
        vt(2) = WinRT_DelegateAddr(AddressOf WinRT_Act_Release)
        Select Case i
            Case WRT_DELEGATE_ACTIVATED: vt(3) = WinRT_DelegateAddr(AddressOf WinRT_Act_InvokeActivated)
            Case WRT_DELEGATE_DISMISSED: vt(3) = WinRT_DelegateAddr(AddressOf WinRT_Act_InvokeDismissed)
            Case WRT_DELEGATE_FAILED:    vt(3) = WinRT_DelegateAddr(AddressOf WinRT_Act_InvokeFailed)
        End Select

        WRT_Act_pVTable(i) = GlobalAlloc(WRT_GPTR, 4 * entrySize)
        If WRT_Act_pVTable(i) = 0 Then Err.Raise 7, "WinRT_EnsureDelegates", "GlobalAlloc vtable failed."
        RtlMoveMemory ByVal WRT_Act_pVTable(i), vt(0), 4 * entrySize

        WRT_Act_pObject(i) = GlobalAlloc(WRT_GPTR, entrySize)
        If WRT_Act_pObject(i) = 0 Then Err.Raise 7, "WinRT_EnsureDelegates", "GlobalAlloc object failed."
        RtlMoveMemory ByVal WRT_Act_pObject(i), WRT_Act_pVTable(i), entrySize

        WRT_Act_RefCount(i) = 1
    Next i

    WRT_Act_DelegatesReady = True
End Sub

' AddressOf で得た関数ポインタをそのまま返す（VBA の AddressOf を LongPtr 化）
Private Function WinRT_DelegateAddr(ByVal addr As LongPtr) As LongPtr
    WinRT_DelegateAddr = addr
End Function

' this ポインタからデリゲート種別(0..2)を判定。見つからなければ -1
Private Function WinRT_DelegateIndexFromThis(ByVal this As LongPtr) As Long
    Dim i As Long
    For i = 0 To WRT_DELEGATE_COUNT - 1
        If WRT_Act_pObject(i) = this Then
            WinRT_DelegateIndexFromThis = i
            Exit Function
        End If
    Next i
    WinRT_DelegateIndexFromThis = -1
End Function

'------------------------- IUnknown vtable（3 デリゲート共通）-------------------------

Private Function WinRT_Act_QueryInterface(ByVal this As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    Dim g As WinRT_GUID
    Dim idx As Long

    If riid = 0 Then
        WinRT_Act_QueryInterface = WRT_E_NOINTERFACE
        Exit Function
    End If

    idx = WinRT_DelegateIndexFromThis(this)
    If idx < 0 Then
        ppvObject = 0
        WinRT_Act_QueryInterface = WRT_E_NOINTERFACE
        Exit Function
    End If

    RtlMoveMemory g, ByVal riid, LenB(g)
    If WinRT_GuidEqual(g, WRT_Act_iidUnknown) Or WinRT_GuidEqual(g, WRT_Act_iidDelegate(idx)) Then
        ppvObject = this
        WRT_Act_RefCount(idx) = WRT_Act_RefCount(idx) + 1
        WinRT_Act_QueryInterface = WRT_S_OK
    Else
        ppvObject = 0
        WinRT_Act_QueryInterface = WRT_E_NOINTERFACE
    End If
End Function

Private Function WinRT_Act_AddRef(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = WinRT_DelegateIndexFromThis(this)
    If idx < 0 Then
        WinRT_Act_AddRef = 1
        Exit Function
    End If
    WRT_Act_RefCount(idx) = WRT_Act_RefCount(idx) + 1
    WinRT_Act_AddRef = WRT_Act_RefCount(idx)
End Function

Private Function WinRT_Act_Release(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = WinRT_DelegateIndexFromThis(this)
    If idx < 0 Then
        WinRT_Act_Release = 1
        Exit Function
    End If
    If WRT_Act_RefCount(idx) > 0 Then WRT_Act_RefCount(idx) = WRT_Act_RefCount(idx) - 1
    WinRT_Act_Release = WRT_Act_RefCount(idx)
End Function

'------------------------- Invoke（種別ごと）-------------------------

' ITypedEventHandler<ToastNotification, IInspectable>::Invoke(this, sender, args)
Private Function WinRT_Act_InvokeActivated(ByVal this As LongPtr, ByVal pSender As LongPtr, ByVal pArgs As LongPtr) As Long
    Dim macroName As String
    Dim groupInfo As String
    Dim userInputs As Object

    On Error GoTo EH
    macroName = WinRT_Act_GetActivatedArguments(pArgs)
    If Len(macroName) > 0 Then
        groupInfo = WinRT_Act_GetToastGroup(pSender)
        Set userInputs = WinRT_Act_BuildUserInputDictionary(pArgs)
        WinRT_Act_RunExcelMacro groupInfo, macroName, userInputs
    End If
    WinRT_Act_InvokeActivated = WRT_S_OK
    Exit Function
EH:
    WinRT_Act_InvokeActivated = WRT_S_OK
End Function

' ITypedEventHandler<ToastNotification, ToastDismissedEventArgs>::Invoke
Private Function WinRT_Act_InvokeDismissed(ByVal this As LongPtr, ByVal pSender As LongPtr, ByVal pArgs As LongPtr) As Long
    Dim groupInfo As String
    Dim dict As Object
    Dim reason As Long

    On Error GoTo EH
    groupInfo = WinRT_Act_GetToastGroup(pSender)

    reason = 0
    If pArgs <> 0 Then WinRT_CallComMethod pArgs, VT_IToastDismissedEventArgs_GetReason, vbLong, WinRT_vbPtr, VarPtr(reason)

    Set dict = CreateObject("Scripting.Dictionary")
    dict("ToastNotification.Tag") = WinRT_Act_GetToastTag(pSender)
    dict("ToastDismissalReason") = CStr(reason)

    WinRT_Act_RunExcelMacro groupInfo, WRT_MacroName_Dismissed, dict
    WinRT_Act_InvokeDismissed = WRT_S_OK
    Exit Function
EH:
    WinRT_Act_InvokeDismissed = WRT_S_OK
End Function

' ITypedEventHandler<ToastNotification, ToastFailedEventArgs>::Invoke
Private Function WinRT_Act_InvokeFailed(ByVal this As LongPtr, ByVal pSender As LongPtr, ByVal pArgs As LongPtr) As Long
    Dim groupInfo As String
    Dim dict As Object
    Dim errorCode As Long

    On Error GoTo EH
    groupInfo = WinRT_Act_GetToastGroup(pSender)

    errorCode = 0
    If pArgs <> 0 Then WinRT_CallComMethod pArgs, VT_IToastFailedEventArgs_GetErrorCode, vbLong, WinRT_vbPtr, VarPtr(errorCode)

    Set dict = CreateObject("Scripting.Dictionary")
    dict("ToastNotification.Tag") = WinRT_Act_GetToastTag(pSender)
    dict("ErrorCode") = "0x" & Right$("00000000" & Hex$(errorCode), 8)

    WinRT_Act_RunExcelMacro groupInfo, WRT_MacroName_Failed, dict
    WinRT_Act_InvokeFailed = WRT_S_OK
    Exit Function
EH:
    WinRT_Act_InvokeFailed = WRT_S_OK
End Function

'------------------------- イベント引数の取り出し（DispCallFunc, no-TLB）-------------------------

' IToastActivatedEventArgs.Arguments（launch / action arguments のマクロ名）
Private Function WinRT_Act_GetActivatedArguments(ByVal pArgs As LongPtr) As String
    Dim pActArgs As LongPtr
    Dim iid As WinRT_GUID
    Dim hStr As LongPtr

    If pArgs = 0 Then Exit Function
    IIDFromString StrPtr(WinRT_IID_IToastActivatedEventArgs), iid
    WinRT_CallComMethod pArgs, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pActArgs)
    If pActArgs = 0 Then Exit Function

    hStr = 0
    WinRT_CallComMethod pActArgs, VT_IToastActivatedEventArgs_GetArguments, vbLong, WinRT_vbPtr, VarPtr(hStr)
    WinRT_Act_GetActivatedArguments = WinRT_HStringToString(hStr)
    If hStr <> 0 Then WindowsDeleteString hStr
    WinRT_CallComMethod pActArgs, VT_RELEASE, vbLong
End Function

' IToastNotification2.Group（"ブック名|PID" 形式）
Private Function WinRT_Act_GetToastGroup(ByVal pSender As LongPtr) As String
    WinRT_Act_GetToastGroup = WinRT_Act_GetToastStringProp(pSender, VT_IToastNotification2_GetGroup)
End Function

' IToastNotification2.Tag
Private Function WinRT_Act_GetToastTag(ByVal pSender As LongPtr) As String
    WinRT_Act_GetToastTag = WinRT_Act_GetToastStringProp(pSender, VT_IToastNotification2_GetTag)
End Function

Private Function WinRT_Act_GetToastStringProp(ByVal pSender As LongPtr, ByVal vtIndex As Long) As String
    Dim pToast2 As LongPtr
    Dim iid As WinRT_GUID
    Dim hStr As LongPtr

    If pSender = 0 Then Exit Function
    IIDFromString StrPtr(WinRT_IID_IToastNotification2), iid
    WinRT_CallComMethod pSender, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pToast2)
    If pToast2 = 0 Then Exit Function

    hStr = 0
    WinRT_CallComMethod pToast2, vtIndex, vbLong, WinRT_vbPtr, VarPtr(hStr)
    WinRT_Act_GetToastStringProp = WinRT_HStringToString(hStr)
    If hStr <> 0 Then WindowsDeleteString hStr
    WinRT_CallComMethod pToast2, VT_RELEASE, vbLong
End Function

' IToastActivatedEventArgs2.UserInput（ValueSet）を Scripting.Dictionary に変換
Private Function WinRT_Act_BuildUserInputDictionary(ByVal pArgs As LongPtr) As Object
    Dim pActArgs2 As LongPtr
    Dim pPropSet As LongPtr
    Dim pIterable As LongPtr
    Dim pIterator As LongPtr
    Dim pPair As LongPtr
    Dim hasCurrent As Byte
    Dim hKey As LongPtr
    Dim pValue As LongPtr
    Dim keyText As String
    Dim iidArgs2 As WinRT_GUID
    Dim iidIterable As WinRT_GUID
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")
    If pArgs = 0 Then GoTo Done

    On Error GoTo Done
    IIDFromString StrPtr(WinRT_IID_IToastActivatedEventArgs2), iidArgs2
    WinRT_CallComMethod pArgs, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidArgs2), WinRT_vbPtr, VarPtr(pActArgs2)
    If pActArgs2 = 0 Then GoTo Done

    WinRT_CallComMethod pActArgs2, VT_IToastActivatedEventArgs2_GetUserInput, vbLong, WinRT_vbPtr, VarPtr(pPropSet)
    If pPropSet = 0 Then GoTo Done

    IIDFromString StrPtr(WinRT_IID_IIterable_KVP_IInspectable), iidIterable
    WinRT_CallComMethod pPropSet, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidIterable), WinRT_vbPtr, VarPtr(pIterable)
    If pIterable = 0 Then GoTo Done

    WinRT_CallComMethod pIterable, VT_IIterable_First, vbLong, WinRT_vbPtr, VarPtr(pIterator)
    If pIterator = 0 Then GoTo Done

    hasCurrent = 0
    WinRT_CallComMethod pIterator, VT_IIterator_GetHasCurrent, vbLong, WinRT_vbPtr, VarPtr(hasCurrent)
    Do While hasCurrent <> 0
        pPair = 0
        WinRT_CallComMethod pIterator, VT_IIterator_GetCurrent, vbLong, WinRT_vbPtr, VarPtr(pPair)
        If pPair <> 0 Then
            hKey = 0
            WinRT_CallComMethod pPair, VT_IKeyValuePair_GetKey, vbLong, WinRT_vbPtr, VarPtr(hKey)
            keyText = WinRT_HStringToString(hKey)
            If hKey <> 0 Then WindowsDeleteString hKey

            pValue = 0
            WinRT_CallComMethod pPair, VT_IKeyValuePair_GetValue, vbLong, WinRT_vbPtr, VarPtr(pValue)
            If Len(keyText) > 0 Then dict(keyText) = WinRT_Act_InspectableToString(pValue)
            If pValue <> 0 Then WinRT_CallComMethod pValue, VT_RELEASE, vbLong
            WinRT_CallComMethod pPair, VT_RELEASE, vbLong
        End If

        hasCurrent = 0
        WinRT_CallComMethod pIterator, VT_IIterator_MoveNext, vbLong, WinRT_vbPtr, VarPtr(hasCurrent)
    Loop

Done:
    On Error Resume Next
    If pIterator <> 0 Then WinRT_CallComMethod pIterator, VT_RELEASE, vbLong
    If pIterable <> 0 Then WinRT_CallComMethod pIterable, VT_RELEASE, vbLong
    If pPropSet <> 0 Then WinRT_CallComMethod pPropSet, VT_RELEASE, vbLong
    If pActArgs2 <> 0 Then WinRT_CallComMethod pActArgs2, VT_RELEASE, vbLong
    On Error GoTo 0
    Set WinRT_Act_BuildUserInputDictionary = dict
End Function

' IInspectable の値を文字列化（IPropertyValue.GetString → IStringable.ToString の順）
Private Function WinRT_Act_InspectableToString(ByVal pValue As LongPtr) As String
    Dim pPropVal As LongPtr
    Dim pStringable As LongPtr
    Dim iid As WinRT_GUID
    Dim hStr As LongPtr

    If pValue = 0 Then Exit Function

    IIDFromString StrPtr(WinRT_IID_IPropertyValue), iid
    On Error Resume Next
    WinRT_CallComMethod pValue, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pPropVal)
    On Error GoTo 0
    If pPropVal <> 0 Then
        hStr = 0
        On Error Resume Next
        WinRT_CallComMethod pPropVal, VT_IPropertyValue_GetString, vbLong, WinRT_vbPtr, VarPtr(hStr)
        On Error GoTo 0
        WinRT_Act_InspectableToString = WinRT_HStringToString(hStr)
        If hStr <> 0 Then WindowsDeleteString hStr
        WinRT_CallComMethod pPropVal, VT_RELEASE, vbLong
        If Len(WinRT_Act_InspectableToString) > 0 Then Exit Function
    End If

    IIDFromString StrPtr(WinRT_IID_IStringable), iid
    On Error Resume Next
    WinRT_CallComMethod pValue, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pStringable)
    On Error GoTo 0
    If pStringable <> 0 Then
        hStr = 0
        On Error Resume Next
        WinRT_CallComMethod pStringable, VT_IStringable_ToString, vbLong, WinRT_vbPtr, VarPtr(hStr)
        On Error GoTo 0
        WinRT_Act_InspectableToString = WinRT_HStringToString(hStr)
        If hStr <> 0 Then WindowsDeleteString hStr
        WinRT_CallComMethod pStringable, VT_RELEASE, vbLong
    End If
End Function

'------------------------- マクロ実行 -------------------------

Private Sub WinRT_Act_RunExcelMacro(ByVal groupInfo As String, ByVal macroName As String, ByVal userInputs As Object)
    Dim bookName As String
    Dim targetPid As Long
    Dim qualifiedName As String

    WinRT_Act_ParseGroupInfo groupInfo, bookName, targetPid

    ' 別 Excel プロセス宛ての通知は無視（多重起動対策）
    If targetPid <> 0 And targetPid <> WinRT_GetCurrentProcessId() Then Exit Sub

    If InStr(macroName, "!") > 0 Then
        qualifiedName = macroName
    ElseIf Len(bookName) > 0 Then
        qualifiedName = "'" & bookName & "'!" & macroName
    Else
        qualifiedName = macroName
    End If

    If userInputs Is Nothing Then
        Application.Run qualifiedName
    Else
        Application.Run qualifiedName, userInputs
    End If
End Sub

Private Sub WinRT_Act_ParseGroupInfo(ByVal groupInfo As String, ByRef bookName As String, ByRef targetPid As Long)
    Dim pos As Long

    bookName = vbNullString
    targetPid = 0
    pos = InStrRev(groupInfo, WRT_GroupDelimiter)
    If pos = 0 Then
        bookName = groupInfo
        Exit Sub
    End If
    bookName = Left$(groupInfo, pos - 1)
    On Error Resume Next
    targetPid = CLng(Mid$(groupInfo, pos + 1))
    On Error GoTo 0
End Sub

Private Function WinRT_GuidEqual(ByRef a As WinRT_GUID, ByRef b As WinRT_GUID) As Boolean
    Dim i As Long
    If a.Data1 <> b.Data1 Then Exit Function
    If a.Data2 <> b.Data2 Then Exit Function
    If a.Data3 <> b.Data3 Then Exit Function
    For i = 0 To 7
        If a.Data4(i) <> b.Data4(i) Then Exit Function
    Next i
    WinRT_GuidEqual = True
End Function

Private Function WinRT_HStringToString(ByVal hStr As LongPtr) As String
    Dim bufLen As Long
    Dim pWchar As LongPtr

    If hStr = 0 Then Exit Function
    pWchar = WindowsGetStringRawBuffer(hStr, bufLen)
    If pWchar = 0 Or bufLen <= 0 Then Exit Function
    WinRT_HStringToString = String$(bufLen, vbNullChar)
    RtlMoveMemory ByVal StrPtr(WinRT_HStringToString), ByVal pWchar, CLngPtr(bufLen) * 2
End Function

' IAsyncOperation / IAsyncAction の完了を待ち、完了後の HRESULT(ErrorCode) を返す。0=成功。
' WinRT 非同期は QI で IAsyncInfo を取り、Status を Started 以外になるまでポーリングする。
Private Function WinRT_WaitAsync(ByVal pAsync As LongPtr) As Long
    Dim iidAsyncInfo As WinRT_GUID
    Dim pAsyncInfo As LongPtr
    Dim status As Long
    Dim errorCode As Long
    Dim waitCount As Long

    If pAsync = 0 Then
        WinRT_WaitAsync = &H80004003
        Exit Function
    End If

    IIDFromString StrPtr(WinRT_IID_IAsyncInfo), iidAsyncInfo
    pAsyncInfo = 0
    WinRT_CallComMethod pAsync, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidAsyncInfo), WinRT_vbPtr, VarPtr(pAsyncInfo)
    If pAsyncInfo = 0 Then
        WinRT_WaitAsync = &H80004002
        Exit Function
    End If

    Do
        status = WinRT_AsyncStatus_Started
        WinRT_CallComMethod pAsyncInfo, VT_IAsyncInfo_GetStatus, vbLong, WinRT_vbPtr, VarPtr(status)
        If status <> WinRT_AsyncStatus_Started Then Exit Do
        DoEvents
        waitCount = waitCount + 1
        If waitCount > 50000 Then Exit Do
    Loop

    If status = WinRT_AsyncStatus_Completed Then
        errorCode = 0
    Else
        errorCode = 0
        WinRT_CallComMethod pAsyncInfo, VT_IAsyncInfo_GetErrorCode, vbLong, WinRT_vbPtr, VarPtr(errorCode)
        If status = WinRT_AsyncStatus_Canceled And errorCode = 0 Then errorCode = &H800704C7
        If status = WinRT_AsyncStatus_Started And errorCode = 0 Then errorCode = &H8001011F
    End If

    WinRT_CallComMethod pAsyncInfo, VT_RELEASE, vbLong
    WinRT_WaitAsync = errorCode
End Function

Private Function WinRT_CallComMethod(ByVal pUnk As LongPtr, ByVal vTableIndex As Long, ByVal retType As Integer, ParamArray args() As Variant) As Variant
    Dim vTableOffset As LongPtr
    Dim argTypes() As Integer
    Dim argPointers() As LongPtr
    Dim argValues() As Variant
    Dim argCount As Long
    Dim i As Long
    Dim hRes As Long
    Dim vResult As Variant

    If pUnk = 0 Then Err.Raise 513, "WinRT_CallComMethod", "COM pointer is null. Index=" & vTableIndex
#If Win64 Then
    vTableOffset = vTableIndex * 8
#Else
    vTableOffset = vTableIndex * 4
#End If

    argCount = (UBound(args) + 1) \ 2
    If argCount > 0 Then
        ReDim argTypes(0 To argCount - 1)
        ReDim argPointers(0 To argCount - 1)
        ReDim argValues(0 To argCount - 1)
        For i = 0 To argCount - 1
            argTypes(i) = CInt(args(i * 2))
            argValues(i) = args(i * 2 + 1)
            argPointers(i) = VarPtr(argValues(i))
        Next i
        hRes = DispCallFunc(pUnk, vTableOffset, 4, retType, argCount, argTypes(0), argPointers(0), vResult)
    Else
        hRes = DispCallFunc(pUnk, vTableOffset, 4, retType, 0, ByVal 0&, ByVal 0&, vResult)
    End If

    If hRes <> 0 Then Err.Raise hRes, "WinRT_CallComMethod", "DispCallFunc failed at vtable index " & vTableIndex & ": 0x" & Hex$(hRes)
    If vTableIndex <> VT_RELEASE Then
        Select Case VarType(vResult)
            Case vbLong, vbInteger, vbByte
                If CLng(vResult) < 0 Then Err.Raise CLng(vResult), "WinRT_CallComMethod", "COM method failed at vtable index " & vTableIndex & ": 0x" & Hex$(CLng(vResult))
            Case vbLongLong
                If vResult < 0 Then Err.Raise CLng(vResult), "WinRT_CallComMethod", "COM method failed at vtable index " & vTableIndex & ": 0x" & Hex$(CLng(vResult))
        End Select
    End If
    WinRT_CallComMethod = vResult
End Function
