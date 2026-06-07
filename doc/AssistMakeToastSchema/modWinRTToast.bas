Attribute VB_Name = "modWinRTToast"
Option Explicit

' modWinRTToast
' WinRT Interfaces x64.txt / ToastCompat_WinRT.bas に基づく VTable 直呼び出し。
' AppNotificationBuilderVBA.dll の Show / Update / Remove / CheckNotificationSetting を置き換える。

Private Declare PtrSafe Function RoInitialize Lib "combase.dll" (ByVal initType As Long) As Long
Private Declare PtrSafe Function RoUninitialize Lib "combase.dll" () As Long
Private Declare PtrSafe Function WindowsCreateString Lib "combase.dll" (ByVal sourceString As LongPtr, ByVal length As Long, ByRef hstring As LongPtr) As Long
Private Declare PtrSafe Function WindowsDeleteString Lib "combase.dll" (ByVal hstring As LongPtr) As Long
Private Declare PtrSafe Function RoGetActivationFactory Lib "combase.dll" (ByVal activatableClassId As LongPtr, ByRef iid As Any, ByRef factory As LongPtr) As Long
Private Declare PtrSafe Function RoActivateInstance Lib "combase.dll" (ByVal activatableClassId As LongPtr, ByRef instance As LongPtr) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Integer, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long

Private Type WinRT_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type WinRT_DateTime
    UniversalTime As Currency
End Type

#If Win64 Then
    Private Const WinRT_vbPtr As Integer = 20
#Else
    Private Const WinRT_vbPtr As Integer = 3
#End If

Private Const VT_QI As Long = 0
Private Const VT_RELEASE As Long = 2
Private Const VT_IToastNotifier_Show As Long = 6
Private Const VT_IToastNotifier_AddToSchedule As Long = 9
Private Const VT_IToastNotification2_SetTag As Long = 6
Private Const VT_IToastNotification2_SetGroup As Long = 7
Private Const VT_IToastNotification2_SetSuppressPopup As Long = 8
Private Const VT_IToastNotification4_SetData As Long = 7
Private Const VT_IToastNotification6_SetExpiresOnReboot As Long = 7
Private Const VT_IToastNotificationHistory_RemoveGroupedTagWithId As Long = 8
Private Const VT_IToastNotificationHistory_RemoveGroupWithId As Long = 7
Private Const VT_IToastNotificationHistory_ClearWithId As Long = 12

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
    ExpirationTime As Date
End Type

Public Type WinRT_DataBinding
    TitleText As String
    ContentsText As String
    AttributionText As String
    ProgressTitle As String
    ProgressValueStringOverride As String
    ProgressStatus As String
    ProgressValue As Double
End Type

Public Sub WinRT_ShowToastNotification(ByRef Config As WinRT_ToastConfig, ByRef Binding As WinRT_DataBinding)
    Dim initialized As Boolean
    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized

    If Config.Schedule_DeliveryTime > 0 Then
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

    pNotifier = WinRT_CreateToastNotifier(Config.AppUserModelID, Config.CollectionID)
    If pNotifier = 0 Then Err.Raise 513, , "CreateToastNotifier failed."

    IIDFromString StrPtr("{354389C6-7C01-4BD5-9C20-604340CD2B74}"), iidNotifier2
    WinRT_CallComMethod pNotifier, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidNotifier2), WinRT_vbPtr, VarPtr(pNotifier2)
    If pNotifier2 = 0 Then Err.Raise 513, , "QueryInterface IToastNotifier2 failed."

    pData = WinRT_CreateNotificationData(Binding)
    If pData = 0 Then Err.Raise 513, , "CreateNotificationData failed."

    hStrTag = WinRT_CreateHString(Config.Tag)
    hStrGroup = WinRT_CreateHString(Config.Group)
    WinRT_CallComMethod pNotifier2, 6, vbLong, WinRT_vbPtr, pData, WinRT_vbPtr, hStrTag, WinRT_vbPtr, hStrGroup, WinRT_vbPtr, VarPtr(updateResult)
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

Public Sub WinRT_RemoveToastNotification(ByRef Config As WinRT_ToastConfig)
    Dim initialized As Boolean
    Dim pHistory As LongPtr
    Dim hStrTag As LongPtr
    Dim hStrGroup As LongPtr
    Dim hStrAppId As LongPtr

    On Error GoTo Cleanup
    WinRT_EnsureRoInitialized initialized

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

    settingValue = CLng(WinRT_CallComMethod(pNotifier, 8, vbLong))
    WinRT_CheckNotificationSetting = settingValue

Cleanup:
    If Err.Number <> 0 Then WinRT_CheckNotificationSetting = Err.Number
    On Error Resume Next
    If pNotifier <> 0 Then WinRT_CallComMethod pNotifier, VT_RELEASE, vbLong
    WinRT_FinalizeRo initialized
    On Error GoTo 0
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

    pNotifier = WinRT_CreateToastNotifier(Config.AppUserModelID, Config.CollectionID)
    If pNotifier = 0 Then Err.Raise 513, , "CreateToastNotifier failed."
    WinRT_CallComMethod pNotifier, VT_IToastNotifier_Show, vbLong, WinRT_vbPtr, pToast

    If hStrTag <> 0 Then WindowsDeleteString hStrTag
    If hStrGroup <> 0 Then WindowsDeleteString hStrGroup
    If pNotifier <> 0 Then WinRT_CallComMethod pNotifier, VT_RELEASE, vbLong
    If pToast <> 0 Then WinRT_CallComMethod pToast, VT_RELEASE, vbLong
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
    Dim deliveryTime As WinRT_DateTime
    Dim iidScheduledFactory As WinRT_GUID
    Dim iidScheduled2 As WinRT_GUID
    Dim iidScheduled4 As WinRT_GUID

    pXmlDoc = WinRT_LoadXmlDocument(Config.XmlTemplate)
    If pXmlDoc = 0 Then Err.Raise 513, , "LoadXmlDocument failed."

    IIDFromString StrPtr("{E7BED191-0BB9-4189-8394-31761B476FD7}"), iidScheduledFactory
    WinRT_GetActivationFactory "Windows.UI.Notifications.ScheduledToastNotification", iidScheduledFactory, pScheduledFactory
    If pScheduledFactory = 0 Then Err.Raise 513, , "ScheduledToastNotification factory failed."

    deliveryTime.UniversalTime = WinRT_VbaDateToUniversalTime(Config.Schedule_DeliveryTime)
    WinRT_CallComMethod pScheduledFactory, 6, vbLong, WinRT_vbPtr, pXmlDoc, WinRT_vbPtr, VarPtr(deliveryTime), WinRT_vbPtr, VarPtr(pScheduled)
    WinRT_CallComMethod pScheduledFactory, VT_RELEASE, vbLong
    WinRT_CallComMethod pXmlDoc, VT_RELEASE, vbLong
    pScheduledFactory = 0
    pXmlDoc = 0
    If pScheduled = 0 Then Err.Raise 513, , "CreateScheduledToastNotification failed."

    IIDFromString StrPtr("{A66EA09C-31B4-43B0-B5DD-7A40E85363B1}"), iidScheduled2
    WinRT_CallComMethod pScheduled, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidScheduled2), WinRT_vbPtr, VarPtr(pScheduled2)
    If pScheduled2 <> 0 Then
        hStrScheduleId = WinRT_CreateHString(Config.Schedule_ID)
        hStrTag = WinRT_CreateHString(Config.Tag)
        hStrGroup = WinRT_CreateHString(Config.Group)
        If hStrScheduleId <> 0 Then WinRT_CallComMethod pScheduled, 8, vbLong, WinRT_vbPtr, hStrScheduleId
        If hStrTag <> 0 Then WinRT_CallComMethod pScheduled2, VT_IToastNotification2_SetTag, vbLong, WinRT_vbPtr, hStrTag
        If hStrGroup <> 0 Then WinRT_CallComMethod pScheduled2, VT_IToastNotification2_SetGroup, vbLong, WinRT_vbPtr, hStrGroup
        If Config.SuppressPopup Then WinRT_CallComMethod pScheduled2, VT_IToastNotification2_SetSuppressPopup, vbLong, vbByte, CByte(1)
        WinRT_CallComMethod pScheduled2, VT_RELEASE, vbLong
        pScheduled2 = 0
    End If

    If Config.ExpirationTime > 0 Then
        IIDFromString StrPtr("{98429E8B-BD32-4A3B-9D15-22AEA49462A1}"), iidScheduled4
        WinRT_CallComMethod pScheduled, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidScheduled4), WinRT_vbPtr, VarPtr(pScheduled4)
        If pScheduled4 <> 0 Then
            Dim pExpirationRef As LongPtr
            pExpirationRef = WinRT_CreateDateTimeReference(Config.ExpirationTime)
            If pExpirationRef <> 0 Then
                WinRT_CallComMethod pScheduled4, 7, vbLong, WinRT_vbPtr, pExpirationRef
                WinRT_CallComMethod pExpirationRef, VT_RELEASE, vbLong
            End If
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
        pNotifier = CLngPtr(WinRT_CallComMethod(pAsync, 7, vbLong))
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

    pHistory = CLngPtr(WinRT_CallComMethod(pManagerStatics2, 6, vbLong))
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
        pHistory = CLngPtr(WinRT_CallComMethod(pAsync, 7, vbLong))
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

    hasProgress = (Len(Binding.ProgressStatus) > 0)
    WinRT_CallComMethod pData, 6, vbLong, WinRT_vbPtr, VarPtr(pMap)
    If pMap <> 0 Then
        If Len(Binding.TitleText) > 0 Then WinRT_InsertMapValue pMap, "TopTextTitle", Binding.TitleText
        If Len(Binding.ContentsText) > 0 Then WinRT_InsertMapValue pMap, "TopTextContents", Binding.ContentsText
        If Len(Binding.AttributionText) > 0 Then WinRT_InsertMapValue pMap, "TopTextAttribution", Binding.AttributionText
        If hasProgress Then
            If Len(Binding.ProgressTitle) > 0 Then WinRT_InsertMapValue pMap, "ProgressTitle", Binding.ProgressTitle
            WinRT_InsertMapValue pMap, "ProgressStatus", Binding.ProgressStatus
            If Binding.ProgressValue < 0 Then
                WinRT_InsertMapValue pMap, "ProgressValue", "Indeterminate"
            Else
                WinRT_InsertMapValue pMap, "ProgressValue", CStr(Binding.ProgressValue)
            End If
            If Len(Binding.ProgressValueStringOverride) > 0 Then WinRT_InsertMapValue pMap, "ProgressValueString", Binding.ProgressValueStringOverride
        End If
        WinRT_CallComMethod pMap, VT_RELEASE, vbLong
    End If

    WinRT_CreateNotificationData = pData
End Function

Private Function WinRT_DataBindingHasContent(ByRef Binding As WinRT_DataBinding) As Boolean
    If Len(Binding.TitleText) > 0 Then WinRT_DataBindingHasContent = True: Exit Function
    If Len(Binding.ContentsText) > 0 Then WinRT_DataBindingHasContent = True: Exit Function
    If Len(Binding.AttributionText) > 0 Then WinRT_DataBindingHasContent = True: Exit Function
    If Len(Binding.ProgressStatus) > 0 Then WinRT_DataBindingHasContent = True
End Function

Private Function WinRT_CreateDateTimeReference(ByVal dt As Date) As LongPtr
    Dim hStrClass As LongPtr
    Dim pPropertyValueStatics As LongPtr
    Dim pPropertyValue As LongPtr
    Dim dateTimeValue As WinRT_DateTime
    Dim iidPropertyValueStatics As WinRT_GUID

    IIDFromString StrPtr("{629DBDCF-4466-40FF-9A1A-484AFD159F5A}"), iidPropertyValueStatics

    hStrClass = WinRT_CreateHString("Windows.Foundation.PropertyValue")
    RoGetActivationFactory hStrClass, iidPropertyValueStatics, pPropertyValueStatics
    If hStrClass <> 0 Then WindowsDeleteString hStrClass
    If pPropertyValueStatics = 0 Then Exit Function

    dateTimeValue.UniversalTime = WinRT_VbaDateToUniversalTime(dt)
    WinRT_CallComMethod pPropertyValueStatics, 13, vbLong, WinRT_vbPtr, VarPtr(dateTimeValue), WinRT_vbPtr, VarPtr(pPropertyValue)
    WinRT_CallComMethod pPropertyValueStatics, VT_RELEASE, vbLong
    WinRT_CreateDateTimeReference = pPropertyValue
End Function

Private Sub WinRT_InsertMapValue(ByVal pMap As LongPtr, ByVal key As String, ByVal val As String)
    Dim hKey As LongPtr
    Dim hVal As LongPtr
    Dim replaced As Byte
    hKey = WinRT_CreateHString(key)
    hVal = WinRT_CreateHString(val)
    WinRT_CallComMethod pMap, 10, vbLong, WinRT_vbPtr, hKey, WinRT_vbPtr, hVal, WinRT_vbPtr, VarPtr(replaced)
    If hKey <> 0 Then WindowsDeleteString hKey
    If hVal <> 0 Then WindowsDeleteString hVal
End Sub

Private Sub WinRT_GetActivationFactory(ByVal className As String, ByRef iid As WinRT_GUID, ByRef factory As LongPtr)
    Dim hStrClass As LongPtr
    hStrClass = WinRT_CreateHString(className)
    RoGetActivationFactory hStrClass, iid, factory
    If hStrClass <> 0 Then WindowsDeleteString hStrClass
End Sub

Private Function WinRT_CreateHString(ByVal s As String) As LongPtr
    Dim hStr As LongPtr
    If Len(s) > 0 Then WindowsCreateString StrPtr(s), Len(s), hStr
    WinRT_CreateHString = hStr
End Function

Private Function WinRT_VbaDateToUniversalTime(ByVal dt As Date) As Currency
    WinRT_VbaDateToUniversalTime = CCur((dt - CDate("1601-01-01")) * 86400# * 10000000#)
End Function

Private Sub WinRT_EnsureRoInitialized(ByRef initialized As Boolean)
    If Not initialized Then
        RoInitialize 1
        initialized = True
    End If
End Sub

Private Sub WinRT_FinalizeRo(ByVal initialized As Boolean)
    If initialized Then RoUninitialize
End Sub

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
    If vTableIndex <> VT_RELEASE And VarType(vResult) = vbLong Then
        If CLng(vResult) < 0 Then Err.Raise CLng(vResult), "WinRT_CallComMethod", "COM method failed at vtable index " & vTableIndex & ": 0x" & Hex$(CLng(vResult))
    End If
    WinRT_CallComMethod = vResult
End Function
