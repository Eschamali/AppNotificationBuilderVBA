Option Explicit

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Integer, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long

Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function WindowsGetStringRawBuffer Lib "combase.dll" (ByVal hstring As LongPtr, ByRef length As Long) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function WinRT_GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function WindowsDeleteString Lib "combase.dll" (ByVal hstring As LongPtr) As Long

#If Win64 Then
    Private Const WinRT_vbPtr As Integer = 20
#Else
    Private Const WinRT_vbPtr As Integer = 3
#End If


Private Type WinRT_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type



Private Const VT_QI As Long = 0
Private Const VT_RELEASE As Long = 2
Private Const WinRT_IID_IAsyncInfo As String = "{00000036-0000-0000-C000-000000000046}"
Private Const WinRT_AsyncStatus_Started As Long = 0
Private Const VT_IAsyncInfo_GetStatus As Long = 7
Private Const WinRT_AsyncStatus_Canceled As Long = 2
Private Const WinRT_AsyncStatus_Completed As Long = 1
Private Const VT_IAsyncInfo_GetErrorCode As Long = 8


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

'==================================================================================
' トーストアクティブ化（イベントコールバック）? DLL 不要・純 VBA + DispCallFunc
'   標準モジュール関数の関数ポインタで IUnknown ベースのネイティブ COM デリゲート
'   (vtable 4 エントリ: QI / AddRef / Release / Invoke) を自前構築し、
'   IToastNotification.add_Activated / add_Dismissed / add_Failed に登録する。
'==================================================================================

' トーストに Activated / Dismissed / Failed の 3 ハンドラを登録する
Public Sub WinRT_RegisterToastEventHandlers(ByVal pToast As LongPtr)
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
