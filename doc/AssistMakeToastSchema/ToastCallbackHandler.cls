VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToastCallbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' ToastCallbackHandler.cls
' トースト Activated / Dismissed / Failed の Invoke 実装と Application.Run ディスパッチ
' =============================================================================

#If Win64 Then
#Else
    #Error "64 ビット VBA 専用"
#End If

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Integer, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" ( _
    ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function WindowsGetStringRawBuffer Lib "combase.dll" ( _
    ByVal hstring As LongPtr, ByRef Length As Long) As LongPtr
Private Declare PtrSafe Function WindowsDeleteString Lib "combase.dll" ( _
    ByVal hstring As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    Destination As Any, Source As Any, ByVal Length As LongPtr)

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
Private Const VT_IToastNotification2_GetTag As Long = 7
Private Const VT_IToastNotification2_GetGroup As Long = 9
Private Const VT_IToastActivatedEventArgs_GetArguments As Long = 6
Private Const VT_IToastActivatedEventArgs2_GetUserInput As Long = 6
Private Const VT_IToastDismissedEventArgs_GetReason As Long = 6
Private Const VT_IToastFailedEventArgs_GetErrorCode As Long = 6
Private Const VT_IIterable_First As Long = 6
Private Const VT_IIterator_GetCurrent As Long = 6
Private Const VT_IIterator_GetHasCurrent As Long = 7
Private Const VT_IIterator_MoveNext As Long = 8
Private Const VT_IKeyValuePair_GetKey As Long = 6
Private Const VT_IKeyValuePair_GetValue As Long = 7
Private Const VT_IPropertyValue_GetString As Long = 19
Private Const VT_IStringable_ToString As Long = 6

Private Const WinRT_IID_IToastNotification2 As String = "{9DFB9FD1-143A-490E-90BF-B9FBA7132DE7}"
Private Const WinRT_IID_IToastActivatedEventArgs As String = "{E3BF92F3-C197-436F-8265-0625824F8DAC}"
Private Const WinRT_IID_IToastActivatedEventArgs2 As String = "{AB7DA512-CC61-568E-81BE-304AC31038FA}"
Private Const WinRT_IID_IIterable_KVP_IInspectable As String = "{FE2F3D47-5D47-5499-8374-430C7CDA0204}"
Private Const WinRT_IID_IPropertyValue As String = "{4BD682DD-7554-40E9-9A9B-82654EDE7E62}"
Private Const WinRT_IID_IStringable As String = "{96369F54-8EB6-48F0-ABCE-C1B211E627C3}"

Private Const WRT_GroupDelimiter As String = "|"
Private Const WRT_S_OK As Long = 0
Private Const WRT_MacroName_Dismissed As String = "ExcelToast_Dismissed"
Private Const WRT_MacroName_Failed As String = "ExcelToast_Failed"


' =============================================================================
' Invoke（COM サンク先・Public）
' =============================================================================

Public Function Invoke_Activated( _
    ByVal pThis As LongPtr, ByVal pSender As LongPtr, ByVal pArgs As LongPtr) As Long

    Dim macroName As String
    Dim groupInfo As String
    Dim userInputs As Object

    On Error GoTo EH
    macroName = GetActivatedArguments(pArgs)
    If Len(macroName) > 0 Then
        groupInfo = GetToastGroup(pSender)
        Set userInputs = BuildUserInputDictionary(pArgs)
        RunExcelMacro groupInfo, macroName, userInputs
    End If
    Invoke_Activated = WRT_S_OK
    Exit Function
EH:
    Invoke_Activated = WRT_S_OK
End Function

Public Function Invoke_Dismissed( _
    ByVal pThis As LongPtr, ByVal pSender As LongPtr, ByVal pArgs As LongPtr) As Long

    Dim groupInfo As String
    Dim dict As Object
    Dim reason As Long

    On Error GoTo EH
    groupInfo = GetToastGroup(pSender)

    reason = 0
    If pArgs <> 0 Then CallComMethod pArgs, VT_IToastDismissedEventArgs_GetReason, vbLong, WinRT_vbPtr, VarPtr(reason)

    Set dict = CreateObject("Scripting.Dictionary")
    dict("ToastNotification.Tag") = GetToastTag(pSender)
    dict("ToastDismissalReason") = CStr(reason)

    RunExcelMacro groupInfo, WRT_MacroName_Dismissed, dict
    Invoke_Dismissed = WRT_S_OK
    Exit Function
EH:
    Invoke_Dismissed = WRT_S_OK
End Function

Public Function Invoke_Failed( _
    ByVal pThis As LongPtr, ByVal pSender As LongPtr, ByVal pArgs As LongPtr) As Long

    Dim groupInfo As String
    Dim dict As Object
    Dim errorCode As Long

    On Error GoTo EH
    groupInfo = GetToastGroup(pSender)

    errorCode = 0
    If pArgs <> 0 Then CallComMethod pArgs, VT_IToastFailedEventArgs_GetErrorCode, vbLong, WinRT_vbPtr, VarPtr(errorCode)

    Set dict = CreateObject("Scripting.Dictionary")
    dict("ToastNotification.Tag") = GetToastTag(pSender)
    dict("ErrorCode") = "0x" & Right$("00000000" & Hex$(errorCode), 8)

    RunExcelMacro groupInfo, WRT_MacroName_Failed, dict
    Invoke_Failed = WRT_S_OK
    Exit Function
EH:
    Invoke_Failed = WRT_S_OK
End Function


' =============================================================================
' イベント引数の取り出し
' =============================================================================

Private Function GetActivatedArguments(ByVal pArgs As LongPtr) As String
    Dim pActArgs As LongPtr
    Dim iid As WinRT_GUID
    Dim hStr As LongPtr

    If pArgs = 0 Then Exit Function
    IIDFromString StrPtr(WinRT_IID_IToastActivatedEventArgs), iid
    CallComMethod pArgs, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pActArgs)
    If pActArgs = 0 Then Exit Function

    hStr = 0
    CallComMethod pActArgs, VT_IToastActivatedEventArgs_GetArguments, vbLong, WinRT_vbPtr, VarPtr(hStr)
    GetActivatedArguments = HStringToString(hStr)
    If hStr <> 0 Then WindowsDeleteString hStr
    CallComMethod pActArgs, VT_RELEASE, vbLong
End Function

Private Function GetToastGroup(ByVal pSender As LongPtr) As String
    GetToastGroup = GetToastStringProp(pSender, VT_IToastNotification2_GetGroup)
End Function

Private Function GetToastTag(ByVal pSender As LongPtr) As String
    GetToastTag = GetToastStringProp(pSender, VT_IToastNotification2_GetTag)
End Function

Private Function GetToastStringProp(ByVal pSender As LongPtr, ByVal vtIndex As Long) As String
    Dim pToast2 As LongPtr
    Dim iid As WinRT_GUID
    Dim hStr As LongPtr

    If pSender = 0 Then Exit Function
    IIDFromString StrPtr(WinRT_IID_IToastNotification2), iid
    CallComMethod pSender, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pToast2)
    If pToast2 = 0 Then Exit Function

    hStr = 0
    CallComMethod pToast2, vtIndex, vbLong, WinRT_vbPtr, VarPtr(hStr)
    GetToastStringProp = HStringToString(hStr)
    If hStr <> 0 Then WindowsDeleteString hStr
    CallComMethod pToast2, VT_RELEASE, vbLong
End Function

Private Function BuildUserInputDictionary(ByVal pArgs As LongPtr) As Object
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
    CallComMethod pArgs, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidArgs2), WinRT_vbPtr, VarPtr(pActArgs2)
    If pActArgs2 = 0 Then GoTo Done

    CallComMethod pActArgs2, VT_IToastActivatedEventArgs2_GetUserInput, vbLong, WinRT_vbPtr, VarPtr(pPropSet)
    If pPropSet = 0 Then GoTo Done

    IIDFromString StrPtr(WinRT_IID_IIterable_KVP_IInspectable), iidIterable
    CallComMethod pPropSet, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iidIterable), WinRT_vbPtr, VarPtr(pIterable)
    If pIterable = 0 Then GoTo Done

    CallComMethod pIterable, VT_IIterable_First, vbLong, WinRT_vbPtr, VarPtr(pIterator)
    If pIterator = 0 Then GoTo Done

    hasCurrent = 0
    CallComMethod pIterator, VT_IIterator_GetHasCurrent, vbLong, WinRT_vbPtr, VarPtr(hasCurrent)
    Do While hasCurrent <> 0
        pPair = 0
        CallComMethod pIterator, VT_IIterator_GetCurrent, vbLong, WinRT_vbPtr, VarPtr(pPair)
        If pPair <> 0 Then
            hKey = 0
            CallComMethod pPair, VT_IKeyValuePair_GetKey, vbLong, WinRT_vbPtr, VarPtr(hKey)
            keyText = HStringToString(hKey)
            If hKey <> 0 Then WindowsDeleteString hKey

            pValue = 0
            CallComMethod pPair, VT_IKeyValuePair_GetValue, vbLong, WinRT_vbPtr, VarPtr(pValue)
            If Len(keyText) > 0 Then dict(keyText) = InspectableToString(pValue)
            If pValue <> 0 Then CallComMethod pValue, VT_RELEASE, vbLong
            CallComMethod pPair, VT_RELEASE, vbLong
        End If

        hasCurrent = 0
        CallComMethod pIterator, VT_IIterator_MoveNext, vbLong, WinRT_vbPtr, VarPtr(hasCurrent)
    Loop

Done:
    On Error Resume Next
    If pIterator <> 0 Then CallComMethod pIterator, VT_RELEASE, vbLong
    If pIterable <> 0 Then CallComMethod pIterable, VT_RELEASE, vbLong
    If pPropSet <> 0 Then CallComMethod pPropSet, VT_RELEASE, vbLong
    If pActArgs2 <> 0 Then CallComMethod pActArgs2, VT_RELEASE, vbLong
    On Error GoTo 0
    Set BuildUserInputDictionary = dict
End Function

Private Function InspectableToString(ByVal pValue As LongPtr) As String
    Dim pPropVal As LongPtr
    Dim pStringable As LongPtr
    Dim iid As WinRT_GUID
    Dim hStr As LongPtr

    If pValue = 0 Then Exit Function

    IIDFromString StrPtr(WinRT_IID_IPropertyValue), iid
    On Error Resume Next
    CallComMethod pValue, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pPropVal)
    On Error GoTo 0
    If pPropVal <> 0 Then
        hStr = 0
        On Error Resume Next
        CallComMethod pPropVal, VT_IPropertyValue_GetString, vbLong, WinRT_vbPtr, VarPtr(hStr)
        On Error GoTo 0
        InspectableToString = HStringToString(hStr)
        If hStr <> 0 Then WindowsDeleteString hStr
        CallComMethod pPropVal, VT_RELEASE, vbLong
        If Len(InspectableToString) > 0 Then Exit Function
    End If

    IIDFromString StrPtr(WinRT_IID_IStringable), iid
    On Error Resume Next
    CallComMethod pValue, VT_QI, vbLong, WinRT_vbPtr, VarPtr(iid), WinRT_vbPtr, VarPtr(pStringable)
    On Error GoTo 0
    If pStringable <> 0 Then
        hStr = 0
        On Error Resume Next
        CallComMethod pStringable, VT_IStringable_ToString, vbLong, WinRT_vbPtr, VarPtr(hStr)
        On Error GoTo 0
        InspectableToString = HStringToString(hStr)
        If hStr <> 0 Then WindowsDeleteString hStr
        CallComMethod pStringable, VT_RELEASE, vbLong
    End If
End Function


' =============================================================================
' マクロ実行
' =============================================================================

Private Sub RunExcelMacro(ByVal groupInfo As String, ByVal macroName As String, ByVal userInputs As Object)
    Dim bookName As String
    Dim targetPid As Long
    Dim qualifiedName As String

    ParseGroupInfo groupInfo, bookName, targetPid

    If targetPid <> 0 And targetPid <> GetCurrentProcessId() Then Exit Sub

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

Private Sub ParseGroupInfo(ByVal groupInfo As String, ByRef bookName As String, ByRef targetPid As Long)
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


' =============================================================================
' COM / 文字列ユーティリティ
' =============================================================================

Private Function HStringToString(ByVal hStr As LongPtr) As String
    Dim bufLen As Long
    Dim pWchar As LongPtr

    If hStr = 0 Then Exit Function
    pWchar = WindowsGetStringRawBuffer(hStr, bufLen)
    If pWchar = 0 Or bufLen <= 0 Then Exit Function
    HStringToString = String$(bufLen, vbNullChar)
    RtlMoveMemory ByVal StrPtr(HStringToString), ByVal pWchar, CLngPtr(bufLen) * 2
End Function

Private Function CallComMethod( _
    ByVal pUnk As LongPtr, ByVal vTableIndex As Long, _
    ByVal retType As Integer, ParamArray args() As Variant) As Variant

    Dim vTableOffset As LongPtr
    Dim argTypes() As Integer
    Dim argPointers() As LongPtr
    Dim argValues() As Variant
    Dim argCount As Long
    Dim i As Long
    Dim hRes As Long
    Dim vResult As Variant

    If pUnk = 0 Then Err.Raise 513, "CallComMethod", "COM pointer is null. Index=" & vTableIndex

    vTableOffset = vTableIndex * 8

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

    If hRes <> 0 Then Err.Raise hRes, "CallComMethod", "DispCallFunc failed at vtable index " & vTableIndex
    If vTableIndex <> VT_RELEASE Then
        Select Case VarType(vResult)
            Case vbLong, vbInteger, vbByte
                If CLng(vResult) < 0 Then Err.Raise CLng(vResult), "CallComMethod", "COM failed at index " & vTableIndex
            Case vbLongLong
                If vResult < 0 Then Err.Raise CLng(vResult), "CallComMethod", "COM failed at index " & vTableIndex
        End Select
    End If
    CallComMethod = vResult
End Function
