VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToastCallbackCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' ToastCallbackCore.cls
' トースト COM デリゲートの vtable / x64 サンク管理
' 関数アドレスは ITypeInfo (BuildFuncPtrCache) で取得（固定 slot より安全）
' VirtualAlloc 領域への書込は MemLongPtr（起動時のみ・Defender 対策）
' COM コールバック内の GUID 読取は CopyMemory（PointerAccessor は VBE 非安全）
' =============================================================================
#If Win64 Then
    Private Const WinRT_vbPtr As Integer = vbLongLong
    Private Const vbLongPtr As Integer = vbLongLong
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize As Long = 8
#Else
    #Error "64 ビット VBA 専用"
#End If

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Long, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Any) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal length As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" ( _
    ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal bstr As LongPtr)

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const CC_STDCALL As Long = 4

Private Const VT_IToastNotification_AddActivated As Long = 11
Private Const VT_IToastNotification_AddDismissed As Long = 9
Private Const VT_IToastNotification_AddFailed As Long = 13

Private Const WRT_DELEGATE_ACTIVATED As Long = 0
Private Const WRT_DELEGATE_DISMISSED As Long = 1
Private Const WRT_DELEGATE_FAILED As Long = 2
Private Const WRT_DELEGATE_COUNT As Long = 3

Private Const THUNK_SIZE As Long = 56
Private Const DELEGATE_BLK_SIZE As Long = 96
Private Const FAKE_OBJ_SIZE As Long = 8
Private Const MEM_POOL_SIZE As Long = 512

Private Const WinRT_IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
Private Const WinRT_IID_TypedEventHandler_Activated As String = "{AB54DE2D-97D9-5528-B6AD-105AFE156530}"
Private Const WinRT_IID_TypedEventHandler_Dismissed As String = "{61C2402F-0ED0-5A18-AB69-59F4AA99A368}"
Private Const WinRT_IID_TypedEventHandler_Failed As String = "{95E3E803-C969-5E3A-9753-EA2AD22A9A33}"

Private Const WRT_S_OK As Long = 0
Private Const WRT_E_NOINTERFACE As Long = &H80004002

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Type WinRT_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TYPEATTR
    guid As guid
    lcid As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    lpstrSchema As LongPtr
    cbSizeInstance As Long
    typekind As Long
    cFuncs As Integer
    cVars As Integer
    cImplTypes As Integer
    cbSizeVft As Integer
    cbAlignment As Integer
    wTypeFlags As Integer
    wMajorVerNum As Integer
    wMinorVerNum As Integer
    tdescAlias As LongPtr
    tdescAlias2 As LongPtr
    idldescType As LongPtr
    idldescType2 As LongPtr
End Type

Private Type FUNCDESC
    memid As Long
    padding1 As Long
    lprgscode As LongPtr
    lprgelemdescParam As LongPtr
    funckind As Long
    invkind As Long
    callconv As Long
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer
    cScodes As Integer
    elemdescFunc_tdesc As LongPtr
    elemdescFunc_tdesc2 As LongPtr
    elemdescFunc_idldesc As LongPtr
    elemdescFunc_idldesc2 As LongPtr
    wFuncFlags As Integer
    padding2 As Integer
    padding3 As Long
End Type

Private m_Handler As ToastCallbackHandler
Private m_funcPtrsCore As Object
Private m_funcPtrsHandler As Object

Private m_pMemPool As LongPtr
Private m_pThunkQI As LongPtr
Private m_pThunkAR As LongPtr
Private m_pThunkRelease As LongPtr

Private m_pVtblBlock(0 To 2) As LongPtr
Private m_pObject(0 To 2) As LongPtr
Private m_RefCount(0 To 2) As Long
Private m_iidDelegate(0 To 2) As WinRT_GUID
Private m_iidUnknown As WinRT_GUID
Private m_DelegatesReady As Boolean


Private Sub Class_Initialize()
    Set m_Handler = New ToastCallbackHandler
    Set m_funcPtrsCore = CreateObject("Scripting.Dictionary")
    m_funcPtrsCore.CompareMode = 1
    Set m_funcPtrsHandler = CreateObject("Scripting.Dictionary")
    m_funcPtrsHandler.CompareMode = 1
End Sub

Private Sub Class_Terminate()
    If m_pMemPool <> 0 Then
        VirtualFree m_pMemPool, 0, MEM_RELEASE
        m_pMemPool = 0
    End If
    m_pThunkQI = 0
    m_pThunkAR = 0
    m_pThunkRelease = 0
    Dim i As Long
    For i = 0 To WRT_DELEGATE_COUNT - 1
        m_pObject(i) = 0
        m_pVtblBlock(i) = 0
    Next i
    Set m_Handler = Nothing
    Set m_funcPtrsCore = Nothing
    Set m_funcPtrsHandler = Nothing
End Sub


Public Sub PrewarmDelegates()
    EnsureDelegates
End Sub

Public Sub RegisterToastHandlers(ByVal pToast As LongPtr)
    Dim token As LongLong
    Dim i As Long

    If pToast = 0 Then Exit Sub

    On Error Resume Next
    EnsureDelegates
    If Not m_DelegatesReady Then Exit Sub

    For i = 0 To WRT_DELEGATE_COUNT - 1
        token = 0
        Select Case i
            Case WRT_DELEGATE_ACTIVATED
                CallComMethod pToast, VT_IToastNotification_AddActivated, vbLong, _
                    WinRT_vbPtr, m_pObject(i), WinRT_vbPtr, VarPtr(token)
            Case WRT_DELEGATE_DISMISSED
                CallComMethod pToast, VT_IToastNotification_AddDismissed, vbLong, _
                    WinRT_vbPtr, m_pObject(i), WinRT_vbPtr, VarPtr(token)
            Case WRT_DELEGATE_FAILED
                CallComMethod pToast, VT_IToastNotification_AddFailed, vbLong, _
                    WinRT_vbPtr, m_pObject(i), WinRT_vbPtr, VarPtr(token)
        End Select
    Next i
    On Error GoTo 0
End Sub


Public Function Com_QueryInterface( _
    ByVal pThis As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long

    Dim g As WinRT_GUID
    Dim idx As Long

    If riid = 0 Then
        Com_QueryInterface = WRT_E_NOINTERFACE
        Exit Function
    End If

    idx = DelegateIndexFromThis(pThis)
    If idx < 0 Then
        ppvObject = 0
        Com_QueryInterface = WRT_E_NOINTERFACE
        Exit Function
    End If

    CopyMemory g, ByVal riid, LenB(g)
    If GuidEqual(g, m_iidUnknown) Or GuidEqual(g, m_iidDelegate(idx)) Then
        ppvObject = pThis
        m_RefCount(idx) = m_RefCount(idx) + 1
        Com_QueryInterface = WRT_S_OK
    Else
        ppvObject = 0
        Com_QueryInterface = WRT_E_NOINTERFACE
    End If
End Function

Public Function Com_AddRef(ByVal pThis As LongPtr) As Long
    Dim idx As Long
    idx = DelegateIndexFromThis(pThis)
    If idx < 0 Then
        Com_AddRef = 1
        Exit Function
    End If
    m_RefCount(idx) = m_RefCount(idx) + 1
    Com_AddRef = m_RefCount(idx)
End Function

Public Function Com_Release(ByVal pThis As LongPtr) As Long
    Dim idx As Long
    idx = DelegateIndexFromThis(pThis)
    If idx < 0 Then
        Com_Release = 1
        Exit Function
    End If
    If m_RefCount(idx) > 0 Then m_RefCount(idx) = m_RefCount(idx) - 1
    Com_Release = m_RefCount(idx)
End Function


Private Sub EnsureDelegates()
    Dim i As Long
    Dim invokeName As String
    Dim poolBase As LongPtr
    Dim off As LongPtr

    If m_DelegatesReady Then Exit Sub

    If m_funcPtrsCore.Count = 0 Then BuildFuncPtrCache ObjPtr(Me), m_funcPtrsCore
    If m_funcPtrsHandler.Count = 0 Then BuildFuncPtrCache ObjPtr(m_Handler), m_funcPtrsHandler
    If m_funcPtrsCore.Count = 0 Or m_funcPtrsHandler.Count = 0 Then Exit Sub

    IIDFromString StrPtr(WinRT_IID_IUnknown), m_iidUnknown
    IIDFromString StrPtr(WinRT_IID_TypedEventHandler_Activated), m_iidDelegate(WRT_DELEGATE_ACTIVATED)
    IIDFromString StrPtr(WinRT_IID_TypedEventHandler_Dismissed), m_iidDelegate(WRT_DELEGATE_DISMISSED)
    IIDFromString StrPtr(WinRT_IID_TypedEventHandler_Failed), m_iidDelegate(WRT_DELEGATE_FAILED)

    poolBase = VirtualAlloc(0, MEM_POOL_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If poolBase = 0 Then Exit Sub
    m_pMemPool = poolBase

    off = poolBase
    m_pThunkQI = off
    If Not WriteThunkCall3At(m_pThunkQI, ObjPtr(Me), m_funcPtrsCore("Com_QueryInterface")) Then GoTo Fail
    off = off + THUNK_SIZE

    m_pThunkAR = off
    If Not WriteThunkCall1At(m_pThunkAR, ObjPtr(Me), m_funcPtrsCore("Com_AddRef")) Then GoTo Fail
    off = off + THUNK_SIZE

    m_pThunkRelease = off
    If Not WriteThunkCall1At(m_pThunkRelease, ObjPtr(Me), m_funcPtrsCore("Com_Release")) Then GoTo Fail
    off = off + THUNK_SIZE

    For i = 0 To WRT_DELEGATE_COUNT - 1
        Select Case i
            Case WRT_DELEGATE_ACTIVATED: invokeName = "Invoke_Activated"
            Case WRT_DELEGATE_DISMISSED: invokeName = "Invoke_Dismissed"
            Case WRT_DELEGATE_FAILED:    invokeName = "Invoke_Failed"
        End Select
        If Not m_funcPtrsHandler.Exists(invokeName) Then GoTo Fail

        m_pVtblBlock(i) = off
        If Not BuildDelegateBlockAt(m_pVtblBlock(i), off + DELEGATE_BLK_SIZE, m_funcPtrsHandler(invokeName)) Then GoTo Fail
        m_pObject(i) = off + DELEGATE_BLK_SIZE
        off = off + DELEGATE_BLK_SIZE + FAKE_OBJ_SIZE
        m_RefCount(i) = 1
    Next i

    m_DelegatesReady = True
    Exit Sub

Fail:
    If m_pMemPool <> 0 Then
        VirtualFree m_pMemPool, 0, MEM_RELEASE
        m_pMemPool = 0
    End If
    m_pThunkQI = 0
    m_pThunkAR = 0
    m_pThunkRelease = 0
End Sub

Private Function BuildDelegateBlockAt( _
    ByVal blk As LongPtr, ByVal fakeObj As LongPtr, ByVal pInvokeFunc As LongPtr) As Boolean

    Dim thkInvoke As LongPtr
    Dim slot(3) As LongPtr

    thkInvoke = blk + 32
    If Not WriteThunkCall3At(thkInvoke, ObjPtr(m_Handler), pInvokeFunc) Then Exit Function

    slot(0) = m_pThunkQI
    slot(1) = m_pThunkAR
    slot(2) = m_pThunkRelease
    slot(3) = thkInvoke
    MemLongPtr(blk + 0) = slot(0)
    MemLongPtr(blk + 8) = slot(1)
    MemLongPtr(blk + 16) = slot(2)
    MemLongPtr(blk + 24) = slot(3)
    MemLongPtr(fakeObj) = blk

    BuildDelegateBlockAt = True
End Function

Private Function DelegateIndexFromThis(ByVal this As LongPtr) As Long
    Dim i As Long
    For i = 0 To WRT_DELEGATE_COUNT - 1
        If m_pObject(i) = this Then
            DelegateIndexFromThis = i
            Exit Function
        End If
    Next i
    DelegateIndexFromThis = -1
End Function


Private Function WriteThunkCall1At( _
    ByVal addr As LongPtr, ByVal pOwner As LongPtr, ByVal pFunc As LongPtr) As Boolean

    Dim b(0 To THUNK_SIZE - 1) As Byte
    Dim i As Long
    Dim k As Long

    If addr = 0 Or pFunc = 0 Then Exit Function

    i = 0
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H28: i = i + 4
    b(i) = &H4C: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H10: i = i + 5
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pOwner: i = i + 8
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pFunc: i = i + 8
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H10: i = i + 4
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H28: i = i + 4
    b(i) = &HC3: i = i + 1
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    For k = 0 To THUNK_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
    WriteThunkCall1At = True
End Function

Private Function WriteThunkCall3At( _
    ByVal addr As LongPtr, ByVal pOwner As LongPtr, ByVal pFunc As LongPtr) As Boolean

    Dim b(0 To THUNK_SIZE - 1) As Byte
    Dim i As Long
    Dim k As Long

    If addr = 0 Or pFunc = 0 Then Exit Function

    i = 0
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pOwner: i = i + 8
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pFunc: i = i + 8
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    b(i) = &HC3: i = i + 1
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    For k = 0 To THUNK_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
    WriteThunkCall3At = True
End Function


Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function

Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pA(0 To 0) As PointerAccessor
    With pA(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pA, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pA(0 To 0) As PointerAccessor
    With pA(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pA, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub


Private Sub BuildFuncPtrCache(ByVal pObj As LongPtr, ByVal dict As Object)
    dict.RemoveAll
    Dim pTI As LongPtr
    Dim hr As Long
    hr = CallCOM3(pObj, 4, 0&, 0&, VarPtr(pTI))
    If hr <> 0 Then Exit Sub

    Dim pTA As LongPtr
    hr = CallCOM1(pTI, 3, VarPtr(pTA))
    If hr <> 0 Then GoTo Cleanup

    Dim uTA As TYPEATTR
    CopyMemory uTA, ByVal pTA, LenB(uTA)
    Dim pVT As LongPtr
    CopyMemory pVT, ByVal pObj, LenB(pVT)

    Dim n As Long
    For n = 0 To uTA.cFuncs - 1
        Dim pFD As LongPtr
        hr = CallCOM2(pTI, 5, CLng(n), VarPtr(pFD))
        If hr <> 0 Then GoTo NextF
        Dim uFD As FUNCDESC
        CopyMemory uFD, ByVal pFD, LenB(uFD)
        Dim pB As LongPtr
        pB = 0
        hr = CallCOM5(pTI, 12, uFD.memid, VarPtr(pB), 0, 0, 0)
        If hr = 0 And pB <> 0 Then
            Dim fn As String
            fn = PtrToStrW(pB)
            SysFreeString pB
            pB = 0
            Dim pA As LongPtr
            CopyMemory pA, ByVal (pVT + uFD.oVft), LenB(pA)
            dict(fn) = pA
        End If
        CallCOM1V pTI, 20, pFD
NextF:
    Next n
    CallCOM1V pTI, 19, pTA
Cleanup:
    CallCOM0V pTI, 2
End Sub

Private Function PtrToStrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToStrW = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToStrW), p, CLngPtr(cch * 2)
End Function


Private Function GuidEqual(ByRef a As WinRT_GUID, ByRef b As WinRT_GUID) As Boolean
    Dim i As Long
    If a.Data1 <> b.Data1 Then Exit Function
    If a.Data2 <> b.Data2 Then Exit Function
    If a.Data3 <> b.Data3 Then Exit Function
    For i = 0 To 7
        If a.Data4(i) <> b.Data4(i) Then Exit Function
    Next i
    GuidEqual = True
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

    If pUnk = 0 Then Exit Function

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

    If hRes = 0 Then CallComMethod = vResult
End Function


Private Function CallCOM0V(ByVal p As LongPtr, ByVal idx As Long) As Long
    Dim r As Variant
    CallCOM0V = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 0, 0, 0, r)
End Function

Private Function CallCOM1(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr) As Long
    Dim r As Variant
    Dim v(0) As Variant
    Dim t(0) As Integer
    Dim vp(0) As LongPtr
    v(0) = a1: t(0) = vbLongPtr: vp(0) = VarPtr(v(0))
    CallCOM1 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), r)
    If CallCOM1 = 0 Then CallCOM1 = CLng(r)
End Function

Private Sub CallCOM1V(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr)
    Dim r As Variant
    Dim v(0) As Variant
    Dim t(0) As Integer
    Dim vp(0) As LongPtr
    v(0) = a1: t(0) = vbLongPtr: vp(0) = VarPtr(v(0))
    DispCallFunc p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), r
End Sub

Private Function CallCOM2(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr) As Long

    Dim r As Variant
    Dim v(1) As Variant
    Dim t(1) As Integer
    Dim vp(1) As LongPtr
    v(0) = a1: v(1) = a2: t(0) = vbLong: t(1) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1))
    CallCOM2 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 2, t(0), vp(0), r)
    If CallCOM2 = 0 Then CallCOM2 = CLng(r)
End Function

Private Function CallCOM3(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As LongPtr) As Long

    Dim r As Variant
    Dim v(2) As Variant
    Dim t(2) As Integer
    Dim vp(2) As LongPtr
    v(0) = a1: v(1) = a2: v(2) = a3
    t(0) = vbLong: t(1) = vbLong: t(2) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    CallCOM3 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 3, t(0), vp(0), r)
    If CallCOM3 = 0 Then CallCOM3 = CLng(r)
End Function

Private Function CallCOM5(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr, ByVal a3 As LongPtr, _
    ByVal a4 As LongPtr, ByVal a5 As LongPtr) As Long

    Dim r As Variant
    Dim v(4) As Variant
    Dim t(4) As Integer
    Dim vp(4) As LongPtr
    v(0) = a1: v(1) = a2: v(2) = a3: v(3) = a4: v(4) = a5
    t(0) = vbLong: t(1) = vbLongPtr: t(2) = vbLongPtr: t(3) = vbLongPtr: t(4) = vbLongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    vp(3) = VarPtr(v(3)): vp(4) = VarPtr(v(4))
    CallCOM5 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 5, t(0), vp(0), r)
    If CallCOM5 = 0 Then CallCOM5 = CLng(r)
End Function
