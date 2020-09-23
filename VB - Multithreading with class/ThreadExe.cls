VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThreadExe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Thread stuff
Private Const CREATE_SUSPENDED As Long = &H4

Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'Synchronization stuff
Private Const INFINITE As Long = &HFFFFFFFF
Private Const WAIT_TIMEOUT As Long = 258&
Private Const STATUS_WAIT_0 As Long = &H0
Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_OBJECT_1 As Long = (WAIT_OBJECT_0 + 1)

Private Const STATUS_PENDING As Long = &H103
Private Const STILL_ACTIVE As Long = STATUS_PENDING

Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'ShellExecute stuff
Private Const SEE_MASK_FLAG_DDEWAIT As Long = &H100
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMINIMIZED As Long = 2
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_MAXIMIZE As Long = 3
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const SW_SHOW As Long = 5
Private Const SW_MINIMIZE As Long = 6
Private Const SW_SHOWMINNOACTIVE As Long = 7
Private Const SW_SHOWNA As Long = 8
Private Const SW_RESTORE As Long = 9
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SW_FORCEMINIMIZE As Long = 11


Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpstrVerb As Long
    lpstrFile As Long
    lpstrParameters As Long
    lpstrDirectory As Long
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpstrClass As Long
    hkeyClass As Long
    dwHotKey As Long
    hIconOrMon As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

'Misc stuff
Private Declare Sub GetArrayDescriptor Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc() As Any, ByVal ByteLen As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long


'Public enum for the showmode wanted
Public Enum teShowModeEnum
    teHide = SW_HIDE
    teShowNormal = SW_SHOWNORMAL
    teShowMinimized = SW_SHOWMINIMIZED
    teShowMaximized = SW_SHOWMAXIMIZED
    teShowNoActivate = SW_SHOWNOACTIVATE
    teShow = SW_SHOW
    teMinimize = SW_MINIMIZE
    teShowMinNoActive = SW_SHOWMINNOACTIVE
    teShowNA = SW_SHOWNA
    teRestore = SW_RESTORE
    teShowDefault = SW_SHOWDEFAULT
    teForceMinimize = SW_FORCEMINIMIZE
End Enum
#If False Then
    Private teHide, teShowNormal, teShowMinimized, teShowMaximized, teShowNoActivate
    Private teShow, teMinimize, teShowMinNoActive, teShowNA, teRestore
    Private teShowDefault, teForceMinimize
#End If

'Public enum for Status property
Public Enum teProcessStatusEnum
    teNoProcess
    teProcessRunning
    teDisconnected
End Enum
#If False Then
    Private teNoProcess, teProcessRunning, teDisconnected
#End If

'Member variables
Private g_MemberVarsCS As CRITICAL_SECTION 'Critical section for member variables
Private m_arrAction() As Byte 'Array containing ANSI version of Action
Private m_arrExeName() As Byte 'Array containing ANSI version of ExeName
Private m_arrDirectory() As Byte 'Array containing ANSI version of Directory
Private m_arrParameters() As Byte 'Array containing ANSI version of Parameters
Private m_ParentWindow As Form 'Parent form that receives system messages while calling ShellExecuteEx
Private m_ShowMode As teShowModeEnum 'Show mode to use in ShellExecuteEx
Private m_LastShellError As Long 'Last error returned by ShellExecuteEx
Private m_hProcess As Long 'Current process handle
Private m_LastExitCode As Long 'Last process exit code for the process being monitored
Private m_WorkerThreadID As Long 'The thread ID of the monitor thread
Private m_hThread As Long 'Monitor thread handle
'If the next flag is true, then there is a thread handle in m_hThread, a process
'handle in m_hProcess and therefore, this object is busy monitoring a process
Private bProcessOnMonitor As Boolean 'Flag indicating existance of a monitored process

'Control variables
'Event used to signal a change in the control variables
Private hEventAccessControlVars As Long 'Event for signaling the wait thread that the control variables have changed
Private g_ControlVarsCS As CRITICAL_SECTION 'Critical section for control variables
Private m_EventSink As IThreadEvents 'Event sink object (an object that implements IThreadEvents)
Private bNoEvents As Boolean 'A boolean value indicating the monitor thread to not call the event sink's methods
Private bKillProcess As Boolean 'A boolean value indicating the monitor thread to kill the monitored process
Private bStopMonitor As Boolean 'A boolean value indicating the monitor thread to stop monitoring

Public Property Get Action() As String
    EnterCriticalSection g_MemberVarsCS
    Action = StrConv(m_arrAction, vbUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Let Action(ByVal vNewValue As String)
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 419
    End If
    Erase m_arrAction
    m_arrAction() = StrConv(vNewValue & vbNullChar, vbFromUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get ExeName() As String
    EnterCriticalSection g_MemberVarsCS
    ExeName = StrConv(m_arrExeName, vbUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Let ExeName(ByVal vNewValue As String)
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 419
    End If
    Erase m_arrExeName
    m_arrExeName() = StrConv(vNewValue & vbNullChar, vbFromUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Function StartProcess() As Boolean

Dim sei As SHELLEXECUTEINFO

    EnterCriticalSection g_ControlVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_ControlVarsCS
        Err.Raise 419
    End If
    LeaveCriticalSection g_ControlVarsCS
    EnterCriticalSection g_MemberVarsCS
    If (GetStringAddress(m_arrExeName) = 0) Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 380
    End If
    If (m_hProcess <> 0) Then
        CloseHandle m_hProcess
        m_hProcess = 0
    End If
    If (m_hThread <> 0) Then
        CloseHandle m_hThread
        m_hThread = 0
    End If
    With sei
        .cbSize = Len(sei)
        .fMask = SEE_MASK_FLAG_DDEWAIT Or SEE_MASK_NOCLOSEPROCESS
        .hWnd = m_ParentWindow.hWnd
        .lpstrDirectory = GetStringAddress(m_arrDirectory)
        .lpstrFile = GetStringAddress(m_arrExeName)
        .lpstrParameters = GetStringAddress(m_arrParameters)
        .lpstrVerb = GetStringAddress(m_arrAction)
        .nShow = m_ShowMode
    End With
    If (ShellExecuteEx(sei) = 0) Then
        StartProcess = False
        m_LastShellError = GetLastError
        LeaveCriticalSection g_MemberVarsCS
        Exit Function
    End If
'    LeaveCriticalSection g_MemberVarsCS
'    EnterCriticalSection g_MemberVarsCS
'    m_WorkerThreadID = 0
    If (sei.hProcess <> 0) And (sei.hInstApp > 32) Then
        m_hProcess = sei.hProcess
    End If
    LeaveCriticalSection g_MemberVarsCS
    StartProcess = LaunchMonitor
End Function

Public Property Get ParentWindow() As Form
    EnterCriticalSection g_MemberVarsCS
    Set ParentWindow = m_ParentWindow
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Set ParentWindow(ByVal vNewValue As Form)
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 419
    End If
    Set m_ParentWindow = vNewValue
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get Directory() As String
    EnterCriticalSection g_MemberVarsCS
    Directory = StrConv(m_arrDirectory, vbUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Let Directory(ByVal vNewValue As String)
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 419
    End If
    Erase m_arrDirectory
    m_arrDirectory() = StrConv(vNewValue & vbNullChar, vbFromUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get Parameters() As String
    EnterCriticalSection g_MemberVarsCS
    Parameters = StrConv(m_arrParameters, vbUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Let Parameters(ByVal vNewValue As String)
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 419
    End If
    Erase m_arrParameters
    m_arrParameters() = StrConv(vNewValue & vbNullChar, vbFromUnicode)
    LeaveCriticalSection g_MemberVarsCS
End Property

Private Function GetStringAddress(ByRef arrStr() As Byte) As Long

Dim lDesc As Long

    GetArrayDescriptor lDesc, arrStr, 4
    If (lDesc <> 0) Then
        lDesc = VarPtr(arrStr(LBound(arrStr)))
    End If
    GetStringAddress = lDesc
End Function

Public Property Get ShowMode() As teShowModeEnum
    EnterCriticalSection g_MemberVarsCS
    ShowMode = m_ShowMode
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Let ShowMode(ByVal vNewValue As teShowModeEnum)
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 419
    End If
    m_ShowMode = vNewValue
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get ReturnCode() As Long
    EnterCriticalSection g_MemberVarsCS
    If (m_hProcess = 0) Then
        ReturnCode = m_LastExitCode
    Else
        GetExitCodeProcess m_hProcess, ReturnCode
    End If
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get LastShellError() As Long
    EnterCriticalSection g_MemberVarsCS
    LastShellError = m_LastShellError
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Sub StopMonitor(Optional ByVal bStopProcess As Boolean = False)
    EnterCriticalSection g_MemberVarsCS
    If (m_hThread = 0) Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 5
    End If
    'The following IF ensures that this method will not be called from the worker
    'thread context because this methods waits on the worker thread.
    If (GetCurrentThreadId = m_WorkerThreadID) Then
        LeaveCriticalSection g_MemberVarsCS
        Err.Raise 5
    End If
    m_WorkerThreadID = 0
    LeaveCriticalSection g_MemberVarsCS
    EnterCriticalSection g_ControlVarsCS
    bKillProcess = bStopProcess
    bNoEvents = True
    bStopMonitor = True
    LeaveCriticalSection g_ControlVarsCS
    SetEvent hEventAccessControlVars
    WaitForSingleObject m_hThread, INFINITE
    EnterCriticalSection g_MemberVarsCS
    CloseHandle m_hThread
    m_hThread = 0
    LeaveCriticalSection g_MemberVarsCS
    EnterCriticalSection g_ControlVarsCS
    bNoEvents = False
    bStopMonitor = False
    bKillProcess = False
    RaiseStatusEvent
    LeaveCriticalSection g_ControlVarsCS
End Sub

Private Sub Class_Initialize()
    InitializeCriticalSection g_MemberVarsCS
    InitializeCriticalSection g_ControlVarsCS
    hEventAccessControlVars = CreateEvent(0, 0, 0, "_MemberVars_" & CStr(ObjPtr(Me)))
End Sub

Private Sub Class_Terminate()
    StopMonitor
    CloseProcessHandle
    DeleteCriticalSection g_MemberVarsCS
    DeleteCriticalSection g_ControlVarsCS
    CloseHandle hEventAccessControlVars
End Sub

Private Sub CloseProcessHandle()
    EnterCriticalSection g_MemberVarsCS
    If (m_hProcess <> 0) Then
        CloseHandle m_hProcess
        m_hProcess = 0
    End If
    LeaveCriticalSection g_MemberVarsCS
End Sub

Public Property Get ProcessStatus() As teProcessStatusEnum
    EnterCriticalSection g_MemberVarsCS
    If bProcessOnMonitor Then
        ProcessStatus = teProcessRunning
    Else
        ProcessStatus = IIf(ReturnCode = STILL_ACTIVE, teDisconnected, teNoProcess)
    End If
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get EventSink() As IThreadEvents
    EnterCriticalSection g_MemberVarsCS
    Set EventSink = m_EventSink
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Set EventSink(ByVal vNewValue As IThreadEvents)
    EnterCriticalSection g_MemberVarsCS
    Set m_EventSink = vNewValue
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Property Get WorkerThreadID() As Long
    EnterCriticalSection g_MemberVarsCS
    WorkerThreadID = m_WorkerThreadID
    LeaveCriticalSection g_MemberVarsCS
End Property

Public Sub StartMonitor()

Dim lResp As Long
Dim bExit As Boolean
Dim lExitCode As Long
Dim arrObjects(1 To 2) As Long

    EnterCriticalSection g_MemberVarsCS
    bExit = (m_hProcess <> 0)
    LeaveCriticalSection g_MemberVarsCS
    If bExit Then
        EnterCriticalSection g_MemberVarsCS
        bProcessOnMonitor = True
        arrObjects(1) = hEventAccessControlVars
        arrObjects(2) = m_hProcess
        LeaveCriticalSection g_MemberVarsCS
        RaiseStatusEvent
        Do
            lResp = WaitForMultipleObjects(2, arrObjects(1), 0, INFINITE)
            If (lResp = WAIT_OBJECT_0) Then
                EnterCriticalSection g_ControlVarsCS
                bExit = bStopMonitor Or bKillProcess
                If bKillProcess Then
                    TerminateProcess arrObjects(2), 0
                    EnterCriticalSection g_MemberVarsCS
'                    GetExitCodeProcess m_hProcess, lExitCode
'                    m_LastExitCode = lExitCode
                    m_LastExitCode = 0
                    CloseHandle m_hProcess
                    m_hProcess = 0
                    bProcessOnMonitor = False
                    LeaveCriticalSection g_MemberVarsCS
                    RaiseStatusEvent
                End If
                LeaveCriticalSection g_ControlVarsCS
            ElseIf (lResp = WAIT_OBJECT_1) Then
                EnterCriticalSection g_ControlVarsCS
                EnterCriticalSection g_MemberVarsCS
                bProcessOnMonitor = False
                LeaveCriticalSection g_MemberVarsCS
                RaiseStatusEvent
                LeaveCriticalSection g_ControlVarsCS
                bExit = True
            End If
        Loop Until bExit
        EnterCriticalSection g_MemberVarsCS
        bProcessOnMonitor = False
        LeaveCriticalSection g_MemberVarsCS
    End If
End Sub

Public Function LaunchMonitor() As Boolean

Dim lThreadID As Long
Dim hThreadExe As Long

    If (m_hThread <> 0) Then
        Err.Raise 5
    End If
    hThreadExe = CreateThread(0, 0, AddressOf ThreadExeStarter, ByVal ObjPtr(Me), CREATE_SUSPENDED, lThreadID)
    If (hThreadExe = 0) Then
        LaunchMonitor = False
        Exit Function
    End If
    EnterCriticalSection g_MemberVarsCS
    m_hThread = hThreadExe
    m_WorkerThreadID = lThreadID
    LeaveCriticalSection g_MemberVarsCS
    ResumeThread hThreadExe
    LaunchMonitor = True
End Function

Private Sub RaiseStatusEvent()
    EnterCriticalSection g_ControlVarsCS
    If Not (bNoEvents) Then
        If Not (m_EventSink Is Nothing) Then
            m_EventSink.StatusChanged Me
        End If
    End If
    LeaveCriticalSection g_ControlVarsCS
End Sub
