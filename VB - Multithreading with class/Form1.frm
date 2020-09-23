VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch and monitor process"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList illvProcesses 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   "Stopped"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0894
            Key             =   "Unplugged"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CE6
            Key             =   "Running"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cbResumeMon 
      Caption         =   "Resume monitor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton cbDelObject 
      Caption         =   "Delete controller"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton cbStopMon 
      Caption         =   "Stop monitor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton cbStopProc 
      Caption         =   "Stop process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   5040
      Width           =   2655
   End
   Begin MSComctlLib.ListView lvProcesses 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2778
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      SmallIcons      =   "illvProcesses"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ExeName"
         Text            =   "Name"
         Object.Width           =   6809
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Status"
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "ExitCode"
         Text            =   "Exit code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "ThreadID"
         Text            =   "Thread ID"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   3840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frProcess 
      Caption         =   "Process to launch"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   7455
      Begin VB.CheckBox ckUseCurrent 
         Caption         =   "Use current controller"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cbStart 
         Caption         =   "Start process"
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox ddAction 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton cbBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   6960
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox tbFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\Winnt\notepad.exe"
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "Action:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "File name:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image imgInfo 
      Height          =   480
      Left            =   120
      Top             =   600
      Width           =   480
   End
   Begin VB.Label etNotes 
      Caption         =   $"Form1.frx":1138
      Height          =   1095
      Left            =   840
      TabIndex        =   14
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label etMainThread 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Const IMAGE_ICON As Long = 1
Private Const LR_DEFAULTCOLOR As Long = &H0
Private Const LR_SHARED As Long = &H8000
Private Const OIC_NOTE As Long = 32516
Private Const OIC_INFORMATION As Long = OIC_NOTE
Private Const PICTYPE_ICON As Long = 3

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PictDesc, ByVal riid As Long, ByVal fOwn As Long, lplpvObj As Any) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private IID_IPicture As Guid


Private colObjects As Collection

Private g_ListViewCS As CRITICAL_SECTION
Private g_LViewCtlsCS As CRITICAL_SECTION

Implements IThreadEvents

Private Sub cbBrowse_Click()
    With cdDialog
        .CancelError = True
        .DialogTitle = "Select document to shell"
        .filename = tbFilename.Text
        .Filter = "Executable files (*.exe)|*.exe|Office documents (*.xls, *.doc, *.ppt)|*.xls; *.doc; *.ppt|All files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly
        On Error Resume Next
        .ShowOpen
        If (Err.Number <> 0) Then
            Exit Sub
        End If
        tbFilename.Text = .filename
    End With
End Sub

Private Sub cbDelObject_Click()

Dim oThread As ThreadExe

    Set oThread = GetSelectedThreadObject
    If (oThread Is Nothing) Then
        Exit Sub
    End If
    If (oThread.ProcessStatus = teProcessRunning) Then
        oThread.StopMonitor
    End If
    ZeroMemory oThread, 4
    EnterCriticalSection g_ListViewCS
    colObjects.Remove lvProcesses.SelectedItem.Key
    With lvProcesses
        .ListItems.Remove .SelectedItem.Index
    End With
    LeaveCriticalSection g_ListViewCS
    UpdateLVControls
End Sub

Private Sub cbResumeMon_Click()

Dim oThread As ThreadExe

    Set oThread = GetSelectedThreadObject
    If (oThread Is Nothing) Then
        Exit Sub
    End If
    If (oThread.ProcessStatus = teDisconnected) Then
        oThread.LaunchMonitor
    End If
    UpdateLVControls
End Sub

Private Sub cbStart_Click()

Dim oThreadObject As ThreadExe

    If (Len(Trim$(tbFilename.Text)) = 0) Then
        MsgBox "You must enter a valid document or program filename.", vbInformation
        Exit Sub
    End If
    Set oThreadObject = GetSelectedThreadObject
    If (ckUseCurrent.Value <> vbChecked) Or (oThreadObject Is Nothing) Then
        Set oThreadObject = New ThreadExe
    End If
    With oThreadObject
        If (.ProcessStatus = teProcessRunning) Then
            .StopMonitor
        End If
        .Action = ddAction.Text
        .ExeName = Trim$(tbFilename.Text)
        Set .ParentWindow = Me
        .ShowMode = teShowNormal
        Set .EventSink = Me
    End With
    AddObjectToList oThreadObject
    oThreadObject.StartProcess
    UpdateLVControls
End Sub

Private Sub cbStopMon_Click()

Dim oThread As ThreadExe

    Set oThread = GetSelectedThreadObject
    If (oThread Is Nothing) Then
        Exit Sub
    End If
    oThread.StopMonitor
    UpdateLVControls
End Sub

Private Sub cbStopProc_Click()

Dim oThread As ThreadExe

    Set oThread = GetSelectedThreadObject
    If (oThread Is Nothing) Then
        Exit Sub
    End If
    oThread.StopMonitor True
    UpdateLVControls
End Sub

Private Sub Form_Load()

Dim arrOptions As Variant
Dim lCount As Long
Dim hPic As Long
Dim pd As PictDesc
Dim lpPicture As Picture

    Set colObjects = New Collection
    InitializeCriticalSection g_ListViewCS
    InitializeCriticalSection g_LViewCtlsCS
    arrOptions = Array("edit", "explore", "find", "open", "print", "properties")
    With ddAction
        For lCount = LBound(arrOptions) To UBound(arrOptions)
            .AddItem arrOptions(lCount)
        Next lCount
        .Text = "open"
    End With
    etMainThread.Caption = "Main thread ID:  " & CStr(GetCurrentThreadId)

    hPic = LoadImage(0, OIC_INFORMATION, IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR Or LR_SHARED)
    If (hPic <> 0) Then
        With IID_IPicture
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(3) = &HAA
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        pd.cbSizeofStruct = Len(pd)
        pd.hImage = hPic
        pd.picType = PICTYPE_ICON
        OleCreatePictureIndirect pd, VarPtr(IID_IPicture), 1, lpPicture
        Set imgInfo.Picture = lpPicture
    End If

    With etNotes
        .Caption = "All users:  The control menu has an Always-on-top item." & _
                vbCrLf & vbCrLf & .Caption
    End With
    AddAlwaysOnTopToSysMenu Me.hWnd
    StartSubclass Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colObjects = Nothing
    DeleteCriticalSection g_ListViewCS
    DeleteCriticalSection g_LViewCtlsCS
    StopSubclass
End Sub

Private Sub AddObjectToList(ByVal oObject As ThreadExe)

Dim li As ListItem

    EnterCriticalSection g_ListViewCS
    On Error Resume Next
    colObjects.Add oObject, "K" & CStr(ObjPtr(oObject))
    If (Err.Number = 457) Then
        Set li = lvProcesses.ListItems("K" & CStr(ObjPtr(oObject)))
    Else
        Set li = lvProcesses.ListItems.Add(, "K" & CStr(ObjPtr(oObject)), oObject.ExeName, , "Stopped")
    End If
    On Error GoTo 0
    li.Text = oObject.ExeName
    li.SubItems(1) = StatusText(oObject.ProcessStatus)
    li.SubItems(2) = ""
    li.SubItems(3) = ""
    LeaveCriticalSection g_ListViewCS
End Sub

Private Sub IThreadEvents_StatusChanged(ByVal oThread As ThreadExe)
    EnterCriticalSection g_ListViewCS
    With lvProcesses.ListItems("K" & CStr(ObjPtr(oThread)))
        .SmallIcon = IconIndex(oThread.ProcessStatus)
        .SubItems(1) = StatusText(oThread.ProcessStatus)
        If (oThread.ProcessStatus = teNoProcess) Then
            .SubItems(2) = oThread.ReturnCode
            .SubItems(3) = "0"
        Else
            .SubItems(3) = CStr(oThread.WorkerThreadID)
        End If
    End With
    UpdateLVControls
    LeaveCriticalSection g_ListViewCS
End Sub

Private Function StatusText(ByVal eStatus As teProcessStatusEnum)
    Select Case eStatus
        Case teNoProcess
            StatusText = "No process"
        Case teProcessRunning
            StatusText = "Running"
        Case teDisconnected
            StatusText = "Unplugged"
    End Select
End Function

Private Function IconIndex(ByVal eStatus As teProcessStatusEnum) As Long
    Select Case eStatus
        Case teNoProcess
            IconIndex = 1
        Case teProcessRunning
            IconIndex = 3
        Case teDisconnected
            IconIndex = 2
    End Select
End Function

Private Function GetSelectedThreadObject() As ThreadExe

Dim lObject As Long
Dim oThread As ThreadExe

    EnterCriticalSection g_ListViewCS
    If (lvProcesses.SelectedItem Is Nothing) Then
        LeaveCriticalSection g_ListViewCS
        Set GetSelectedThreadObject = Nothing
        Exit Function
    End If
    lObject = CLng(Mid(lvProcesses.SelectedItem.Key, 2))
    LeaveCriticalSection g_ListViewCS
    CopyMemory oThread, lObject, 4
    Set GetSelectedThreadObject = oThread
    ZeroMemory oThread, 4
End Function

Private Sub lvProcesses_Click()
    UpdateLVControls
End Sub

Private Sub lvProcesses_ItemClick(ByVal Item As MSComctlLib.ListItem)
    UpdateLVControls
End Sub

Private Sub UpdateLVControls()

Dim oThread As ThreadExe

    EnterCriticalSection g_LViewCtlsCS
    Set oThread = GetSelectedThreadObject
    If (oThread Is Nothing) Then
        cbStopMon.Enabled = False
        cbStopProc.Enabled = False
        cbResumeMon.Enabled = False
        cbDelObject.Enabled = False
    Else
        cbStopMon.Enabled = (oThread.ProcessStatus = teProcessRunning)
        cbStopProc.Enabled = (oThread.ProcessStatus = teProcessRunning)
        cbResumeMon.Enabled = (oThread.ProcessStatus = teDisconnected)
        cbDelObject.Enabled = True
    End If
    LeaveCriticalSection g_LViewCtlsCS
End Sub
