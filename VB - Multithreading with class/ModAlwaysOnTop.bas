Attribute VB_Name = "ModAlwaysOnTop"
Option Explicit

'System menu stuff
Private Const MF_BYCOMMAND As Long = &H0&
Private Const SC_CLOSE As Long = &HF060
Private Const SC_ALWAYSONTOP As Long = &H1

Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_STRING As Long = &H0&
Private Const MFT_SEPARATOR As Long = MF_SEPARATOR
Private Const MFT_STRING As Long = MF_STRING
Private Const MF_CHECKED As Long = &H8&
Private Const MF_UNCHECKED As Long = &H0&
Private Const MFS_CHECKED As Long = MF_CHECKED
Private Const MFS_UNCHECKED As Long = MF_UNCHECKED

Private Const MIIM_ID As Long = &H2
Private Const MIIM_TYPE As Long = &H10
Private Const MIIM_STATE As Long = &H1


Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long

'Subclassing stuff
Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_SYSCOMMAND As Long = &H112

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Always-on-top stuff
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Module varaibles
Private g_hWnd As Long
Private g_lOldProc As Long

Public Sub StartSubclass(ByVal hWnd As Long)
    g_hWnd = hWnd
    g_lOldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewProc)
End Sub

Public Sub StopSubclass()
    SetWindowLong g_hWnd, GWL_WNDPROC, g_lOldProc
End Sub

Private Function NewProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Static bOnTop As Boolean
Dim lZOrder As Long
Dim mif As MENUITEMINFO
Dim hMenu As Long
Dim lCount As Long

    Select Case uMsg
        Case WM_SYSCOMMAND
            If (wParam = SC_ALWAYSONTOP) Then
                bOnTop = Not (bOnTop)
                lZOrder = IIf(bOnTop, HWND_TOPMOST, HWND_NOTOPMOST)
                SetWindowPos hWnd, lZOrder, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                hMenu = GetSystemMenu(hWnd, 0)
                lCount = GetMenuItemCount(hMenu)
                With mif
                    .cbSize = Len(mif)
                    .fMask = MIIM_STATE
                    .fState = IIf(bOnTop, MFS_CHECKED, MFS_UNCHECKED)
                End With
                SetMenuItemInfo hMenu, lCount - 1, 1, mif
            Else
                NewProc = CallWindowProc(g_lOldProc, hWnd, uMsg, wParam, lParam)
            End If
        Case Else
            NewProc = CallWindowProc(g_lOldProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

Public Sub AddAlwaysOnTopToSysMenu(ByVal hWnd As Long)

Dim hMenu As Long
Dim lCount As Long
Dim mif As MENUITEMINFO

    'Get the system menu handle
    hMenu = GetSystemMenu(hWnd, 0)
    'Find the menu count
    lCount = GetMenuItemCount(hMenu)
    'Add a menu item separator
    With mif
        .cbSize = Len(mif)
        .fMask = MIIM_TYPE
        .fType = MFT_SEPARATOR
    End With
    InsertMenuItem hMenu, lCount, 1, mif
    'Add the wanted menu item
    With mif
        .fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .dwTypeData = "Always on top"
        .cch = Len(.dwTypeData)
        .wID = SC_ALWAYSONTOP
    End With
    InsertMenuItem hMenu, lCount + 1, 1, mif
End Sub
