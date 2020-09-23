Attribute VB_Name = "ModSync"
Option Explicit

Public Type CRITICAL_SECTION
    dummy1 As Long
    dummy2 As Long
    dummy3 As Long
    dummy4 As Long
    dummy5 As Long
    dummy6 As Long
End Type

Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
