Attribute VB_Name = "ModThread"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Function ThreadExeStarter(ByVal lpExe As Long) As Long

Dim oExe As ThreadExe

    CopyMemory oExe, lpExe, 4
    oExe.StartMonitor
    ZeroMemory oExe, 4
End Function
