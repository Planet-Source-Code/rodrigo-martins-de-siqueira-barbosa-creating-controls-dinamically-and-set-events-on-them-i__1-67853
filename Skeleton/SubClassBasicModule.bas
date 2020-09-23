Attribute VB_Name = "SubClassBasicModule"
Option Explicit
' USER32 - Functions
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
(ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" _
(ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
' KERNEL32 - Functions
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
' CONSTANTS
Public Const WH_CALLWNDPROC = 4
Public Const WM_MOUSEMOVE = 132
Public Const WM_Click = 533
' STRUCTS
Type MYSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type
' REST
Public lHook As Long
Public Function WindowProc(ByVal Hookid As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim SWINP As MYSTRUCT
    CopyMemory SWINP, ByVal lParam, Len(SWINP)
    WindowProc = CallNextHookEx(lHook, Hookid, wParam, ByVal lParam)
    FrmSkeleton.EventRoute SWINP.hwnd, SWINP.message
End Function





