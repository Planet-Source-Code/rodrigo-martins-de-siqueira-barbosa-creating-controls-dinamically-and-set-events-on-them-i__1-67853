VERSION 5.00
Begin VB.Form FrmSkeleton 
   Caption         =   "Skeleton Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmSkeleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========[            Setting events using windows messages              ]==========
' * This is the skeleton program for using subclasses to set events on runtime         *
' * See the example calculator for a idea of how to use it. Here i just wanted to      *
' * keep it simple and clean.                                                          *
' *                                                                         HAVE FUN   *
' *                                              Rodrigo Martins de Siqueira Barbosa   *
' ======================================================================================
Private Sub Form_Load()
lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf WindowProc, App.hInstance, App.ThreadID)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Give up global application hook
UnhookWindowsHookEx lHook
End Sub

Sub EventRoute(chwnd As Long, MessageHandle As Long)
On Error Resume Next
Select Case MessageHandle
Case WM_Click
    For Each ctrl In Controls
        If ctrl.hwnd = chwnd Then
            'How the control will react to the event Click -- do it here
        End If
    Next
Case WM_MOUSEMOVE
    For Each ctrl In Controls
        If ctrl.hwnd = chwnd Then
            'How the control will react to the event MouseMove -- do it here
        End If
    Next
End Select
End Sub

