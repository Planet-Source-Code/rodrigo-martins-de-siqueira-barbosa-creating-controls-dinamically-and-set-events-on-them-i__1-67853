VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form FrmScriptCalc 
   Caption         =   "Creating a calculator on run-time example"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGenerate 
      Caption         =   "Generate Controls"
      Height          =   555
      Left            =   4320
      TabIndex        =   2
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton CmdAddCode 
      Caption         =   "Add Code"
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   4095
   End
   Begin VB.TextBox TxtCodeScript 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FrmScriptCalc.frx":0000
      Top             =   600
      Width           =   6975
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   3960
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rodrigo_Martins@email.com"
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   6720
      Width           =   2025
   End
   Begin VB.Label Label3 
      Caption         =   "Created by Rodrigo Martins de Siqueira Barbosa - Feb 2007"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   6480
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmScriptCalc.frx":0831
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   9855
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmScriptCalc.frx":08BB
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "FrmScriptCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========[ Creating controls at Run-Time and SETTING EVENTS AT RUN-TIME! ]==========
' * Here i use the MS Script control for setting up events for controls created at     *
' * run-time. Creating the controls was the easy part, but i couldn't find a way to    *
' * set events on runtime to my mewly created controls... so I had the idea of use     *
' * subclassing to do it.                                                              *
' * After some research and time, here it is!                                          *
' * I hope someone will find it as useful as I found it enjoyable creating it.         *
' * Oh yeah, and if you like it, please give some credits to Marcel A. Fritsch,        *
' * his code on subclassing helped me a lot. Here's his code address:                  *
' * http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=32360&lngWId=1   *
' *                                                                                    *
' * Actually, only 2 events are working: Click and MouseMove.                          *
' * It's pretty simple to add other events, so I just created those two so that        *
' * one of them (click) is really useful and the other is just a proof of concept.     *
' *                                                                         HAVE FUN   *
' *                                              Rodrigo Martins de Siqueira Barbosa   *
' ======================================================================================

Private Sub CmdAddCode_Click()
ScriptControl1.AddCode TxtCodeScript.Text
End Sub

Private Sub CmdGenerate_Click()
'loads the calculator design from a text file
Dim LineValue As String
Dim ObjCreated As Object
Open App.Path & "\calc.template" For Input As #1
i = 0
Do While Not EOF(1)
    Line Input #1, LineValue
    i = i + 1
    If Mid(LineValue, 1, 1) <> "#" Then 'allows comments on text file
   
      Set ObjCreated = Controls.Add(GetValueFromLine(LineValue, "ControlType"), GetValueFromLine(LineValue, "ControlName"))
      ObjCreated.Width = GetValueFromLine(LineValue, "ControlWidth")
      ObjCreated.Height = GetValueFromLine(LineValue, "ControlHeight")
      ObjCreated.Top = GetValueFromLine(LineValue, "ControlTop")
      ObjCreated.Left = GetValueFromLine(LineValue, "ControlLeft")
      StrTmp = GetValueFromLine(LineValue, "ControlCaption") 'just for not repeat the search on the string
      If StrTmp <> "" Then
        ObjCreated.Caption = StrTmp
      End If
      ObjCreated.Visible = True
    End If
Loop
Close #1
LoadControlsOnScript
End Sub


Private Sub Form_Load()
lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf WindowProc, App.hInstance, App.ThreadID)
LoadControlsOnScript
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Give up global application hook
UnhookWindowsHookEx lHook
End Sub

Sub LoadControlsOnScript()
On Error Resume Next
For Each ctrl In Controls
    ScriptControl1.AddObject ctrl.Name, ctrl
Next
ScriptControl1.AddObject "Form1", Form1
End Sub

'gets value from a line like: name=xxxxxxx;type=yyyyyyyy
'so GetValueFromLine("name") will return "xxxxxxx"
Function GetValueFromLine(StrLine As String, Key As String) As String
startsel = 0
For i = 1 To Len(StrLine)
    If startsel = 0 Then
        If UCase(Mid(StrLine, i, Len(Key))) = UCase(Key) Then
            startsel = i + Len(Key) + 1
        End If
    End If
    If startsel <> 0 Then
        If Mid(StrLine, i, 1) = ";" Then
            endsel = i
            Exit For
        End If
    End If
Next i
GetValueFromLine = Mid(StrLine, startsel, endsel - startsel)
End Function

Sub EventRoute(chwnd As Long, MessageHandle As Long)
On Error Resume Next
Select Case MessageHandle
Case WM_Click
    For Each ctrl In Controls
        If ctrl.hwnd = chwnd Then
            ScriptControl1.ExecuteStatement ctrl.Name & "_Click()"
        End If
    Next
Case WM_MOUSEMOVE
    For Each ctrl In Controls
        If ctrl.hwnd = chwnd Then
            ScriptControl1.ExecuteStatement ctrl.Name & "_MouseMove()"
        End If
    Next
End Select
End Sub

