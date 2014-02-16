VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   2040
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Const GW_HWNDNEXT = 2: Const WM_GETTEXTLENGTH = &HE: Const WM_GETTEXT = &HD
Private Type POINTAPI
    X As Long:     Y As Long
End Type

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const WM_SETTEXT = &HC
Dim hDlg As Long, hButton As Long
Dim r As String





Private Sub Command1_Click()
Dim a&
    a = WinGetHandle&(TextBox1)
    TextBox2 = "Handle: " & a & vbCr & WinGetText$(a)
End Sub

Public Function WinGetHandle&(szTitle$)
    Dim test_hwnd&, Child_Hwnd&
    test_hwnd& = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd& <> 0
        If InStr(LCase(WinGetText$(test_hwnd&)), LCase(szTitle$)) Then
            WinGetHandle& = test_hwnd&
            Exit Function
        End If
        test_hwnd& = GetWindow(test_hwnd&, GW_HWNDNEXT)
    Loop
End Function
 
Public Function WinGetText$(hwnd&)
    Dim Length&
    Length& = SendMessage(hwnd&, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1
    WinGetText$ = Space(Length&)
    SendMessage hwnd&, WM_GETTEXT, ByVal Length&, ByVal StrPtr(WinGetText$)
    WinGetText$ = Left$(WinGetText$, Length&)
End Function

Private Sub Timer1_Timer()

  If GetAsyncKeyState(1) = 0 Then
        'Label1.Caption = "Your Left Mouse Button Is UP"
    Else
       Dim Pt As POINTAPI, mWnd As Long
    GetCursorPos Pt 'Get the current cursor position
    mWnd = WindowFromPoint(Pt.X, Pt.Y) 'Get the window under the cursor
    TextBox3 = "Handle: " & mWnd & vbCr & WinGetText$(mWnd)
    End If



End Sub

