VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pro-Status"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9015
   StartUpPosition =   1  'CenterOwner
   Begin Project1.isButton isButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      icon            =   "Form1.frx":617A
      style           =   10
      caption         =   "About"
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "Form1.frx":6196
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin Project1.UniTextBox Box4 
      Height          =   855
      Left            =   1560
      TabIndex        =   16
      Top             =   5160
      Width           =   6015
      _extentx        =   10610
      _extenty        =   1508
      font            =   "Form1.frx":61BE
      forecolor       =   16744576
      text            =   "UniTextBox2"
      multiline       =   -1  'True
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox Box1 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1931
      font            =   "Form1.frx":61E6
      text            =   "buoc 1 nhap doan t?t"
      multiline       =   -1  'True
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":620E
      forecolor       =   255
      text            =   "P"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin VB.TextBox T2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3360
      Width           =   6975
   End
   Begin Project1.isButton isButton1 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      icon            =   "Form1.frx":623A
      style           =   8
      caption         =   "Copy"
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "Form1.frx":6256
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin Project1.isButton Covernt 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      icon            =   "Form1.frx":627E
      style           =   8
      caption         =   "Go"
      inonthemestyle  =   0
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "Form1.frx":629A
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin Project1.UniTextBox T1 
      Height          =   1455
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   6975
      _extentx        =   12303
      _extenty        =   2566
      font            =   "Form1.frx":62C2
      backcolor       =   -2147483643
      text            =   ""
      multiline       =   -1  'True
      scrollbar       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":62EA
      forecolor       =   33023
      text            =   "r"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":6316
      forecolor       =   49344
      text            =   "o"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   495
      _extentx        =   873
      _extenty        =   1508
      font            =   "Form1.frx":6342
      forecolor       =   65280
      text            =   "S"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   4
      Left            =   4320
      TabIndex        =   8
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":636E
      forecolor       =   12632064
      text            =   "t"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   5
      Left            =   4800
      TabIndex        =   9
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":639A
      forecolor       =   16711680
      text            =   "a"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   6
      Left            =   5280
      TabIndex        =   10
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":63C6
      forecolor       =   16711935
      text            =   "t"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   7
      Left            =   5760
      TabIndex        =   11
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":63F2
      forecolor       =   255
      text            =   "u"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox UniTextBox1 
      Height          =   855
      Index           =   8
      Left            =   6240
      TabIndex        =   12
      Top             =   360
      Width           =   375
      _extentx        =   661
      _extenty        =   1508
      font            =   "Form1.frx":641E
      forecolor       =   33023
      text            =   "s"
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
      alignment       =   2
   End
   Begin Project1.UniTextBox Box2 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1085
      font            =   "Form1.frx":644A
      text            =   "buoc 1 nhap doan t?t"
      multiline       =   -1  'True
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
   End
   Begin Project1.UniTextBox Box3 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1508
      font            =   "Form1.frx":6472
      text            =   "buoc 1 nhap doan t?t"
      multiline       =   -1  'True
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
   End
   Begin Project1.UniTextBox tBox 
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      font            =   "Form1.frx":649A
      text            =   "buoc 1 nhap doan t?t"
      multiline       =   -1  'True
      locked          =   -1  'True
      enabled         =   0   'False
      borderstyle     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const WM_SETTEXT = &HC
Dim hDlg As Long, hButton As Long
Dim r As String

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Const GW_HWNDNEXT = 2



Function S1(ByVal St1 As String) As String
Dim a
Dim I
On Error GoTo Loi
a = Split(St1, " ")
For I = 0 To UBound(a)
   a(I) = UCase(Left(a(I), 1)) & Right(a(I), Len(a(I)) - 1)
Next
S1 = Join(a, " ")
Loi:
End Function







Private Sub Covernt_Click()
T1.Text = S1(T1.Text)
T2 = ""
Dim I As Integer
For I = 1 To Len(T1.Text)
r = Mid$(T1.Text, I, 1)
'*************************

If r = "t" Then r = "†"
If r = "T" Then r = "†"
If r = "G" Then r = "(¯¬"
If r = "H" Then r = "]-["
If r = "M" Then r = "¶v¶"
If r = "V" Then r = "\/"
If r = "X" Then r = "><"
If r = "x" Then r = "×"

If r = "h" Then r = "|¬"
If r = "E" Then r = "€"
If r = "L" Then r = "£"
If r = "u" Then r = "µ"
If r = "C" Then r = "Ç"
If r = "B" Then r = "ß"
If r = "c" Then r = "¢"
If r = "i" Then r = "j"
If r = "Y" Then r = "¥"
If r = "D" Then r = "|)"
If r = "K" Then r = "|<"
If r = "k" Then r = "|«"
If r = "S" Then r = "§"
If r = "s" Then r = "š"
If r = "e" Then r = "ë"
If r = "o" Then r = "ø"
If r = "O" Then r = "0"
If r = "N" Then r = "Ñ"
If r = "Q" Then r = "Ø"
If r = "y" Then r = "ÿ"




If r = ChrW$(&H1ED3) Then r = "ô`"
If r = ChrW$(&H1ED1) Then r = "ô'"
If r = ChrW$(&H1ED5) Then r = "ô?"
If r = ChrW$(&H1ED7) Then r = "ô~"
If r = ChrW$(&H1ED9) Then r = "ô."

If r = ChrW$(&H1A1) Then r = "o*"
If r = ChrW$(&H1EDD) Then r = "o*"
If r = ChrW$(&H1EE3) Then r = "o*"
If r = ChrW$(&H1EE1) Then r = "o*"
If r = ChrW$(&H1EDB) Then r = "o*"
If r = ChrW$(&H1EDF) Then r = "o*"

'ChrW$(&H1EDF)
If r = ChrW$(&H1EBF) Then r = "ê'"
If r = ChrW$(&H1EC1) Then r = "ê`"
If r = ChrW$(&H1EC3) Then r = "ê?"
If r = ChrW$(&H1EC5) Then r = "ê~"
If r = ChrW$(&H1EC7) Then r = "ê."

If r = ChrW$(&H1ECF) Then r = "o?"
If r = ChrW$(&H1ECD) Then r = "o."


If r = ChrW$(&H1EA5) Then r = "â'"
If r = ChrW$(&H1EA7) Then r = "â`"
If r = ChrW$(&H1EA9) Then r = "â?"
If r = ChrW$(&H1EAB) Then r = "â~"
If r = ChrW$(&H1EAD) Then r = "â."

If r = ChrW$(&H103) Then r = "a"
If r = ChrW$(&H1EAF) Then r = "a'"
If r = ChrW$(&H1EB1) Then r = "a`"
If r = ChrW$(&H1EB3) Then r = "a?"
If r = ChrW$(&H1EB5) Then r = "a~"
If r = ChrW$(&H1EB7) Then r = "a."

If r = ChrW$(&H1EE7) Then r = "u?"
If r = ChrW$(&H1EE5) Then r = "u."

If r = ChrW$(&H1B0) Then r = "u*"
If r = ChrW$(&H1EE9) Then r = "u*"
If r = ChrW$(&H1EEB) Then r = "u*"
If r = ChrW$(&H1EED) Then r = "u*"
If r = ChrW$(&H1EEF) Then r = "u*"
If r = ChrW$(&H1EF1) Then r = "u*"

If r = ChrW$(&H1EA3) Then r = "a?"
If r = ChrW$(&H1EA1) Then r = "a."

If r = ChrW$(&HED) Then r = "j'"
If r = ChrW$(&HEC) Then r = "j`"
If r = ChrW$(&H1EC9) Then r = "j?"
If r = ChrW$(&H129) Then r = "j~"
If r = ChrW$(&H1ECB) Then r = "j."

If r = ChrW$(&H1EA5) Then r = "â'"
If r = ChrW$(&H1EA7) Then r = "â`"
If r = ChrW$(&H1EA9) Then r = "â?"
If r = ChrW$(&H1EAB) Then r = "â~"
If r = ChrW$(&H1EAD) Then r = "â."

If r = ChrW$(&H1EF3) Then r = "y`"
If r = ChrW$(&H1EF7) Then r = "y?"
If r = ChrW$(&H1EF9) Then r = "y~"
If r = ChrW$(&H1EF5) Then r = "y."

If r = ChrW$(&H110) Then r = "+)"



T2 = T2 & r
Next I
End Sub




Private Sub Form_Load()
Box1.Text = ChrW$(&H42) & ChrW$(&H1B0) & ChrW$(&H1EDB) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H20) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H6F) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EA7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H75) & ChrW$(&H79) & ChrW$(&H1EC3) & ChrW$(&H6E) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EA7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&HEC) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H2E)
Box2.Text = ChrW$(&H42) & ChrW$(&H1B0) & ChrW$(&H1EDB) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H32) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H47) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1ECB) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H2E) & ChrW$(&H2E) & ChrW$(&H2E)
Box3.Text = ChrW$(&H42) & ChrW$(&H1B0) & ChrW$(&H1EDB) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H33) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H6F) & ChrW$(&H70) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&HE0) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H61) & ChrW$(&H73) & ChrW$(&H74) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H53) & ChrW$(&H74) & ChrW$(&H61) & ChrW$(&H74) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H2E)
Box4.Text = ChrW$(&H43) & ChrW$(&H68) & ChrW$(&HFA) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EEF) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&HFA) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H75) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EBB) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EDB) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H2E)
tBox.Text = ChrW$(&H42) & ChrW$(&H1B0) & ChrW$(&H1EDB) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H34) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H3A)
T1.Text = ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1ED3) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC3) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H75) & ChrW$(&HF4) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EDB) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H5F) _
& ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&HE2) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&H1EE5) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EAD) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H2E)


End Sub


Private Sub isButton1_Click()
Clipboard.Clear
Clipboard.SetText T2.Text
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H6F) & ChrW$(&H70) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EC9) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EA7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H6C) & ChrW$(&H20) & ChrW$(&H2B) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H24) & ChrW$(&H5F) & ChrW$(&H24), vbOKOnly, "Xong", Me.hwnd
End Sub




Private Sub isButton2_Click()
UniMsgBox vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H72) & ChrW$(&H6F) & ChrW$(&H53) & ChrW$(&H74) & ChrW$(&H61) & ChrW$(&H74) & ChrW$(&H75) & ChrW$(&H73) & vbCrLf & ChrW$(&H50) & ChrW$(&H68) & ChrW$(&HE1) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H48) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H4D) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H68) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H21) & vbCrLf _
& ChrW$(&H54) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H47) & ChrW$(&H69) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H51) & ChrW$(&H75) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H39) & ChrW$(&H30) & ChrW$(&H40) & ChrW$(&H59) & ChrW$(&H61) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H6F) & ChrW$(&H2E) & ChrW$(&H43) & ChrW$(&H6F) & ChrW$(&H6D) & vbCrLf & _
ChrW$(&H43) & ChrW$(&H1EA3) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EEF) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EE5) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H21) & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbOKOnly, "Thank's You", Me.hwnd
End Sub

