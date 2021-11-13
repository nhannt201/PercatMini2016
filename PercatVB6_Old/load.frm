VERSION 5.00
Begin VB.Form load 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "PerCat 2.0"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
   Icon            =   "load.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox t12 
      Height          =   285
      Left            =   11760
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1920
      Width           =   150
   End
   Begin VB.TextBox t11 
      Height          =   285
      Left            =   11760
      TabIndex        =   9
      Top             =   1080
      Width           =   150
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   10080
      Top             =   6720
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1320
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   99
      Left            =   600
      Top             =   5520
   End
   Begin VB.Timer nnt 
      Interval        =   800
      Left            =   600
      Top             =   600
   End
   Begin VB.Label ss 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   3600
      Width           =   0
   End
   Begin VB.Label q7 
      BackStyle       =   0  'Transparent
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   7800
      TabIndex        =   7
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label q6 
      BackStyle       =   0  'Transparent
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   6
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label q5 
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   5
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label q4 
      BackStyle       =   0  'Transparent
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   4
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label q3 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   3
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label q2 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label q1 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3600
      TabIndex        =   1
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label so 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   15
   End
   Begin VB.Image percat 
      Height          =   7140
      Left            =   2400
      Picture         =   "load.frx":57E2
      Top             =   0
      Width           =   7620
   End
End
Attribute VB_Name = "load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Const REG_SZ = 1 ' Unicode nul terminated string
Private Const REG_BINARY = 3 ' Free form binary
Private Const HKEY_CURRENT_USER = &H80000001
Dim PerCats  As String

Private Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey Ret
End Sub

Private Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegDeleteValue Ret, strValue
    RegCloseKey Ret
End Sub
Private Sub Form_Load()
 If Len(App.Path) <> 3 Then
        PerCats = App.Path + "\" + App.EXEName + (".exe")
    Else
        PerCats = App.Path + App.EXEName + (".exe")
    End If

    'Thêm kh?i d?ng cùng Windows
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Percat", PerCats
t12.Text = ""
t11.Text = ""
End Sub

Private Sub nnt_Timer()
so.Caption = Int(Rnd * 13)
If so.Caption = "0" Then
percat.Picture = a.d1.Picture
ElseIf so.Caption = "1" Then
percat.Picture = a.d2.Picture
ElseIf so.Caption = "2" Then
percat.Picture = a.d3.Picture
ElseIf so.Caption = "3" Then
percat.Picture = a.d4.Picture
ElseIf so.Caption = "4" Then
percat.Picture = a.d5.Picture
ElseIf so.Caption = "5" Then
percat.Picture = a.d6.Picture
ElseIf so.Caption = "6" Then
percat.Picture = a.d7.Picture
ElseIf so.Caption = "7" Then
percat.Picture = a.d8.Picture
ElseIf so.Caption = "8" Then
percat.Picture = a.d9.Picture
ElseIf so.Caption = "9" Then
percat.Picture = a.d10.Picture
ElseIf so.Caption = "10" Then
percat.Picture = a.d11.Picture
ElseIf so.Caption = "11" Then
percat.Picture = a.d12.Picture
ElseIf so.Caption = "12" Then
percat.Picture = a.d13.Picture
ElseIf so.Caption = "13" Then
percat.Picture = a.d5.Picture
End If
End Sub


Private Sub Timer1_Timer()
If ss.Caption = "0" Then
q7.ForeColor = &HC0FFC0
q1.ForeColor = &HFF&
ss.Caption = "1"
ElseIf ss.Caption = "1" Then
q1.ForeColor = &HC0FFC0
q2.ForeColor = &HFF&
ss.Caption = "2"
ElseIf ss.Caption = "2" Then
q2.ForeColor = &HC0FFC0
q3.ForeColor = &HFF&
ss.Caption = "3"
ElseIf ss.Caption = "3" Then
q3.ForeColor = &HC0FFC0
q4.ForeColor = &HFF&
ss.Caption = "4"
ElseIf ss.Caption = "4" Then
q4.ForeColor = &HC0FFC0
q5.ForeColor = &HFF&
ss.Caption = "5"
ElseIf ss.Caption = "5" Then
q5.ForeColor = &HC0FFC0
q6.ForeColor = &HFF&
ss.Caption = "6"
ElseIf ss.Caption = "6" Then
q6.ForeColor = &HC0FFC0
q7.ForeColor = &HFF&
ss.Caption = "0"
End If
End Sub

Private Sub Timer2_Timer()
If ss.Caption = "0" Then
q7.ForeColor = &HFF00&
q1.ForeColor = &HFFFF&
ss.Caption = "1"
ElseIf ss.Caption = "1" Then
q1.ForeColor = &HFF00&
q2.ForeColor = &HFFFF&
ss.Caption = "2"
ElseIf ss.Caption = "2" Then
q2.ForeColor = &HFF00&
q3.ForeColor = &HFFFF&
ss.Caption = "3"
ElseIf ss.Caption = "3" Then
q3.ForeColor = &HFF00&
q4.ForeColor = &HFFFF&
ss.Caption = "4"
ElseIf ss.Caption = "4" Then
q4.ForeColor = &HFF00&
q5.ForeColor = &HFFFF&
ss.Caption = "5"
ElseIf ss.Caption = "5" Then
q5.ForeColor = &HFF00&
q6.ForeColor = &HFFFF&
ss.Caption = "6"
ElseIf ss.Caption = "6" Then
q6.ForeColor = &HFF00&
q7.ForeColor = &HFFFF&
ss.Caption = "0"
End If
End Sub

Private Sub Timer3_Timer()
pc.Show
Unload Me
Timer3.Enabled = False
End Sub
