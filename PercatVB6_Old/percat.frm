VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form pc 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PerCat 6.0"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   Enabled         =   0   'False
   Icon            =   "percat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   8520
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin percat.jcbutton sha 
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   5
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   32768
      Caption         =   "Scan Picture"
      ForeColor       =   16777215
   End
   Begin VB.Timer q2 
      Enabled         =   0   'False
      Interval        =   411
      Left            =   9480
      Top             =   6000
   End
   Begin VB.Timer q1 
      Enabled         =   0   'False
      Interval        =   411
      Left            =   9480
      Top             =   5640
   End
   Begin VB.Timer sss 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9600
      Top             =   7080
   End
   Begin VB.TextBox t01 
      Height          =   285
      Left            =   10680
      TabIndex        =   16
      Top             =   9720
      Width           =   150
   End
   Begin VB.TextBox t0 
      Height          =   195
      Left            =   10560
      TabIndex        =   15
      Top             =   9600
      Width           =   255
   End
   Begin VB.Timer timm 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   840
      Top             =   6360
   End
   Begin VB.TextBox ttt 
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   9840
      Width           =   375
   End
   Begin VB.Timer nnt 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   5520
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   9120
      Width           =   255
   End
   Begin VB.TextBox t1 
      Height          =   375
      Left            =   10560
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.Timer n1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1080
      Top             =   4800
   End
   Begin percat.jcbutton bt1 
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "&OK"
   End
   Begin VB.TextBox chat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Text            =   "Contents - Néi dung"
      Top             =   7800
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.TextBox ten 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "Your name - Tªn b¹n"
      Top             =   7200
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   3480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   840
   End
   Begin percat.jcbutton send 
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "&Send"
   End
   Begin percat.jcbutton svd 
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   5
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "Scan Video"
      ForeColor       =   16777215
   End
   Begin VB.Image anh 
      Height          =   5775
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Avatar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label sa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search from Bing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search from Yahoo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search from Google"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image bing 
      Height          =   1095
      Left            =   9480
      Picture         =   "percat.frx":57E2
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image yahoo 
      Height          =   1095
      Left            =   9480
      Picture         =   "percat.frx":6B88
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image google 
      Height          =   1095
      Left            =   9480
      Picture         =   "percat.frx":B37E
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label nn00 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   10800
      TabIndex        =   10
      Top             =   9720
      Width           =   15
   End
   Begin VB.Label so 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7680
      Width           =   15
   End
   Begin VB.Label vi 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   9240
      Width           =   9855
   End
   Begin VB.Label ct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   10095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   10560
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label en 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   8640
      Width           =   9855
   End
   Begin VB.Image percat 
      Height          =   7140
      Left            =   1680
      Picture         =   "percat.frx":F7A8
      Top             =   -240
      Width           =   7620
   End
End
Attribute VB_Name = "pc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
 
Public Function GoogleSpeak(ByVal sText As String, Optional ByVal Language As String = "en", Optional ByVal bDoevents As Boolean) As Boolean
    On Error Resume Next
    Dim sTempPath As String, ml As String
    Dim FileLength As Long
 
    sText = Replace(sText, vbCrLf, " ")
 
    If Len(sText) > 100 Then Exit Function
 
    sTempPath = Environ("Temp") & "\TempMP3.MP3"
 
    If URLDownloadToFile(0&, "http://translate.google.com/translate_tts?tl=" & Language & "&q=" & sText, sTempPath, 0&, 0&) = 0 Then
 
        If mciSendString("open " & Chr$(34) & sTempPath & Chr$(34) & " type MpegVideo" & " alias myfile", 0&, 0&, 0&) = 0 Then
 
            ml = String(30, 0)
            Call mciSendString("status myfile length ", ml, 30, 0&)
            FileLength = Val(ml)
            If FileLength Then
                If mciSendString("play myFile", 0&, 0&, 0&) = 0 Then
                    Do While mciSendString("status myfile position ", ml, 30, 0&) = 0
                        If Val(ml) = FileLength Then GoogleSpeak = True: Exit Do
                        If bDoevents Then DoEvents
                    Loop
                End If
            End If
            Call mciSendString("close myfile", 0&, 0&, 0&)
 
        End If
 
        Kill sTempPath
    End If
 
End Function
 



Private Sub bing_Click()
Dim rtt As String
td.Show
rtt = Replace(ttt.Text, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/search?q=" & rtt
Debug.Print GoogleSpeak("You are viewing results from bing", "en", True)
End Sub

Private Sub bing_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.Visible = True
Label1.Visible = False
Label2.Visible = False
End Sub

Private Sub bt1_Click()
Dim bd As String
percat.Picture = a.d2.Picture
en.Caption = "Hello, " & ten.Text
vi.Caption = "Xin chµo, " & ten.Text
bd = GetUrlSource("http://percat.esy.es/get.php?name=" & ten.Text)
Debug.Print GoogleSpeak("Hello ," & ten.Text, "en", True)
percat.Picture = a.d1.Picture
chat.Visible = True
send.Visible = True
ten.Visible = False
bt1.Visible = False
End Sub

Private Sub chat_Change()
send.Default = True
End Sub

Private Sub chat_Click()
chat.Text = ""
End Sub




Private Sub Form_Load()
'If Dir(App.Path & "\name.percat") = Empty Then
'Timer1.Enabled = True
'so.Caption = "1"
'Else
'Timer1.Enabled = False
'so.Caption = "1"
'n1.Enabled = True
'bing.Visible = False
'google.Visible = False
'yahoo.Visible = False
'End If
Dim na As String
Dim nb As String
Dim nc As String
na = GetUrlSource("http://percat.esy.es/ip.php")
t0.Text = na
nb = GetUrlSource("http://percat.esy.es/user/" & t0.Text & ".txt")
t01.Text = nb
If t01.Text = "no" Then
Timer1.Enabled = True
so.Caption = "1"
Else
Timer1.Enabled = False
so.Caption = "1"
n1.Enabled = True
bing.Visible = False
google.Visible = False
yahoo.Visible = False
End If
End Sub



Private Sub Image1_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub

Private Sub google_Click()
Dim rtt As String
td.Show
rtt = Replace(ttt.Text, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rtt
Debug.Print GoogleSpeak("You are viewing results from google", "en", True)
End Sub

Private Sub google_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.Visible = True
Label2.Visible = False
Label3.Visible = False
End Sub

Private Sub Label4_Click()
Dim bg As String
cd.ShowOpen
bg = cd.FileName
anh.Picture = LoadPicture(bg)
End Sub

Private Sub n1_Timer()
Dim yx As String
percat.Picture = a.d2.Picture
en.Caption = "Hello, " & t01.Text
vi.Caption = "Xin chµo, " & t01.Text
Debug.Print GoogleSpeak("Hello ," & t01.Text, "en", True)
percat.Picture = a.d1.Picture
chat.Visible = True
send.Visible = True
pc.Enabled = True
percat.Picture = a.d4.Picture
yx = GetUrlSource("http://percat.esy.es/wel.txt")
Text1.Text = yx
Debug.Print GoogleSpeak(Text1.Text, "en", True)
percat.Picture = a.d1.Picture
n1.Enabled = False
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

Private Sub percat_Click()
nnt.Enabled = True
en.Caption = "I am is Percat"
vi.Caption = "Xin chµo, T«i lµ Percat"
Debug.Print GoogleSpeak("I am is Percat", "en", True)
nnt.Enabled = False
End Sub

Private Sub rest_Timer()

End Sub

Private Sub q1_Timer()
nnt.Enabled = True
ttt.Text = chat.Text
timm.Enabled = True
' neu rong
If chat.Text = "" Then
ct.Caption = "You did not enter content!"
Debug.Print GoogleSpeak("You did not enter content!", "en", True)
en.Caption = ct.Caption
vi.Caption = "B¹n ch­a nhËp néi dung!"
' duyet am thanh
ElseIf InStr(chat.Text, "speak") > 0 Then
   chat.Text = Replace(chat.Text, "speak", "")
Dim do1i1 As String
do1i1 = chat.Text
Debug.Print GoogleSpeak(do1i1, "en", True)
' duyet web
ElseIf InStr(ct.Caption, "web") > 0 Then
   ct.Caption = Replace(ct.Caption, "web", "")
Dim do1i As String
td.Show
do1i = ct.Caption
td.brwWebBrowser.Navigate do1i
' duyet wiki
ElseIf InStr(chat.Text, "wiki") > 0 Then
   chat.Text = Replace(chat.Text, "wiki", "")
Dim doi11 As String
Dim rtt11 As String
td.Show
doi11 = chat.Text
rtt11 = Replace(doi11, " ", "+")
td.brwWebBrowser.Navigate "http://vi.wikipedia.org/w/index.php?search=" & rtt11
' kiem tra tra qn trg
ElseIf InStr(chat.Text, "scan") > 0 Then
   chat.Text = Replace(chat.Text, "scan", "")
Dim doi0 As String
Dim rtt0 As String
td.Show
doi0 = chat.Text
rtt0 = Replace(doi0, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rtt0
' kiem tra picre
ElseIf InStr(chat.Text, "image") > 0 Then
   chat.Text = Replace(chat.Text, "image", "")
Dim anh As String
Dim rttt As String
td.Show
anh = chat.Text
rttt = Replace(anh, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rttt & "&source=lnms&tbm=isch"
'anh lan 2
ElseIf InStr(chat.Text, "picture") > 0 Then
   chat.Text = Replace(chat.Text, "picture", "")
Dim an2h As String
Dim rttt0 As String
td.Show
an2h = chat.Text
rttt0 = Replace(an2h, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rttt0 & "&source=lnms&tbm=isch"
' kiem hoat hinh
ElseIf InStr(chat.Text, "animation") > 0 Then
   chat.Text = Replace(chat.Text, "animation", "hoat hinh")
Dim an2h23 As String
Dim rttt13 As String
td.Show
an2h23 = chat.Text
rttt13 = Replace(an2h23, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & rttt13
' kiem hoat hinh 2
ElseIf InStr(chat.Text, "ani") > 0 Then
   chat.Text = Replace(chat.Text, "ani", "hoat hinh")
Dim an2h233 As String
Dim rttt133 As String
td.Show
an2h233 = chat.Text
rttt133 = Replace(an2h233, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & rttt133
' hoat hinh vn
ElseIf InStr(chat.Text, "hoat hinh") > 0 Then
   chat.Text = Replace(chat.Text, "hoat hinh", "phim hoat hinh")
Dim an2h2334 As String
Dim rttt1334 As String
td.Show
an2h2334 = chat.Text
rttt1334 = Replace(an2h2334, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & rttt1334
' hoat hinh youtubbee
ElseIf InStr(chat.Text, "youtube") > 0 Then
   chat.Text = Replace(chat.Text, "youtube", "")
Dim awe As String
Dim aew As String
td.Show
awe = chat.Text
aew = Replace(awe, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & aew
' bing
ElseIf InStr(chat.Text, "bing") > 0 Then
   chat.Text = Replace(chat.Text, "bing", "")
Dim big As String
Dim gib As String
td.Show
big = chat.Text
gib = Replace(big, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/search?q=" & gib
' bing image
ElseIf InStr(chat.Text, "image bing") > 0 Then
   chat.Text = Replace(chat.Text, "image bing", "")
Dim bigim As String
Dim gibim As String
td.Show
bigim = chat.Text
gibim = Replace(bigim, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/images/search?q=" & gibim
' bing image2
ElseIf InStr(chat.Text, "bing image") > 0 Then
   chat.Text = Replace(chat.Text, "bing image", "")
Dim bigim2 As String
Dim gibim2 As String
td.Show
bigim2 = chat.Text
gibim2 = Replace(bigim2, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/images/search?q=" & gibim2
' google
ElseIf InStr(chat.Text, "google") > 0 Then
   chat.Text = Replace(chat.Text, "google", "")
Dim ggl As String
Dim glg As String
td.Show
ggl = chat.Text
glg = Replace(ggl, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & glg
' yahoo
ElseIf InStr(chat.Text, "yahoo") > 0 Then
   chat.Text = Replace(chat.Text, "yahoo", "")
Dim yah As String
Dim yahoo As String
td.Show
yah = chat.Text
yahoo = Replace(yah, " ", "+")
td.brwWebBrowser.Navigate "https://vn.search.yahoo.com/search?q=" & yahoo
' image yahoo
ElseIf InStr(chat.Text, "yahoo image") > 0 Then
   chat.Text = Replace(chat.Text, "yahoo image", "")
Dim yahim As String
Dim yahooim As String
td.Show
yahim = chat.Text
yahooim = Replace(yahim, " ", "+")
td.brwWebBrowser.Navigate "https://vn.images.search.yahoo.com/search/images?p=" & yahooim
' image yahoo 2
ElseIf InStr(chat.Text, "image yahoo") > 0 Then
   chat.Text = Replace(chat.Text, "image yahoo", "")
Dim yahim2 As String
Dim yahooim2 As String
td.Show
yahim2 = chat.Text
yahooim2 = Replace(yahim2, " ", "+")
td.brwWebBrowser.Navigate "https://vn.images.search.yahoo.com/search/images?p=" & yahooim2
' toan hoc math
ElseIf InStr(chat.Text, "math") > 0 Then
   chat.Text = Replace(chat.Text, "math", "")
Dim math As String
Dim mathh As String
td.Show
math = chat.Text
mathh = Replace(math, " ", "+")
td.brwWebBrowser.Navigate "http://coccoc.com/search/#query=" & mathh
' ping
ElseIf chat.Text = "ping" Then
Dim kw As String
kw = GetUrlSource("http://percat.esy.es/ip.php")
Debug.Print GoogleSpeak("Ip of you is " & kw, "en", True)
MsgBox "Ip of you is " & kw
' kiem tra oofline
ElseIf chat.Text = "mini" Then
hidemini.Show
Unload Me
ElseIf chat.Text = "what is the time now" Then
ct.Caption = "Now is " & Time
en.Caption = ct.Caption
vi.Caption = "V©y giê lµ " & Time
Debug.Print GoogleSpeak("Now is " & Time, "en", True)
ElseIf chat.Text = "what time is it" Then
ct.Caption = "Now is " & Time
en.Caption = ct.Caption
vi.Caption = "V©y giê lµ " & Time
Debug.Print GoogleSpeak("Now is " & Time, "en", True)
ElseIf chat.Text = "how old" Then
td.Show
td.brwWebBrowser.Navigate "http://how-old.net"
Debug.Print GoogleSpeak("open browser visit http://how-old.net", "en", True)
ElseIf chat.Text = "google" Then
td.Show
td.brwWebBrowser.Navigate "www.google.com.vn"
ElseIf chat.Text = "coccoc" Then
td.Show
td.brwWebBrowser.Navigate "www.coccoc.com/search"
ElseIf chat.Text = "end" Then
Debug.Print GoogleSpeak("Goodbye, see you again", "en", True)
End
ElseIf chat.Text = "exit" Then
Debug.Print GoogleSpeak("Goodbye, see you again", "en", True)
End
ElseIf chat.Text = "mp3" Then
 Dim r As Long
   r = ShellExecute(0, "open", "http://www.mp3.zing.vn", 0, 0, 1)
ElseIf chat.Text = "luutru360" Then
td.Show
td.brwWebBrowser.Navigate "www.luutru360.com"
ElseIf chat.Text = "image" Then
td.Show
td.brwWebBrowser.Navigate "www.bing.com/?scope=images"
ElseIf chat.Text = "search image" Then
td.Show
td.brwWebBrowser.Navigate "www.bing.com/?scope=images"
ElseIf chat.Text = "scan image" Then
td.Show
td.brwWebBrowser.Navigate "www.bing.com/?scope=images"
ElseIf chat.Text = "youtube" Then
ct.Caption = "Please install flash to see the video => https://get.adobe.com/flashplayer/"
Debug.Print GoogleSpeak("Please install flash to see the video", "en", True)
Dim qqqq As Long
   qqqq = ShellExecute(0, "open", "https://get.adobe.com/flashplayer/", 0, 0, 1)
ElseIf chat.Text = "rename" Then
MsgBox "Xin loi , ten chi duoc dat mot lan!", vbInformation, "Percat"
' nguoc lai thi getlink
Else
Dim noi As String
Dim kq As String
Dim boin As String
nnt.Enabled = True
en.Enabled = True
boin = StrConv(chat.Text, 2)
noi = Replace(boin, " ", "+")
kq = GetUrlSource("http://percat.esy.es/sv/1.php?chat=" & noi)
ct.Caption = kq
t1.Text = ct.Caption
en.Caption = "I will not get if you add these characters to"
nnt.Enabled = True
Debug.Print GoogleSpeak(t1.Text, "en", True)
End If
nnt.Enabled = False
vi.Caption = "Vui lßng kh«ng viÕt in hoa , kh«ng thªm c¸c kÝ tù hay dÊu chÊm ? ! @,..."
timm.Enabled = False
q1.Enabled = False
End Sub

Private Sub q2_Timer()
nnt.Enabled = True
ttt.Text = chat.Text
timm.Enabled = True
' neu rong
If chat.Text = "" Then
ct.Caption = "You did not enter content!"
Debug.Print GoogleSpeak("You did not enter content!", "en", True)
en.Caption = ct.Caption
vi.Caption = "B¹n ch­a nhËp néi dung!"
' duyet am thanh
ElseIf InStr(chat.Text, "speak") > 0 Then
   chat.Text = Replace(chat.Text, "speak", "")
Dim do1i11 As String
do1i11 = chat.Text
Debug.Print GoogleSpeak(do1i11, "en", True)
' duyet web
ElseIf InStr(ct.Caption, "web") > 0 Then
   ct.Caption = Replace(ct.Caption, "web", "")
Dim do11i As String
td.Show
do11i = ct.Caption
td.brwWebBrowser.Navigate do11i
' duyet wiki
ElseIf InStr(chat.Text, "wiki") > 0 Then
   chat.Text = Replace(chat.Text, "wiki", "")
Dim doi112 As String
Dim rtt112 As String
td.Show
doi112 = chat.Text
rtt112 = Replace(doi112, " ", "+")
td.brwWebBrowser.Navigate "http://vi.wikipedia.org/w/index.php?search=" & rtt11
' kiem tra tra qn trg
ElseIf InStr(chat.Text, "scan") > 0 Then
   chat.Text = Replace(chat.Text, "scan", "")
Dim doi01 As String
Dim rtt01 As String
td.Show
doi01 = chat.Text
rtt01 = Replace(doi01, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rtt01
' kiem tra picre
ElseIf InStr(chat.Text, "image") > 0 Then
   chat.Text = Replace(chat.Text, "image", "")
Dim anh1 As String
Dim rttt1 As String
td.Show
anh1 = chat.Text
rttt1 = Replace(anh1, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rttt1 & "&source=lnms&tbm=isch"
'anh lan 2
ElseIf InStr(chat.Text, "picture") > 0 Then
   chat.Text = Replace(chat.Text, "picture", "")
Dim an2h0 As String
Dim rttt00 As String
td.Show
an2h0 = chat.Text
rttt00 = Replace(an2h0, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rttt00 & "&source=lnms&tbm=isch"
' kiem hoat hinh
ElseIf InStr(chat.Text, "animation") > 0 Then
   chat.Text = Replace(chat.Text, "animation", "hoat hinh")
Dim an2h23s As String
Dim rttt13s As String
td.Show
an2h23s = chat.Text
rttt13s = Replace(an2h23s, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & rttt13s
' kiem hoat hinh 2
ElseIf InStr(chat.Text, "ani") > 0 Then
   chat.Text = Replace(chat.Text, "ani", "hoat hinh")
Dim an2h2303 As String
Dim rttt1303 As String
td.Show
an2h2303 = chat.Text
rttt1303 = Replace(an2h2303, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & rttt1303
' hoat hinh vn
ElseIf InStr(chat.Text, "hoat hinh") > 0 Then
   chat.Text = Replace(chat.Text, "hoat hinh", "phim hoat hinh")
Dim an2h23341 As String
Dim rttt13341 As String
td.Show
an2h23341 = chat.Text
rttt13341 = Replace(an2h23341, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & rttt13341
' hoat hinh youtubbee
ElseIf InStr(chat.Text, "youtube") > 0 Then
   chat.Text = Replace(chat.Text, "youtube", "")
Dim awe7 As String
Dim aew7 As String
td.Show
awe7 = chat.Text
aew7 = Replace(awe7, " ", "+")
td.brwWebBrowser.Navigate "https://www.youtube.com/results?search_query=" & aew7
' bing
ElseIf InStr(chat.Text, "bing") > 0 Then
   chat.Text = Replace(chat.Text, "bing", "")
Dim big8 As String
Dim gib8 As String
td.Show
big8 = chat.Text
gib8 = Replace(big8, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/search?q=" & gib8
' bing image
ElseIf InStr(chat.Text, "image bing") > 0 Then
   chat.Text = Replace(chat.Text, "image bing", "")
Dim bigimq As String
Dim gibimq As String
td.Show
bigimq = chat.Text
gibimq = Replace(bigimq, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/images/search?q=" & gibimq
' bing image2
ElseIf InStr(chat.Text, "bing image") > 0 Then
   chat.Text = Replace(chat.Text, "bing image", "")
Dim bigim20 As String
Dim gibim20 As String
td.Show
bigim20 = chat.Text
gibim20 = Replace(bigim20, " ", "+")
td.brwWebBrowser.Navigate "http://www.bing.com/images/search?q=" & gibim20
' google
ElseIf InStr(chat.Text, "google") > 0 Then
   chat.Text = Replace(chat.Text, "google", "")
Dim ggle As String
Dim glge As String
td.Show
ggle = chat.Text
glge = Replace(ggle, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & glge
' yahoo
ElseIf InStr(chat.Text, "yahoo") > 0 Then
   chat.Text = Replace(chat.Text, "yahoo", "")
Dim yahm As String
Dim yahoom As String
td.Show
yahm = chat.Text
yahoom = Replace(yahm, " ", "+")
td.brwWebBrowser.Navigate "https://vn.search.yahoo.com/search?q=" & yahoom
' image yahoo
ElseIf InStr(chat.Text, "yahoo image") > 0 Then
   chat.Text = Replace(chat.Text, "yahoo image", "")
Dim yahimr As String
Dim yahooimr As String
td.Show
yahimr = chat.Text
yahooimr = Replace(yahimr, " ", "+")
td.brwWebBrowser.Navigate "https://vn.images.search.yahoo.com/search/images?p=" & yahooimr
' image yahoo 2
ElseIf InStr(chat.Text, "image yahoo") > 0 Then
   chat.Text = Replace(chat.Text, "image yahoo", "")
Dim yahim2q As String
Dim yahooim2q As String
td.Show
yahim2q = chat.Text
yahooim2q = Replace(yahim2q, " ", "+")
td.brwWebBrowser.Navigate "https://vn.images.search.yahoo.com/search/images?p=" & yahooim2q
' toan hoc math
ElseIf InStr(chat.Text, "math") > 0 Then
   chat.Text = Replace(chat.Text, "math", "")
Dim mathq As String
Dim mathhq As String
td.Show
mathq = chat.Text
mathhq = Replace(mathq, " ", "+")
td.brwWebBrowser.Navigate "http://coccoc.com/search/#query=" & mathhq
' ping
ElseIf chat.Text = "ping" Then
Dim kwa As String
kwa = GetUrlSource("http://percat.esy.es/ip.php")
Debug.Print GoogleSpeak("Ip of you is " & kwa, "en", True)
MsgBox "Ip of you is " & kwa
' kiem tra oofline
ElseIf chat.Text = "mini" Then
hidemini.Show
Unload Me
ElseIf chat.Text = "what is the time now" Then
ct.Caption = "Now is " & Time
en.Caption = ct.Caption
vi.Caption = "V©y giê lµ " & Time
Debug.Print GoogleSpeak("Now is " & Time, "en", True)
ElseIf chat.Text = "what time is it" Then
ct.Caption = "Current time " & Time
en.Caption = ct.Caption
vi.Caption = "V©y giê lµ " & Time
Debug.Print GoogleSpeak("Current time " & Time, "en", True)
ElseIf chat.Text = "how old" Then
td.Show
td.brwWebBrowser.Navigate "http://how-old.net"
Debug.Print GoogleSpeak("open browser visit http://how-old.net", "en", True)
ElseIf chat.Text = "google" Then
td.Show
td.brwWebBrowser.Navigate "www.google.com.vn"
ElseIf chat.Text = "coccoc" Then
td.Show
td.brwWebBrowser.Navigate "www.coccoc.com/search"
ElseIf chat.Text = "end" Then
Debug.Print GoogleSpeak("Goodbye, see you again", "en", True)
End
ElseIf chat.Text = "exit" Then
Debug.Print GoogleSpeak("Bye bye!", "en", True)
End
ElseIf chat.Text = "mp3" Then
 Dim rr As Long
   rr = ShellExecute(0, "open", "http://www.mp3.zing.vn", 0, 0, 1)
ElseIf chat.Text = "luutru360" Then
td.Show
td.brwWebBrowser.Navigate "www.luutru360.com"
ElseIf chat.Text = "image" Then
td.Show
td.brwWebBrowser.Navigate "www.bing.com/?scope=images"
ElseIf chat.Text = "search image" Then
td.Show
td.brwWebBrowser.Navigate "www.bing.com/?scope=images"
ElseIf chat.Text = "scan image" Then
td.Show
td.brwWebBrowser.Navigate "www.bing.com/?scope=images"
ElseIf chat.Text = "youtube" Then
ct.Caption = "Please install flash to see the video => https://get.adobe.com/flashplayer/"
Debug.Print GoogleSpeak("Please install flash to see the video", "en", True)
Dim qqqqq As Long
   qqqqq = ShellExecute(0, "open", "https://get.adobe.com/flashplayer/", 0, 0, 1)
ElseIf chat.Text = "rename" Then
MsgBox "Xin loi , ten chi duoc dat mot lan!", vbInformation, "Percat"
' nguoc lai thi getlink
Else
Dim noi1 As String
Dim kq1 As String
Dim boin1 As String
nnt.Enabled = True
en.Enabled = True
boin1 = StrConv(chat.Text, 2)
noi1 = Replace(boin1, " ", "+")
kq1 = GetUrlSource("http://percat.esy.es/sv/2.php?chat=" & noi1)
ct.Caption = kq1
t1.Text = ct.Caption
en.Caption = "I will not get if you add these characters to"
nnt.Enabled = True
Debug.Print GoogleSpeak(t1.Text, "en", True)
End If
nnt.Enabled = False
vi.Caption = "Vui lßng kh«ng viÕt in hoa , kh«ng thªm c¸c kÝ tù hay dÊu chÊm ? ! @,..."
timm.Enabled = False
q2.Enabled = False
End Sub

Private Sub send_Click()
sss.Enabled = True
End Sub

Private Sub sha_Click()
Dim haa As String
ha.Show
haa = Replace(ttt.Text, " ", "+")
ha.wb.Navigate "https://www.google.com.vn/search?q=" & haa & "&&tbm=isch"
End Sub

Private Sub sss_Timer()
sa.Caption = Int(Rnd * 2)
If sa.Caption = "0" Then
q1.Enabled = True
sss.Enabled = False
ElseIf sa.Caption = "1" Then
q2.Enabled = True
sss.Enabled = False
End If
End Sub

Private Sub svd_Click()
Dim vdd As String
vd.Show
vdd = Replace(ttt.Text, " ", "+")
vd.wb.Navigate "https://www.google.com.vn/search?q=" & vdd & "&tbm=vid"
End Sub

Private Sub ten_Change()
bt1.Default = True
End Sub

Private Sub ten_Click()
ten.Text = ""
End Sub

Private Sub Timer1_Timer()
percat.Picture = a.d3.Picture
en.Caption = "Hello, I am is Percat"
vi.Caption = "Xin chµo, T«i lµ Percat"
Debug.Print GoogleSpeak("Hello, I am is Percat", "en", True)
percat.Picture = a.d5.Picture
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
percat.Picture = a.d4.Picture
en.Caption = "I am a virtual assistant"
vi.Caption = "T«i lµ mét trî lÝ ¶o"
Debug.Print GoogleSpeak("I am a virtual assistant", "en", True)
percat.Picture = a.d5.Picture
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
percat.Picture = a.d2.Picture
en.Caption = "I will tell you by what I know"
vi.Caption = "T«i sÏ cho b¹n biÕt nh÷ng g× t«i biÕt"
Debug.Print GoogleSpeak("I will tell you by what I know", "en", True)
percat.Picture = a.d1.Picture
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
percat.Picture = a.d5.Picture
en.Caption = "Now, let's begin!"
vi.Caption = "B©y giê, chóng ta h·y b¾t ®Çu!"
Debug.Print GoogleSpeak("Now, let's begin!", "en", True)
percat.Picture = a.d3.Picture
Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
percat.Picture = a.d2.Picture
en.Caption = "What is your name?"
vi.Caption = "B¹n tªn lµ g× ?"
Debug.Print GoogleSpeak("What is your name?", "en", True)
percat.Picture = a.d1.Picture
pc.Enabled = True
ten.Visible = True
bt1.Visible = True
Timer5.Enabled = False
End Sub



Private Sub timm_Timer()
If ct.Caption = "Sorry, I do not find, I will open google to assist you!" Then
Dim doi78 As String
Dim rtt78 As String
td.Show
ttt.Text = chat.Text
doi78 = chat.Text
rtt78 = Replace(doi78, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rtt78
bing.Visible = True
yahoo.Visible = True
google.Visible = True
sha.Visible = True
svd.Visible = True
timm.Enabled = False
ElseIf InStr(ct.Caption, "web") > 0 Then
   ct.Caption = Replace(ct.Caption, "web", "")
Dim do1i As String
td.Show
do1i = ct.Caption
td.brwWebBrowser.Navigate do1i
End If
End Sub

Private Sub yahoo_Click()
Dim rtt As String
td.Show
rtt = Replace(ttt.Text, " ", "+")
td.brwWebBrowser.Navigate "https://vn.search.yahoo.com/search?q=" & rtt
Debug.Print GoogleSpeak("You are viewing results from yahoo", "en", True)
End Sub

Private Sub yahoo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label2.Visible = True
Label1.Visible = False
Label3.Visible = False
End Sub
