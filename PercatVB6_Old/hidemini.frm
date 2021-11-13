VERSION 5.00
Begin VB.Form hidemini 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Percat Mini 4.0"
   ClientHeight    =   7875
   ClientLeft      =   5010
   ClientTop       =   2835
   ClientWidth     =   8340
   Icon            =   "hidemini.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin percat.jcbutton jcbutton1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      ButtonStyle     =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      Caption         =   "X"
   End
   Begin VB.TextBox t1 
      Height          =   285
      Left            =   8400
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.Timer timm 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   840
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
      Left            =   120
      TabIndex        =   1
      Text            =   "Contents - Néi dung"
      Top             =   6480
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   3000
   End
   Begin VB.Timer nnt 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   360
      Top             =   4320
   End
   Begin percat.jcbutton send 
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "&Send"
   End
   Begin VB.Image Image1 
      Height          =   7935
      Left            =   6480
      Picture         =   "hidemini.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
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
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   6255
   End
   Begin VB.Label so 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   135
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Image percat 
      Height          =   5100
      Left            =   480
      Picture         =   "hidemini.frx":910D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5340
   End
End
Attribute VB_Name = "hidemini"
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
 
Private Sub chat_Click()
chat.Text = ""
End Sub

Private Sub chat_DblClick()
chat.Text = ""
End Sub

Private Sub Form_Load()
Timer1.Interval = 10
    Me.Top = Screen.Height
    Me.Left = Screen.Width - Me.Width
    Timer1.Interval = 10
Unload pc
End Sub
Private Sub chat_Change()
send.Default = True
End Sub


Private Sub Image1_Click()
nnt.Enabled = True
Debug.Print GoogleSpeak("P e r c a t", "en", True)
Debug.Print GoogleSpeak("Per Cat", "en", True)
Debug.Print GoogleSpeak("Percat", "en", True)
nnt.Enabled = False
End Sub

Private Sub jcbutton1_Click()
End
End Sub

Private Sub send_Click()
nnt.Enabled = True
timm.Enabled = True
' neu rong
If chat.Text = "" Then
ct.Caption = "You did not enter content!"
Debug.Print GoogleSpeak("You did not enter content!", "en", True)
' kiem tra tra qn trg
ElseIf InStr(chat.Text, "scan") > 0 Then
   chat.Text = Replace(chat.Text, "scan", "")
Dim doi As String
Dim rtt As String
td.Show
doi = chat.Text
rtt = Replace(doi, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rtt
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
ElseIf chat.Text = "pc" Then
pc.Show
Unload Me
ElseIf chat.Text = "what is the time now" Then
ct.Caption = "Now is " & Time
Debug.Print GoogleSpeak("Now is " & Time, "en", True)
ElseIf chat.Text = "what time is it" Then
ct.Caption = "Now is " & Time
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
ElseIf chat.Text = "delete user" Then
Dim tenbn As String
Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(App.Path & "\name.percat", 1, , -2)
   tenbn = FSO.Readall
ct.Caption = "Delete User " & tenbn & " success!"
Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(App.Path & "\name.percat", True)
Set FSO = Nothing
Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(App.Path & "\name.percat", 2, , -1)
          FSO.Write ""
rest.Enabled = True
ElseIf chat.Text = "rename" Then
MsgBox "Xin loi , ten chi duoc dat mot lan!", vbInformation, "Percat"
' nguoc lai thi getlink
Else
Dim noi As String
Dim kq As String
Dim boin As String
nnt.Enabled = True
boin = StrConv(chat.Text, 2)
noi = Replace(boin, " ", "+")
kq = GetUrlSource("http://percat.esy.es/?chat=" & noi)
ct.Caption = kq
t1.Text = ct.Caption
nnt.Enabled = True
Debug.Print GoogleSpeak(t1.Text, "en", True)
End If
nnt.Enabled = False
timm.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Me.Top = Me.Top - 50
    If Me.Top < Screen.Height - Me.Height - 300 Then Timer1.Enabled = False
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

Private Sub timm_Timer()
If ct.Caption = "Sorry, I do not find, I will open google to assist you!" Then
Dim doi As String
Dim rtt As String
td.Show
doi = chat.Text
rtt = Replace(doi, " ", "+")
td.brwWebBrowser.Navigate "https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & rtt
timm.Enabled = False
End If
End Sub
