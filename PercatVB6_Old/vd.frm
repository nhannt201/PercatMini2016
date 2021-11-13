VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form vd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Percat 6.0 - Video Scan"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15345
   Icon            =   "vd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   15345
   StartUpPosition =   2  'CenterScreen
   Begin percat.jcbutton nx 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   1508
      ButtonStyle     =   6
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
      BackColor       =   14800597
      Caption         =   "View Each Video"
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   15495
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   15495
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   15495
      ExtentX         =   27331
      ExtentY         =   13785
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "vd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub nx_Click()
AutoPlayNextLoop
End Sub

Private Sub Form_Load()
wb.Navigate "https://www.google.com.vn/search?q=doraemon&biw=1360&bih=679&tbm=vid&source=lnms&sa=X&ei=nfBcVauYE4K7mQWXqIHYAg&ved=0CAcQ_AUoAQ&dpr=1"
End Sub

Sub AutoPlayNextLoop()
    Static iIndex%
    iIndex = iIndex + 1
    If iIndex = List1.ListCount Then iIndex = 0
   wb.Navigate List1.List(iIndex)
End Sub



Private Sub wb_DownloadComplete()
  'you must add the "Microsoft HTML Object Library"!!!!!!!!!
   Dim HTMLdoc As HTMLDocument
        Dim HTMLlinks As HTMLAnchorElement
            Dim STRtxt As String
    ' List the links.
   On Error Resume Next
        Set HTMLdoc = wb.Document
            For Each HTMLlinks In HTMLdoc.links
                STRtxt = STRtxt & HTMLlinks.href & vbCrLf
            Next HTMLlinks
        Text1.Text = STRtxt
        Dim i%: Dim x$(): x = Split(Text1.Text, vbNewLine)
For i = 0 To UBound(x)
If InStr(1, x(i), "http://www.youtube.com/watch?v=") > 0 Or InStr(1, x(i)) > 0 Then List1.AddItem x(i)
Next
End Sub


