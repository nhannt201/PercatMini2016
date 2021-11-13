VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Percat 6.0 - Picture Scan"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15510
   Icon            =   "ha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin percat.jcbutton st 
      Height          =   855
      Left            =   12840
      TabIndex        =   5
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
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
      Caption         =   "Stop Auto"
      ForeColor       =   16777215
   End
   Begin percat.jcbutton at 
      Height          =   825
      Left            =   10080
      TabIndex        =   4
      Top             =   960
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   1455
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
      BackColor       =   11169024
      Caption         =   "Auto"
      ForeColor       =   16777215
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   11040
      Top             =   600
   End
   Begin percat.jcbutton nx 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
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
      Caption         =   "View Each Photo"
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   7815
      Left            =   0
      TabIndex        =   2
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
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   15495
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   15495
   End
End
Attribute VB_Name = "ha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub at_Click()
t1.Enabled = True
End Sub

Private Sub nx_Click()
AutoPlayNextLoop
End Sub

Private Sub Form_Load()
wb.Navigate "https://www.google.com.vn/search?q=doraemon&source=lnms&tbm=isch&sa=X&ei=m_BcVZXBGuHLmwWExoHgCQ&ved=0CAgQ_AUoAg&biw=1360&bih=679"
End Sub

Sub AutoPlayNextLoop()
    Static iIndex%
    iIndex = iIndex + 1
    If iIndex = List1.ListCount Then iIndex = 0
   wb.Navigate List1.List(iIndex)
End Sub



Private Sub st_Click()
t1.Enabled = False
End Sub

Private Sub t1_Timer()
AutoPlayNextLoop
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
If InStr(1, x(i), "http://www.google.com.vn/imgres?imgurl=") > 0 Or InStr(1, x(i)) > 0 Then List1.AddItem x(i)
Next
End Sub

