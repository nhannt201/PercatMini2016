VERSION 5.00
Begin VB.Form ann 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   630
   ClientLeft      =   13905
   ClientTop       =   9630
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin percat.jcbutton jcbutton1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Caption         =   "&Show"
   End
End
Attribute VB_Name = "ann"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
hidemini.Show
Unload Me
End Sub
