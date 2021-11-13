VERSION 5.00
Begin VB.Form a 
   BorderStyle     =   0  'None
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image n15 
      Height          =   255
      Left            =   3720
      Picture         =   "a.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n14 
      Height          =   255
      Left            =   3120
      Picture         =   "a.frx":202F
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n13 
      Height          =   255
      Left            =   2520
      Picture         =   "a.frx":405D
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n12 
      Height          =   255
      Left            =   1920
      Picture         =   "a.frx":608F
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n11 
      Height          =   255
      Left            =   1320
      Picture         =   "a.frx":80C0
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n10 
      Height          =   255
      Left            =   720
      Picture         =   "a.frx":A0F3
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n9 
      Height          =   255
      Left            =   120
      Picture         =   "a.frx":C126
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image n8 
      Height          =   255
      Left            =   4200
      Picture         =   "a.frx":E158
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n7 
      Height          =   255
      Left            =   3720
      Picture         =   "a.frx":10189
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n6 
      Height          =   255
      Left            =   3120
      Picture         =   "a.frx":121BB
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n5 
      Height          =   255
      Left            =   2520
      Picture         =   "a.frx":141EE
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n4 
      Height          =   255
      Left            =   1920
      Picture         =   "a.frx":16220
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n3 
      Height          =   255
      Left            =   1320
      Picture         =   "a.frx":18252
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n2 
      Height          =   255
      Left            =   720
      Picture         =   "a.frx":1A283
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image n1 
      Height          =   255
      Left            =   120
      Picture         =   "a.frx":1C2B5
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image d13 
      Height          =   495
      Left            =   3480
      Picture         =   "a.frx":1E2E6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image d11 
      Height          =   495
      Left            =   2760
      Picture         =   "a.frx":1F3E6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image d12 
      Height          =   495
      Left            =   3240
      Picture         =   "a.frx":20241
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image d6 
      Height          =   495
      Left            =   0
      Picture         =   "a.frx":21341
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image d7 
      Height          =   495
      Left            =   720
      Picture         =   "a.frx":22441
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image d8 
      Height          =   495
      Left            =   1320
      Picture         =   "a.frx":23541
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image d9 
      Height          =   495
      Left            =   1680
      Picture         =   "a.frx":24641
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image d10 
      Height          =   495
      Left            =   2280
      Picture         =   "a.frx":25741
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image d5 
      Height          =   495
      Left            =   2520
      Picture         =   "a.frx":26841
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image d4 
      Height          =   495
      Left            =   1920
      Picture         =   "a.frx":27941
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Image d3 
      Height          =   495
      Left            =   1440
      Picture         =   "a.frx":28A41
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image d2 
      Height          =   495
      Left            =   840
      Picture         =   "a.frx":29B41
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Image d1 
      Height          =   495
      Left            =   120
      Picture         =   "a.frx":2AC41
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
