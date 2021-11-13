VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTitleUni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "USER32.DLL" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DefWindowProcW Lib "USER32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETTEXT As Long = &HC


Public Property Let TitleBarCaption(ByVal hWnd As Long, ByVal NewValue As String)
   DefWindowProcW hWnd, WM_SETTEXT, 0, ByVal StrPtr(Uni(NewValue))
End Property
