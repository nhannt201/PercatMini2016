VERSION 5.00
Begin VB.UserControl jcbutton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   DefaultCancel   =   -1  'True
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "jcButton.ctx":0000
End
Attribute VB_Name = "jcbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'***************************************************************************
'*  Title:      JC button
'*  Function:   An ownerdrawn multistyle button
'*  Author:     Juned Chhipa
'*  Created:    November 2008
'*  Contact me: juned.chhipa@yahoo.com
'*
'*  Copyright © 2008-2009 Juned Chhipa. All rights reserved.
'***************************************************************************
'* This control can be used as an alternative to Command Button. It is
'* a lightweight button control which will emulate new command buttons.
'* Compile to get more faster results
'*
'* This control uses self-subclassing routines of Paul Caton.
'* Feel free to use this control. Please read Licence.txt
'* Please send comments/suggestions/bug reports to juned.chhipa@yahoo.com
'****************************************************************************
'*
'* - CREDITS:
'* - Paul Caton  :-  Self-Subclass Routines
'* - Noel Dacara :-  For helping me (Also, his dcbutton helped me a lot)
'* - Jim Jose    :-  To make grayscale (disabled) bitmap/icon
'* - Carles P.V. :-  For fastest gradient routines
'*   If any bugs found, please report  :- juned.chhipa@yahoo.com
'*
'* I have tested this control many times and I have tried my best to make
'* it work as a real command button. But still, I cannot guarantee that
'* that this is FREE OF BUGS. So please let me know if u find any.

'****************************************************************************
'* This software is provided "as-is" without any express/implied warranty.  *
'* In no event shall the author be held liable for any damages arising      *
'* from the use of this software.                                           *
'* If you do not agree with these terms, do not install "JCButton". Use     *
'* of the program implicitly means you have agreed to these terms.          *        *
'                                                                           *
'* Permission is granted to anyone to use this software for any purpose,    *
'* including commercial use, and to alter and redistribute it, provided     *
'* that the following conditions are met:                                   *
'*                                                                          *
'* 1.All redistributions of source code files must retain all copyright     *
'*   notices that are currently in place, and this list of conditions       *
'*   without any modification.                                              *
'*                                                                          *
'* 2.All redistributions in binary form must retain all occurrences of      *
'*   above copyright notice and web site addresses that are currently in    *
'*   place (for example, in the About boxes).                               *
'*                                                                          *
'* 3.Modified versions in source or binary form must be plainly marked as   *
'*   such, and must not be misrepresented as being the original software.   *
'****************************************************************************

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINT, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

'User32 Declares
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TransparentBlt Lib "MSIMG32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'==========================================================================================================================================================================================================================================================================================
' Subclassing Declares
Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum
Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

'Windows Messages
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOVING As Long = &H216
Private Const WM_NCACTIVATE As Long = &H86
Private Const WM_ACTIVATE As Long = &H6
Private Const ALL_MESSAGES As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04 As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05 As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08 As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09 As Long = 137                                      'Table A (after) entry count patch offset
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hwndTrack As Long
    dwHoverTime As Long
End Type

'for subclass
Private Type tSubData                                                            'Subclass data type
    hwnd As Long                                            'Handle of the window being subclassed
    nAddrSub As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig As Long                                            'The address of the pre-existing WndProc
    nMsgCntA As Long                                            'Msg after table entry count
    nMsgCntB As Long                                            'Msg before table entry count
    aMsgTblA() As Long                                            'Msg after table array
    aMsgTblB() As Long                                            'Msg Before table array
End Type

'for subclass
Private sc_aSubData() As tSubData                                        'Subclass data array
Private bTrack As Boolean
Private bTrackUser32 As Boolean

'Kernel32 declares used by the Subclasser
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'  End of Subclassing Declares
'==========================================================================================================================================================================================================================================================================================================

'[Enumerations]
Public Enum enumButtonStlyes
    [eStandard]                 '1) Standard VB Button
    [eFlat]                     '2) Standard Toolbar Button
    [eWindowsXP]                '3) Famous Win XP Button
    [eXPToolbar]                '4) XP Toolbar
    [eVistaAero]                '5) The New Vista Aero Button
    [eAOL]                      '6) AOL Buttons
    [eInstallShield]            '7) InstallShield?!?~?
    [eOutlook2007]              '8) Office 2007 Outlook Button
    [eVistaToolbar]             '9) Vista Toolbar Button
    [eVisualStudio]            '10) Visual Studio 2005 Button
    [eGelButton]               '11) Gel Button
    [e3DHover]                 '13) 3D Hover Button
    [eFlatHover]               '12) Flat Hover Button
    [eOffice2003]              '13) Office 2003 Style
    [eAqua]                    '14) Aqua (Near to MACOSX)
    [eSleek]                   '15) Somewhere I saw this (Not Bad!)
End Enum
#If False Then
    Private eStandard, eFlat, eVistaAero, eVistaToolbar, eInstallShield, eFlatHover, eOffice2003, eSleek
    Private eWindowsXP, eAqua, eXPToolbar, eVisualStudio, e3DHover, eGelButton, eOutlook2007, eAOL
#End If
Public Enum enumButtonStates
    [eStateNormal]              'Normal State
    [eStateOver]                'Hover State
    [eStateDown]                'Down State
End Enum
#If False Then
    'A trick to preserve casing when typing in IDE
    Private eStateNormal, eStateOver, eStateDown, eStateDisabled, eStateFocus
#End If
Public Enum enumCaptionAlign
    [ecLeftAlign]
    [ecCenterAlign]
    [ecRightAlign]
End Enum
#If False Then
    'A trick to preserve casing when typing in IDE
    Private ecLeftAlign, ecCenterAlign, ecRightAlign
#End If
Public Enum enumPictureAlign
    [epLeftEdge]
    [epLeftOfCaption]
    [epRightEdge]
    [epRightOfCaption]
    [epCenter]
    [epTopEdge]
    [epTopOfCaption]
    [epBottomEdge]
    [epBottomOfCaption]
End Enum
#If False Then
    Private epLeftEdge, epRightEdge, epRightOfCaption, epLeftOfCaption, epCenter
    Private epTopEdge, epTopOfCaption, epBottomEdge, epBottomOfCaption
#End If
Public Enum enumPictureSize
    [epsNormal]
    [eps16x16]
    [eps24x24]
    [eps32x32]
End Enum
#If False Then
    Private epsNormal, eps16x16, eps24x24, eps32x32
#End If
Public Enum GradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum
#If False Then
    Private gdHorizontal, gdVertical, gdDownwardDiagonal, gdUpwardDiagonal
#End If
'  used for Button colors
Private Type tButtonColors
    tBackColor As Long
    tDisabledColor As Long
    tForeColor As Long
    tGreyText As Long
End Type
'  used to define various graphics areas
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINT
    x As Long
    Y As Long
End Type
'  RGB Colors structure
Private Type RGBColor
    r As Single
    g As Single
    B As Single
End Type
'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128    '* Maintenance string for PSS usage.
End Type
' --constants for unicode support
Private Const VER_PLATFORM_WIN32_NT = 2
' --constants for  Flat Button
Private Const BDR_RAISEDINNER As Long = &H4
' --constants for Win 98 style buttons
Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
' --System Hand Pointer
Private Const IDC_HAND As Long = 32649
' --Color Constant
Private Const CLR_INVALID As Long = &HFFFF
Private Const DIB_RGB_COLORS As Long = 0
' --Formatting Text Consts
Private Const DT_SINGLELINE As Long = &H20
' --for drawing Icon Constants
Private Const DI_NORMAL As Long = &H3
' --Property Variables:
Private m_Picture As StdPicture           'Icon of button
Private m_PicOver As StdPicture
Private m_PicSize As enumPictureSize
Private m_PictureAlign As enumPictureAlign     'Picture Alignments
Private PicSizeW As Long                 'Picture's Height
Private PicSizeH As Long                 'Picture's Width
Private m_ButtonStyle As enumButtonStlyes     'Choose your Style
Private m_Buttonstate As enumButtonStates     'Normal / Over / Down
Private m_bIsDown As Boolean              'Is button is pressed?
Private m_bMouseInCtl As Boolean              'Is Mouse in Control
Private m_bHasFocus As Boolean              'Has focus?
Private m_bHandPointer As Boolean              'Use Hand Pointer
Private m_lCursor As Long
Private m_bDefault As Boolean              'Is Default?
Private m_bCheckBoxMode As Boolean              'Is checkbox?
Private m_bValue As Boolean              'Value (Checked/Unchekhed)
Private m_bShowFocus As Boolean              'Bool to show focus
Private m_bParentActive As Boolean              'Parent form Active or not
Private m_lParenthWnd As Long                 'Is parent active?
Private m_WindowsNT As Long                 'OS Supports Unicode?
Private m_bEnabled As Boolean              'Enabled/Disabled
Private m_Caption As String               'String to draw caption
Private m_TextRect As RECT                 'Text Position
Private m_CapRect As RECT                 'For InstallShield style
Private m_CaptionAlign As enumCaptionAlign
Private m_bColors As tButtonColors        'Button Colors
Private m_bUseMaskColor As Boolean              'Transparent areas
Private m_lMaskColor As Long                 'Set Transparent color
Private m_lButtonRgn As Long                 'Button Region
Private m_bIsSpaceBarDown As Boolean              'Space bar down boolean
Private m_ButtonRect As RECT                 'Button Position
Private m_FocusRect As RECT
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private m_lDownButton As Integer              'For click/Dblclick events
Private m_lDShift As Integer              'A flag for dblClick
Private m_lDX As Single
Private m_lDY As Single
Private lh As Long                 'ScaleHeight of button
Private lw As Long                 'ScaleWidth of button
Private XPos As Long                 'X position of picture
Private YPos As Long                 'Y Position of Picture
'  Events
Public Event Click()
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAcsii As Integer)

'  PRIVATE ROUTINES
Private Function PaintGrayScale(ByVal lhdc As Long, ByVal hPicture As Long, ByVal lLeft As Long, ByVal lTop As Long, Optional ByVal lWidth As Long = -1, Optional ByVal lHeight As Long = -1) As Boolean
'****************************************************************************
'*  Converts an icon/bitmap to grayscale (used for Disabled buttons)        *
'*  Author:  Jim Jose                                                       *
'*  Modified by me for Disabled Bitmaps (for Maskcolor)
'*  All Credits goes to Jim Jose                                            *
'****************************************************************************
    Dim BMP As BITMAP
    Dim BMPiH As BITMAPINFOHEADER
    Dim lBits() As Byte    'Packed DIB
    Dim lTrans() As Byte    'Packed DIB
    Dim TmpDC As Long
    Dim x As Long
    Dim xMax As Long
    Dim TmpCol As Long
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long
    Dim bIsIcon As Boolean
    Dim hdcSrc As Long
    Dim hOldob As Long
    Dim PicSize As Long
    Dim oPic As New StdPicture
    Set oPic = m_Picture
    '  Get the Image format
    If (GetObjectType(hPicture) = 0) Then
        Dim mIcon As ICONINFO
        bIsIcon = True
        GetIconInfo hPicture, mIcon
        hPicture = mIcon.hbmColor
    End If
    '  Get image info
    GetObject hPicture, Len(BMP), BMP
    '  Prepare DIB header and redim. lBits() array
    With BMPiH
        .biSize = Len(BMPiH)    '40
        .biPlanes = 1
        .biBitCount = 24
        .biWidth = BMP.bmWidth
        .biHeight = BMP.bmHeight
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        If lWidth = -1 Then lWidth = .biWidth
        If lHeight = -1 Then lHeight = .biHeight
    End With
    ReDim lBits(Len(BMPiH) + BMPiH.biSizeImage)   '[Header + Bits]
    '  Create TemDC and Get the image bits
    TmpDC = CreateCompatibleDC(lhdc)
    GetDIBits TmpDC, hPicture, 0, BMP.bmHeight, lBits(0), BMPiH, 0
    '  Loop through the array... (grayscale - average!!)
    xMax = BMPiH.biSizeImage - 1
    For x = 0 To xMax - 3 Step 3
        R1 = lBits(x)
        G1 = lBits(x + 1)
        B1 = lBits(x + 2)
        TmpCol = (R1 + G1 + B1) \ 3
        lBits(x) = TmpCol
        lBits(x + 1) = TmpCol
        lBits(x + 2) = TmpCol
    Next x
    '  Paint it!
    If bIsIcon Then
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, 0  ' Get the mask
        StretchDIBits lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd    ' Draw the mask
        PaintGrayScale = StretchDIBits(lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)  'Draw the gray
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    Else
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, 0  ' Get the mask
        StretchDIBits lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd    ' Draw the mask
        PaintGrayScale = StretchDIBits(lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    End If
    '   Clear memory
    DeleteDC TmpDC
End Function

Private Sub DrawLineApi(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'****************************************************************************
'*  draw lines
'****************************************************************************
    Dim pt As POINT
    Dim hPen As Long
    Dim hPenOld As Long
    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(UserControl.hdc, hPen)
    MoveToEx UserControl.hdc, X1, Y1, pt
    LineTo UserControl.hdc, X2, Y2
    SelectObject UserControl.hdc, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
End Sub

Private Function BlendColors(ByVal lBackColorFrom As Long, ByVal lBackColorTo As Long) As Long
'***************************************************************************
'*  Combines (mix) two colors                                              *
'***************************************************************************
    BlendColors = RGB(((lBackColorFrom And &HFF) + (lBackColorTo And &HFF)) / 2, (((lBackColorFrom \ &H100) And &HFF) + ((lBackColorTo \ &H100) And &HFF)) / 2, (((lBackColorFrom \ &H10000) And &HFF) + ((lBackColorTo \ &H10000) And &HFF)) / 2)
End Function

Private Sub DrawRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)
'****************************************************************************
'*  Draws a rectangle specified by coords and color of the rectangle        *
'****************************************************************************
    Dim brect As RECT
    Dim hBrush As Long
    Dim ret As Long
    brect.Left = x
    brect.Top = Y
    brect.Right = x + Width
    brect.Bottom = Y + Height
    hBrush = CreateSolidBrush(Color)
    ret = FrameRect(hdc, brect, hBrush)
    ret = DeleteObject(hBrush)
End Sub

Private Sub DrawFocusRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
'****************************************************************************
'*  Draws a Focus Rectangle inside button if m_bShowFocus property is True  *
'****************************************************************************
    Dim brect As RECT
    Dim retval As Long
    brect.Left = x
    brect.Top = Y
    brect.Right = x + Width
    brect.Bottom = Y + Height
    retval = DrawFocusRect(hdc, brect)
End Sub

Private Sub DrawGradientEx(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As GradientDirectionCts)
'****************************************************************************
'* Draws very fast Gradient in four direction.                              *
'* Author: Carles P.V (Gradient Master)                                     *
'* This routine works as a heart for this control.                          *
'* Thank you so much Carles.                                                *
'****************************************************************************
    Dim uBIH As BITMAPINFOHEADER
    Dim lBits() As Long
    Dim lGrad() As Long
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long
    Dim R2 As Long
    Dim G2 As Long
    Dim B2 As Long
    Dim dR As Long
    Dim dG As Long
    Dim dB As Long
    Dim Scan As Long
    Dim i As Long
    Dim iEnd As Long
    Dim iOffset As Long
    Dim j As Long
    Dim jEnd As Long
    Dim iGrad As Long
    '-- A minor check
    'If (Width < 1 Or Height < 1) Then Exit Sub
    If (Width < 1 Or Height < 1) Then
        Exit Sub
    End If
    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    R1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    G1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    B1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    R2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    G2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    B2 = Color2 Mod &H100&
    '-- Get color distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    '-- Size gradient-colors array
    Select Case GradientDirection
    Case [gdHorizontal]
        ReDim lGrad(0 To Width - 1)
    Case [gdVertical]
        ReDim lGrad(0 To Height - 1)
    Case Else
        ReDim lGrad(0 To Width + Height - 2)
    End Select
    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
    Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    '-- Render gradient DIB
    Select Case GradientDirection
    Case [gdHorizontal]
        For j = 0 To jEnd
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(i - iOffset)
            Next i
            iOffset = iOffset + Scan
        Next j
    Case [gdVertical]
        For j = jEnd To 0 Step -1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(j)
            Next i
            iOffset = iOffset + Scan
        Next j
    Case [gdDownwardDiagonal]
        iOffset = jEnd * Scan
        For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset - Scan
            iGrad = j
        Next j
    Case [gdUpwardDiagonal]
        iOffset = 0
        For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset + Scan
            iGrad = j
        Next j
    End Select
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    '-- Paint it!
    StretchDIBits UserControl.hdc, x, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy
End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette As Long = 0) As Long
'****************************************************************************
'*  System color code to long rgb                                           *
'****************************************************************************
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub RedrawButton()
'****************************************************************************
'*  The main routine of this usercontrol. Everything is drawn here.         *
'****************************************************************************
    UserControl.Cls                                'Clears usercontrol
    lh = ScaleHeight
    lw = ScaleWidth
    SetRect m_ButtonRect, 0, 0, lw, lh             'Sets the button rectangle
    If (m_bCheckBoxMode) Then                      'If Checkboxmode True
        If Not (m_ButtonStyle = eStandard Or m_ButtonStyle = eXPToolbar Or m_ButtonStyle = eVisualStudio) Then
            If m_bValue Then m_Buttonstate = eStateDown
        End If
    End If
    Select Case m_ButtonStyle
    Case eStandard
        DrawStandardButton m_Buttonstate
    Case e3DHover
        DrawStandardButton m_Buttonstate
    Case eFlat
        DrawStandardButton m_Buttonstate
    Case eFlatHover
        DrawStandardButton m_Buttonstate
    Case eWindowsXP
        DrawWinXPButton m_Buttonstate
    Case eXPToolbar
        DrawXPToolbar m_Buttonstate
    Case eGelButton
        DrawGelButton m_Buttonstate
    Case eAOL
        DrawAOLButton m_Buttonstate
    Case eInstallShield
        DrawInstallShieldButton m_Buttonstate
    Case eVistaAero
        DrawVistaButton m_Buttonstate
    Case eVistaToolbar
        DrawVistaToolbarStyle m_Buttonstate
    Case eVisualStudio
        DrawVisualStudio2005 m_Buttonstate
    Case eOutlook2007
        DrawOutlook2007 m_Buttonstate
    Case eOffice2003
        DrawOffice2003 m_Buttonstate
    Case eAqua
        DrawAquaButton m_Buttonstate
    Case eSleek
        DrawSleekButton m_Buttonstate
    End Select
    DrawPicwithCaption
End Sub

Private Sub CreateRegion()
'***************************************************************************
'*  Create region everytime you redraw a button.                           *
'*  Because some settings may have changed the button regions              *
'***************************************************************************
    If m_lButtonRgn Then DeleteObject m_lButtonRgn
    Select Case m_ButtonStyle
    Case eWindowsXP, eVistaAero, eVistaToolbar, eInstallShield
        m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 3, 3)
    Case eGelButton, eXPToolbar
        m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 4, 4)
    Case eAqua
        m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 2, lh + 2, 18, 18)
    Case Else
        m_lButtonRgn = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
    End Select
    SetWindowRgn UserControl.hwnd, m_lButtonRgn, True       'Set Button Region
    DeleteObject m_lButtonRgn                               'Free memory
End Sub

Private Sub DrawPicwithCaption()
'****************************************************************************
'* Draws a Picture in Enabled / Disabled mode along with Caption            *
'* Also captions are drawn here calculating all rects                       *
'* Routine to make GrayScale images is the work of Jim Jose.                *
'****************************************************************************
'=========================================================================
'* This routine resulted longer and difficult to read due to 9 picture
'* alignments along with 3 caption alignments.
'* And with Picture Normal and Picture Over features along with
'* UseMaskColor option
'* But it's still fast enough!
'=========================================================================
    Dim PicX As Long                   'X position of picture
    Dim PicY As Long                   'Y Position of Picture
    Dim tmpPic As New StdPicture      'Temp picture (Normal)
    Dim tmpPicOver As New StdPicture      'Over picture
    Dim hdcSrc As Long
    Dim hOldob As Long
    Dim lpRect As RECT                  'RECT to draw caption
    Dim CaptionW As Long                  'Width of Caption
    Dim CaptionH As Long                  'Height of Caption
    Dim CaptionX As Long                  'Left of Caption
    Dim CaptionY As Long                  'Top of Caption
    lw = ScaleWidth                          'Height of Button
    lh = ScaleHeight                         'Width of Button
    '  Get the Caption's height and Width
    CaptionW = TextWidth(m_Caption)          'Caption's Width
    CaptionH = TextHeight(m_Caption)         'Caption's Height
    '  Copy the original picture into a temp var
    Set tmpPic = m_Picture
    Set tmpPicOver = m_PicOver
    ' --Adjust Picture Sizes
    Select Case m_PicSize
    Case epsNormal
        PicSizeH = ScaleX(tmpPic.Height, vbHimetric, vbPixels)
        PicSizeW = ScaleX(tmpPic.Width, vbHimetric, vbPixels)
    Case eps16x16
        PicSizeH = 16
        PicSizeW = 16
    Case eps24x24
        PicSizeH = 24
        PicSizeW = 24
    Case eps32x32
        PicSizeH = 32
        PicSizeW = 32
    End Select
    ' --This button not allows different types for different states.
    ' --I could do this, but it makes great overload and waste f time
    ' --See the coding at the end of routine :(
    If tmpPic.Type <> tmpPicOver.Type Then
        Set tmpPicOver = m_Picture
    End If
    '=========================================================================
    ' --Set text pos and Pic pos
    Select Case m_PictureAlign
    Case epLeftOfCaption
        PicX = (lw - (PicSizeW + CaptionW)) \ 2
        If PicX < 4 Then PicX = 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2 - CaptionW \ 2) + (PicSizeW \ 2) + 3    'Some distance of 3
        If CaptionX < (PicSizeW + 8) Then CaptionX = PicSizeW + 8    'Text shouldn't draw over picture
        CaptionY = (lh \ 2 - CaptionH \ 2)
    Case epLeftEdge
        PicX = 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2) - (CaptionW \ 2) + (PicSizeW \ 2)
        If CaptionX < (PicSizeW + 8) Then CaptionX = PicSizeW + 8    'Text shouldn't draw over picture
        CaptionY = (lh \ 2 - CaptionH \ 2)
    Case epRightEdge
        PicX = lw - PicSizeW - 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw - CaptionW - 4) - PicSizeW
        CaptionY = (lh \ 2 - CaptionH \ 2)
    Case epRightOfCaption
        PicX = (lw - (PicSizeW - CaptionW)) \ 2
        If PicX > (lw - PicSizeW - 4) Then PicX = lw - PicSizeW - 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2 - CaptionW \ 2) - (PicSizeW \ 2) - 3
        If CaptionX + CaptionW < CaptionW Then
            CaptionX = (lw - CaptionW - 4) - PicSizeW
        End If
        CaptionY = lh \ 2 - (CaptionH \ 2)
    Case epCenter
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2) - (CaptionW \ 2)
        CaptionY = (lh \ 2) - CaptionH \ 2
    Case epBottomEdge
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - PicSizeH) - 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 - PicSizeH \ 2 - CaptionH \ 2) - 2
    Case epBottomOfCaption
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - (PicSizeH - CaptionH)) \ 2
        If PicY > lh - PicSizeH - 4 Then PicY = lh - PicSizeH - 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 - PicSizeH \ 2 - CaptionH \ 2) - 2
    Case epTopEdge
        PicX = (lw - PicSizeW) \ 2
        PicY = 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 + PicSizeH \ 2 - CaptionH \ 2) + 2
    Case epTopOfCaption
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - (PicSizeH + CaptionH)) \ 2
        If PicY < 4 Then PicY = 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 + PicSizeH \ 2 - CaptionH \ 2) + 2
    End Select
    ' --Minor check if picture's size exceeds button size
    If PicX < 1 Then PicX = 1
    If PicY < 1 Then PicY = 1
    If PicX + PicSizeW > ScaleWidth Then PicSizeW = ScaleWidth - 8
    If PicY + PicSizeH > ScaleHeight Then PicSizeH = ScaleHeight - 4
    '========================================================================
    ' --Calculate caption rects with Caption Alignment
    If m_Picture Is Nothing Then
        ' --Calculate caption rects if no picture available
        Select Case m_CaptionAlign
        Case ecLeftAlign
            CaptionX = 4
        Case ecCenterAlign
            CaptionX = (lw \ 2) - (CaptionW \ 2)
        Case ecRightAlign
            CaptionX = (lw - CaptionW - 4)
        End Select
        CaptionY = (lh \ 2) - (CaptionH \ 2)
        PicX = 0
        PicY = 0
    Else
        ' --There is a picture, so calc rects with that too.. (depending on Picture Align)
        Select Case m_CaptionAlign
        Case ecLeftAlign
            If m_PictureAlign = epLeftEdge Then
                CaptionX = PicSizeW + 8
            ElseIf m_PictureAlign = epLeftOfCaption Then
                PicX = 4
                CaptionX = PicX + PicSizeW + 4
            ElseIf m_PictureAlign = epRightEdge Then
                If CaptionX < 4 Then
                    CaptionX = (lw - CaptionW - 4) - PicSizeW
                Else
                    CaptionX = 4
                End If
            ElseIf m_PictureAlign = epRightOfCaption Then
                CaptionX = 4
                PicX = CaptionW + 4
            Else
                CaptionX = 4
            End If
        Case ecRightAlign
            If m_PictureAlign = epRightEdge Then
                CaptionX = (lw - CaptionW - 4) - PicSizeW
            ElseIf m_PictureAlign = epRightOfCaption Then
                PicX = lw - PicSizeW - 4
                CaptionX = (lw - CaptionW - 4) - PicSizeW
            ElseIf m_PictureAlign = epLeftEdge Then
                CaptionX = (lw - CaptionW - 4)
                If CaptionX < PicSizeW + 4 Then
                    CaptionX = PicSizeW + 4
                End If
            ElseIf m_PictureAlign = epLeftOfCaption Then
                CaptionX = (lw - CaptionW - 4)
                PicX = CaptionX - PicSizeW - 4
            Else
                CaptionX = (lw - CaptionW - 4)
            End If
        Case ecCenterAlign
            If m_PictureAlign = epRightEdge Then
                If CaptionX + CaptionW < CaptionW Then
                    CaptionX = (lw - CaptionW - 4) - PicSizeW
                Else
                    CaptionX = (lw \ 2) - (CaptionW \ 2)
                End If
            End If
        End Select
    End If
    ' --Uncomment the below lines and see what happens!! Oops
    ' --The caption draws awkwardly with accesskeys!
    If UserControl.AccessKeys <> vbNullString Then
        CaptionX = CaptionX + 3
    End If
    '=========================================================================
    ' --Adjust Picture Positions
    Select Case m_ButtonStyle
    Case eStandard, eFlat, eVistaToolbar, eXPToolbar
        If m_Buttonstate = eStateDown Then
            PicX = PicX + 1
            PicY = PicY + 1
        End If
    Case eAOL
        If m_Buttonstate = eStateDown Then
            PicX = PicX + 2     'More depth for AOL
            PicY = PicY + 2
        Else
            PicX = PicX - 1     'For AOL
            PicY = PicY - 1
        End If
    End Select
    '=========================================================================
    ' --If picture available, Set text rects with Picture
    If m_Buttonstate = eStateDown Then
        Select Case m_ButtonStyle
        Case eStandard, eFlat, eVistaToolbar, eXPToolbar
            ' --Caption pos for Standard/Flat buttons on down state
            SetRect lpRect, CaptionX + 1, CaptionY + 1, (CaptionW + CaptionX) + 1, (CaptionH + CaptionY) + 1
        Case eAOL
            ' --Caption RECT for AOL buttons
            SetRect lpRect, CaptionX + 1, CaptionY + 2, (CaptionW + CaptionX) + 1, (CaptionH + CaptionY) + 1
        Case Else
            ' --for other buttons on down state
            SetRect lpRect, CaptionX, CaptionY, CaptionW + CaptionX, CaptionH + CaptionY
        End Select
    Else
        Select Case m_ButtonStyle
        Case eAOL
            SetRect lpRect, CaptionX - 2, CaptionY - 2, CaptionW + CaptionX - 2, CaptionH + CaptionY - 2
        Case Else
            SetRect lpRect, CaptionX, CaptionY, CaptionW + CaptionX, CaptionH + CaptionY
            ' --For drawing Focus rect exactly around Caption
            SetRect m_CapRect, CaptionX - 2, CaptionY, CaptionW + CaptionX + 1, CaptionH + CaptionY + 1
        End Select
    End If
    '=======================================================================
    ' --Draw Picture Enabled/Disabled depending of Pic type and button state
    Select Case tmpPic.Type
    Case vbPicTypeIcon
        If m_bEnabled Then
            Select Case m_Buttonstate
            Case eStateNormal
                DrawIconEx hdc, PicX, PicY, tmpPic.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
            Case eStateOver
                If m_PicOver Is Nothing Then
                    DrawIconEx hdc, PicX, PicY, tmpPic.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hdc, PicX, PicY, tmpPicOver.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                End If
            Case eStateDown
                If m_PicOver Is Nothing Then
                    DrawIconEx hdc, PicX, PicY, tmpPic.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hdc, PicX, PicY, tmpPicOver.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                End If
            End Select
            ' --For checkboxmode
            If m_bCheckBoxMode And m_bValue Then
                If m_PicOver Is Nothing Then
                    DrawIconEx hdc, PicX, PicY, tmpPic.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hdc, PicX, PicY, tmpPicOver.handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                End If
            End If

        Else
            ' --Draw grayed picture (Thanks to Jim Jose)
            PaintGrayScale hdc, tmpPic.handle, PicX, PicY, PicSizeW, PicSizeH
        End If
    Case vbPicTypeBitmap
        If m_bEnabled Then
            ' --Draw picture with Maskcolor
            hdcSrc = CreateCompatibleDC(0)
            Select Case m_Buttonstate
            Case eStateNormal
                hOldob = SelectObject(hdcSrc, tmpPic.handle)
            Case eStateOver
                If m_PicOver Is Nothing Then
                    hOldob = SelectObject(hdcSrc, tmpPic.handle)
                Else
                    hOldob = SelectObject(hdcSrc, tmpPicOver.handle)
                End If
            Case eStateDown
                If m_PicOver Is Nothing Then
                    hOldob = SelectObject(hdcSrc, tmpPic.handle)
                Else
                    hOldob = SelectObject(hdcSrc, tmpPicOver.handle)
                End If
            End Select
            If m_bCheckBoxMode And m_bValue Then
                If m_PicOver Is Nothing Then
                    hOldob = SelectObject(hdcSrc, tmpPic.handle)
                Else
                    hOldob = SelectObject(hdcSrc, tmpPicOver.handle)
                End If
            End If
            If m_bUseMaskColor Then
                TransparentBlt UserControl.hdc, PicX, PicY, PicSizeW, PicSizeH, hdcSrc, 0, 0, PicSizeW, PicSizeH, m_lMaskColor
                SelectObject hdcSrc, hOldob
                DeleteDC hdcSrc
            Else
                TransparentBlt UserControl.hdc, PicX, PicY, PicSizeW, PicSizeH, hdcSrc, 0, 0, PicSizeW, PicSizeH, -1
                SelectObject hdcSrc, hOldob
                DeleteDC hdcSrc
            End If
        Else
            ' --Disabled Bitmap (Thanks to Jim Jose.)
            PaintGrayScale hdc, tmpPic.handle, PicX, PicY, PicSizeW, PicSizeH
        End If
    End Select
    '=========================================================================
    ' --At last, draw text
    SetTextColor hdc, IIf(m_bEnabled, TranslateColor(m_bColors.tForeColor), TranslateColor(vbGrayText))
    If Not m_WindowsNT Then
        ' --Unicode not supported
        DrawText hdc, m_Caption, Len(m_Caption), lpRect, DT_SINGLELINE  'Button looks good in SingleLine!
    Else
        ' --Supports Unicode (i.e above Windows NT)
        DrawTextW hdc, StrPtr(m_Caption), Len(m_Caption), lpRect, DT_SINGLELINE
    End If
    ' --Clear memory
    Set tmpPic = Nothing
    Set tmpPicOver = Nothing
End Sub

Private Sub SetAccessKey()
    Dim i As Long
    UserControl.AccessKeys = ""
    If Len(m_Caption) > 1 Then
        i = InStr(1, m_Caption, "&", vbTextCompare)
        If (i < Len(m_Caption)) And (i > 0) Then
            If Mid$(m_Caption, i + 1, 1) <> "&" Then
                AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
            Else
                i = InStr(i + 2, m_Caption, "&", vbTextCompare)
                If Mid$(m_Caption, i + 1, 1) <> "&" Then
                    AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
                End If
            End If
        End If
    End If
End Sub

Private Sub DrawCorners(Color As Long)
'****************************************************************************
'* Draws four Corners of the button specified by Color                      *
'****************************************************************************
    With UserControl
        lh = .ScaleHeight
        lw = .ScaleWidth
        SetPixel .hdc, 1, 1, Color
        SetPixel .hdc, 1, lh - 2, Color
        SetPixel .hdc, lw - 2, 1, Color
        SetPixel .hdc, lw - 2, lh - 2, Color
    End With
End Sub

Private Sub DrawStandardButton(ByVal vState As enumButtonStates)
'****************************************************************************
' Draws  four different styles in one procedure                             *
' Makes reading the code difficult, but saves much space!! ;)               *
'****************************************************************************
    Dim FocusRect As RECT
    Dim tmpRect As RECT
    lh = ScaleHeight
    lw = ScaleWidth
    SetRect m_ButtonRect, 0, 0, lw, lh
    If Not m_bEnabled Then
        '     Draws raised edge border
        DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
    End If
    If m_bCheckBoxMode And m_bValue Then
        PaintRect ShiftColor(TranslateColor(m_bColors.tBackColor), 0.02), m_ButtonRect
        If m_ButtonStyle <> eFlatHover Then
            DrawEdge hdc, m_ButtonRect, BDR_SUNKEN95, BF_RECT
            If m_bShowFocus And m_bHasFocus And m_ButtonStyle = eStandard Then
                DrawRectangle 4, 4, lw - 7, lh - 7, TranslateColor(vbApplicationWorkspace)
            End If
        End If
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        ' --Draws flat raised edge border
        Select Case m_ButtonStyle
        Case eStandard
            DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
        Case eFlat
            DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
        End Select
    Case eStateOver
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        Select Case m_ButtonStyle
        Case eFlatHover, eFlat
            ' --Draws flat raised edge border
            DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
        Case Else
            ' --Draws 3d raised edge border
            DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
        End Select
    Case eStateDown
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        Select Case m_ButtonStyle
        Case eStandard
            DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(&H99A8AC)
            DrawRectangle 0, 0, lw, lh, TranslateColor(vbBlack)
        Case e3DHover
            DrawEdge hdc, m_ButtonRect, BDR_SUNKEN95, BF_RECT
        Case eFlatHover, eFlat
            ' --Draws flat pressed edge
            DrawRectangle 0, 0, lw, lh, TranslateColor(vbWhite)
            DrawRectangle 0, 0, lw + 1, lh + 1, TranslateColor(vbGrayText)
        End Select
    End Select
    ' --Button has focus but not downstate Or button is Default
    If m_bHasFocus Or m_bDefault Then
        If m_bShowFocus And Ambient.UserMode Then
            If m_ButtonStyle = e3DHover Or m_ButtonStyle = eStandard Then
                SetRect FocusRect, 4, 4, lw - 4, lh - 4
            Else
                SetRect FocusRect, 3, 3, lw - 3, lh - 3
            End If
            If m_bParentActive Then
                DrawFocusRect hdc, FocusRect
            End If
        End If
        If vState <> eStateDown And m_ButtonStyle = eStandard Then
            SetRect tmpRect, 0, 0, lw - 1, lh - 1
            DrawEdge hdc, tmpRect, BDR_RAISED95, BF_RECT
            DrawRectangle 0, 0, lw - 1, lh - 1, TranslateColor(vbApplicationWorkspace)
            DrawRectangle 0, 0, lw, lh, TranslateColor(vbBlack)
        End If
    End If
End Sub

Private Sub DrawXPToolbar(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    Dim bColor As Long
    lh = ScaleHeight
    lw = ScaleWidth
    UserControl.BackColor = Ambient.BackColor
    bColor = TranslateColor(m_bColors.tBackColor)
    If m_bCheckBoxMode And m_bValue Then
        ' --Check with XP Toolbar!
        If m_bIsDown Then vState = eStateDown
    End If
    If m_bCheckBoxMode And m_bValue And vState <> eStateDown Then
        SetRect lpRect, 0, 0, lw, lh
        PaintRect TranslateColor(&HFEFEFE), lpRect
        m_bColors.tForeColor = TranslateColor(vbButtonText)
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HAF987A)
        DrawCorners ShiftColor(TranslateColor(&HC1B3A0), -0.2)
        If vState = eStateOver Then
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HEDF0F2)  'Right Line
            DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HD8DEE4)   'Bottom
            DrawLineApi 1, lh - 3, lw - 1, lh - 3, TranslateColor(&HE8ECEF)  'Bottom
            DrawLineApi 1, lh - 4, lw - 1, lh - 4, TranslateColor(&HF8F9FA)   'Bottom
        End If
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        PaintRect bColor, m_ButtonRect
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(&HFDFEFE), TranslateColor(&HEEF4F4), gdVertical
        DrawGradientEx 0, lh / 2, lw, lh / 2, TranslateColor(&HEEF4F4), TranslateColor(&HEAF1F1), gdVertical
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE0E7EA)    'right line
        DrawLineApi lw - 3, 2, lw - 3, lh - 2, TranslateColor(&HEAF0F0)
        DrawLineApi 0, lh - 4, lw, lh - 4, TranslateColor(&HE5EDEE)    'Bottom
        DrawLineApi 0, lh - 3, lw, lh - 3, TranslateColor(&HD6E1E4)    'Bottom
        DrawLineApi 0, lh - 2, lw, lh - 2, TranslateColor(&HC6D2D7)    'Bottom
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HC3CECE)
        DrawCorners ShiftColor(TranslateColor(&HC9D4D4), -0.05)
    Case eStateDown
        PaintRect TranslateColor(&HDDE4E5), m_ButtonRect             'Paint with Darker color
        DrawLineApi 1, 1, lw - 2, 1, ShiftColor(TranslateColor(&HD1DADC), -0.02)      'Topmost Line
        DrawLineApi 1, 2, lw - 2, 2, ShiftColor(TranslateColor(&HDAE1E3), -0.02)      'A lighter top line
        DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(TranslateColor(&HDEE5E6), 0.02)    'Bottom Line
        DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(TranslateColor(&HE5EAEB), 0.02)
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H929D9D)
        DrawCorners ShiftColor(TranslateColor(&HABB4B5), -0.2)
    End Select
    If vState = eStateDown Then
        m_bColors.tForeColor = TranslateColor(vbWhite)
    Else
        m_bColors.tForeColor = TranslateColor(vbButtonText)
    End If
End Sub

Private Sub DrawWinXPButton(ByVal vState As enumButtonStates)
'****************************************************************************
'* Windows XP Button                                                        *
'* I made this in just 4 hours                                              *
'* Totally written from Scratch and coded by Me!!                           *
'****************************************************************************
    Dim lpRect As RECT
    Dim bColor As Long
    lh = ScaleHeight
    lw = ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    SetRect m_ButtonRect, 0, 0, lw, lh
    If Not m_bEnabled Then
        CreateRegion
        PaintRect ShiftColor(bColor, 0.03), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.1)
        DrawCorners ShiftColor(bColor, 0.2)
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
        DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.2), ShiftColor(bColor, 0.08), gdVertical
        DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, -0.09)    'BottomMost line
        DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(bColor, -0.05)    'Bottom Line
        DrawLineApi 1, lh - 4, lw - 2, lh - 4, ShiftColor(bColor, -0.01)    'Bottom Line
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, -0.08)    'Right Line
        DrawLineApi 1, 1, 1, lh - 2, BlendColors(TranslateColor(vbWhite), (bColor))    'Left Line
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
        DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.2), ShiftColor(bColor, 0.08), gdVertical
        DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&H89D8FD)       'uppermost inner hover
        DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HCFF0FF)       'uppermost outer hover
        DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&H49BDF9)       'Leftmost Line
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&H49BDF9)    'Rightmost Line
        DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&H7AD2FC)       'Left Line
        DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&H7AD2FC)    'Right Line
        DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&H30B3F8)    'BottomMost Line
        DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&H97E5&)    'Bottom Line
    Case eStateDown
        PaintRect ShiftColor(bColor, -0.05), m_ButtonRect           'Paint with Darker color
        DrawLineApi 1, 1, lw - 2, 1, ShiftColor(bColor, -0.16)      'Topmost Line
        DrawLineApi 1, 2, lw - 2, 2, ShiftColor(bColor, -0.1)      'A lighter top line
        DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, 0.07)    'Bottom Line
        DrawLineApi 1, 1, 1, lh - 2, ShiftColor(bColor, -0.16)    'Leftmost Line
        DrawLineApi 2, 2, 2, lh - 2, ShiftColor(bColor, -0.1)    'Left1 Line
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, 0.04)    'Right Line
    End Select
    If m_bParentActive Then
        If (m_bHasFocus Or m_bDefault) And (vState <> eStateDown And vState <> eStateOver) Then
            DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&HF6D4BC)           'uppermost inner hover
            DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HFFE7CE)           'uppermost outer hover
            DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&HE6AF8E)           'Leftmost Line
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE6AF8E)    'Rightmost Line
            DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&HF4D1B8)           'Left Line
            DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&HF4D1B8)    'Right Line
            DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&HE4AD89)    'BottomMost Line
            DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HEE8269)    'Bottom Line
        End If
    End If
    On Error Resume Next
    If m_bParentActive Then
        If m_bShowFocus And m_bParentActive And (m_bHasFocus Or m_bDefault) Then  'show focusrect at runtime only
            SetRect lpRect, 2, 2, lw - 2, lh - 2     'I don't like this ugly focusrect!!
            DrawFocusRect hdc, lpRect
        End If
    End If
    DrawRectangle 0, 0, lw, lh, TranslateColor(&H743C00)
    DrawCorners ShiftColor(TranslateColor(&H743C00), 0.3)
End Sub

Private Sub DrawVisualStudio2005(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    Dim bColor As Long
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    SetRect m_ButtonRect, 0, 0, lw, lh
    If Not m_bEnabled Then
        DrawGradientEx 0, 0, lw, lh, BlendColors(ShiftColor(bColor, 0.26), TranslateColor(vbWhite)), bColor, gdVertical
    End If
    If m_bCheckBoxMode And m_bValue Then
        PaintRect TranslateColor(&HE8E6E1), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(&H6F4B4B), 0.05)
        If vState = eStateOver Then
            PaintRect TranslateColor(&HE2B598), m_ButtonRect
            DrawRectangle 0, 0, lw, lh, TranslateColor(&HC56A31)
        End If
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        DrawGradientEx 0, 0, lw, lh, BlendColors(ShiftColor(bColor, 0.26), TranslateColor(vbWhite)), bColor, gdVertical
    Case eStateOver
        PaintRect TranslateColor(&HEED2C1), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HC56A31)
    Case eStateDown
        PaintRect TranslateColor(&HE2B598), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H6F4B4B)
    End Select
End Sub

Private Sub DrawAOLButton(ByVal vState As enumButtonStates)
'****************************************************************************
'* AOL (American Online) buttons.                                           *
'****************************************************************************
    Dim lpRect As RECT
    Dim FocusRect As RECT
    Dim bColor As Long
    bColor = TranslateColor(m_bColors.tBackColor)
    If Not m_bEnabled Then                   'Draw Disabled button
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        On Error GoTo H:
        UserControl.BackColor = Ambient.BackColor    'Transparent?!?
        ' --Shadows
        DrawRectangle 6, 6, lw - 9, lh - 9, TranslateColor(&H808080)
        DrawRectangle 5, 5, lw - 7, lh - 7, TranslateColor(&HA0A0A0)
        DrawRectangle 4, 4, lw - 5, lh - 5, TranslateColor(&HC0C0C0)
        SetRect lpRect, 0, 0, lw - 5, lh - 5
        PaintRect bColor, lpRect
        DrawRectangle 0, 0, lw - 4, lh - 4, ShiftColor(bColor, 0.3)
    Case eStateOver
        UserControl.BackColor = Ambient.BackColor
        ' --Shadows
        DrawRectangle 6, 6, lw - 9, lh - 9, TranslateColor(&H808080)
        DrawRectangle 5, 5, lw - 7, lh - 7, TranslateColor(&HA0A0A0)
        DrawRectangle 4, 4, lw - 5, lh - 5, TranslateColor(&HC0C0C0)
        SetRect lpRect, 0, 0, lw - 5, lh - 5
        PaintRect bColor, lpRect
        DrawRectangle 0, 0, lw - 4, lh - 4, ShiftColor(bColor, 0.3)
    Case eStateDown
        UserControl.BackColor = Ambient.BackColor
        SetRect lpRect, 3, 3, lw, lh
        PaintRect bColor, lpRect
        DrawRectangle 3, 3, lw - 3, lh - 3, ShiftColor(bColor, 0.3)
    End Select
    If m_bParentActive Then
        If m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
            UserControl.DrawMode = 6        'For exact AOL effect
            If vState = eStateDown Then
                SetRect lpRect, 6, 6, lw - 3, lh - 3
            Else
                SetRect lpRect, 3, 3, lw - 6, lh - 6
            End If
            DrawFocusRect hdc, lpRect
        End If
    End If
H:
    'Client Site not available (Error in Ambient.BackColor) rarely occurs
End Sub

Private Sub DrawInstallShieldButton(ByVal vState As enumButtonStates)
'****************************************************************************
'* I saw this style while installing JetAudio in my PC.                     *
'* I liked it, so I implemented and gave it a name 'InstallShield'          *
'* hehe .....
'****************************************************************************
    Dim FocusRect As RECT
    Dim lpRect As RECT
    lh = ScaleHeight
    lw = ScaleWidth
    If Not m_bEnabled Then
        vState = eStateNormal                 'Simple draw normal state for Disabled
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        SetRect m_ButtonRect, 0, 0, lw, lh    'Maybe have changed before!
        ' --Draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh, TranslateColor(m_bColors.tBackColor), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.2)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.25)
    Case eStateOver
        ' --Draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh, TranslateColor(m_bColors.tBackColor), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.2)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.25)
    Case eStateDown
        ' --draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1), gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1), ShiftColor(TranslateColor(m_bColors.tBackColor), -0.05), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.23)
        DrawCorners ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.4)
    End Select
    DrawCorners ShiftColor(TranslateColor(m_bColors.tBackColor), 0.05)
    If m_bParentActive And m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
        DrawFocusRect hdc, m_CapRect
    End If
End Sub

Private Sub DrawGelButton(ByVal vState As enumButtonStates)
'****************************************************************************
' Draws a Gelbutton                                                         *
'****************************************************************************
    Dim lpRect As RECT                          'RECT to fill regions
    Dim bColor As Long                          'Original backcolor
    lh = ScaleHeight
    lw = ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    Select Case vState
    Case eStateNormal                            'Normal State
        CreateRegion
        ' --Fill the button region with backVungDatVang color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect bColor, lpRect
        ' --Make a shining Upper Light
        DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.1), bColor, gdVertical
        DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.05), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.1)), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.33)
    Case eStateOver
        ' --Fill the button region with backVungDatVang color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, 0.05), lpRect
        ' --Make a shining Upper Light
        DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(ShiftColor(bColor, 0.05), TranslateColor(vbWhite)), 0.15), ShiftColor(bColor, 0.05), gdVertical
        DrawGradientEx 0, 6, lw, lh - 1, bColor, BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.15)), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.28)
    Case eStateDown
        ' --fill the button region with backVungDatVang color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, -0.03), lpRect
        ' --Make a shining Upper Light
        DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.1), bColor, gdVertical
        DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.08), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.07)), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.36)
    End Select
    DrawCorners ShiftColor(bColor, -0.5)
End Sub

Private Sub DrawVistaToolbarStyle(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    Dim FocusRect As RECT
    lh = ScaleHeight
    lw = ScaleWidth
    If Not m_bEnabled Then
        ' --Draw Disabled button
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        DrawCorners TranslateColor(m_bColors.tBackColor)
        Exit Sub
    End If
    If vState = eStateNormal Then
        CreateRegion
        ' --Set the rect to fill back color
        SetRect lpRect, 0, 0, lw, lh
        ' --Simply fill the button with one color (No gradient effect here!!)
        PaintRect TranslateColor(m_bColors.tBackColor), lpRect
    ElseIf vState = eStateOver Then
        ' --Draws a gradient effect with the folowing colors
        DrawGradientEx 1, 1, lw - 2, lh - 2, TranslateColor(&HFDF9F1), TranslateColor(&HF8ECD0), gdVertical
        ' --Draws a gradient in half region to give a Light Effect
        DrawGradientEx 1, lh / 1.7, lw - 2, lh - 2, TranslateColor(&HF8ECD0), TranslateColor(&HF8ECD0), gdVertical
        ' --Draw outside borders
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HCA9E61)
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)
    ElseIf vState = eStateDown Then
        DrawGradientEx 1, 1, lw - 2, lh - 2, TranslateColor(&HF1DEB0), TranslateColor(&HF9F1DB), gdVertical
        ' --Draws outside borders
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HCA9E61)
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)
    End If
    If vState = eStateDown Or vState = eStateOver Then
        DrawCorners ShiftColor(TranslateColor(&HCA9E61), 0.3)
    End If
End Sub

Private Sub DrawVistaButton(ByVal vState As enumButtonStates)
'*************************************************************************
'* Draws a cool Vista Aero Style Button                                  *
'* Use a light backVungDatVang color for best result                          *
'*************************************************************************
    Dim lpRect As RECT        'Used to set rect for drawing rectangles
    Dim Color1 As Long        'Shifted / Blended color
    Dim bColor As Long        'Original back Color
    lh = ScaleHeight
    lw = ScaleWidth
    Color1 = ShiftColor(TranslateColor(m_bColors.tBackColor), 0.05)
    bColor = TranslateColor(m_bColors.tBackColor)
    If Not m_bEnabled Then
        ' --Draw the Disabled Button
        CreateRegion
        ' --Fill the button with disabled color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, 0.03), lpRect
        ' --Draws outside disabled color rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.25)
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.25)
        DrawCorners ShiftColor(bColor, -0.03)
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        ' --Draws a gradient in the full region
        DrawGradientEx 1, 1, lw - 1, lh, Color1, bColor, gdVertical
        ' --Draws a gradient in half region to give a glassy look
        DrawGradientEx 1, lh / 2, lw - 2, lh - 2, ShiftColor(bColor, -0.02), ShiftColor(bColor, -0.15), gdVertical
        ' --Draws border rectangle
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H707070)    'outer
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)    'inner
    Case eStateOver
        ' --Make gradient in the upper half region
        DrawGradientEx 1, 1, lw - 2, lh / 2, TranslateColor(&HFFF7E4), TranslateColor(&HFFF3DA), gdVertical
        ' --Draw gradient in half button downside to give a glass look
        DrawGradientEx 1, lh / 2, lw - 2, lh - 2, TranslateColor(&HFFE9C1), TranslateColor(&HFDE1AE), gdVertical
        ' --Draws left side gradient effects horizontal
        DrawGradientEx 1, 3, 5, lh / 2 - 2, TranslateColor(&HFFEECD), TranslateColor(&HFFF7E4), gdHorizontal    'Left
        DrawGradientEx 1, lh / 2, 5, lh - (lh / 2) - 1, TranslateColor(&HFAD68F), ShiftColor(TranslateColor(&HFDE1AC), 0.01), gdHorizontal    'Left
        ' --Draws right side gradient effects horizontal
        DrawGradientEx lw - 6, 3, 5, lh / 2 - 2, TranslateColor(&HFFF7E4), TranslateColor(&HFFEECD), gdHorizontal    'Right
        DrawGradientEx lw - 6, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HFDE1AC), 0.01), TranslateColor(&HFAD68F), gdHorizontal    'Right
        ' --Draws border rectangle
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HA77532)    'outer
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)    'inner
    Case eStateDown
        ' --Draw a gradent in full region
        DrawGradientEx 1, 1, lw - 1, lh, TranslateColor(&HF6E4C2), TranslateColor(&HF6E4C2), gdVertical
        ' --Draw gradient in half button downside to give a glass look
        DrawGradientEx 1, lh / 2, lw - 2, lh - 2, TranslateColor(&HF0D29A), TranslateColor(&HF0D29A), gdVertical
        ' --Draws down rectangle
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H5C411D)    '
        DrawLineApi 1, 1, lw - 1, 1, TranslateColor(&HB39C71)    '\Top Lines
        DrawLineApi 1, 2, lw - 1, 2, TranslateColor(&HD6C6A9)    '/
        DrawLineApi 1, 3, lw - 1, 3, TranslateColor(&HECD9B9)    '
        DrawLineApi 1, 1, 1, lh / 2 - 1, TranslateColor(&HCFB073)    'Left upper
        DrawLineApi 1, lh / 2, 1, lh - (lh / 2) - 1, TranslateColor(&HC5912B)    'Left Bottom
        ' --Draws left side gradient effects horizontal
        DrawGradientEx 1, 3, 5, lh / 2 - 2, ShiftColor(TranslateColor(&HE6C891), 0.02), ShiftColor(TranslateColor(&HF6E4C2), -0.01), gdHorizontal    'Left
        DrawGradientEx 1, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HDCAB4E), 0.02), ShiftColor(TranslateColor(&HF0D29A), -0.01), gdHorizontal    'Left
        ' --Draws right side gradient effects horizontal
        DrawGradientEx lw - 6, 3, 5, lh / 2 - 2, ShiftColor(TranslateColor(&HF6E4C2), -0.01), ShiftColor(TranslateColor(&HE6C891), 0.02), gdHorizontal    'Right
        DrawGradientEx lw - 6, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HF0D29A), -0.01), ShiftColor(TranslateColor(&HDCAB4E), 0.02), gdHorizontal    'Right
    End Select
    ' --Draw a focus rectangle if button has focus
    If m_bParentActive Then
        If (m_bHasFocus Or m_bDefault) And vState = eStateNormal Then
            ' --Draw darker outer rectangle
            DrawRectangle 0, 0, lw, lh, TranslateColor(&HA77532)
            ' --Draw light inner rectangle
            DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(&HFBD848)
        End If
        If (m_bShowFocus And m_bHasFocus) Then
            SetRect lpRect, 1.5, 1.5, lw - 2, lh - 2
            DrawFocusRect hdc, lpRect
        End If
    End If
    ' --Create four corners which will be common to all states
    DrawCorners TranslateColor(&HBE965F)
End Sub

Private Sub DrawOutlook2007(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    Dim bColor As Long
    lh = ScaleHeight
    lw = ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    If m_bCheckBoxMode And m_bValue Then
        DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&HA9D9FF), TranslateColor(&H6FC0FF), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H3FABFF), TranslateColor(&H75E1FF), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        If m_bMouseInCtl Then
            DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&H58C1FF), TranslateColor(&H51AFFF), gdVertical
            DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H468FFF), TranslateColor(&H5FD3FF), gdVertical
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        End If
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        PaintRect bColor, m_ButtonRect
        DrawGradientEx 0, 0, lw, lh / 2.7, BlendColors(ShiftColor(bColor, 0.09), TranslateColor(vbWhite)), BlendColors(ShiftColor(bColor, 0.07), bColor), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), bColor, ShiftColor(bColor, 0.03), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&HE1FFFF), TranslateColor(&HACEAFF), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H67D7FF), TranslateColor(&H99E4FF), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    Case eStateDown
        DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&H58C1FF), TranslateColor(&H51AFFF), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H468FFF), TranslateColor(&H5FD3FF), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    End Select
End Sub

Private Sub DrawAquaButton(ByVal vState As enumButtonStates)
    Dim bColor As Long, i As Long, j As Long    'Temp vars
    Dim hHlrgn As Long                      'Highlight shining region
    Dim lrgn As Long                        'Button round region
    Dim ptColor As Long                     'Hightlighting color
    Dim lColor As Long                      'temp color
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    ' --Checkboxmode ??
    If m_bCheckBoxMode And m_bValue Then
        vState = eStateDown
    End If
    If m_bDefault Or m_bHasFocus Then           'MAC OS X draws
        If vState = eStateDown Then             'Hotstate
            vState = eStateDown                 'for focused buttons
        Else
            vState = eStateOver
        End If
    End If
    On Error Resume Next    ' --Forget it!
    UserControl.BackColor = Ambient.BackColor
    Select Case vState
    Case eStateNormal
        CreateRegion
    Case eStateOver
        bColor = TranslateColor(&HD58F54)
    Case eStateDown
        bColor = ShiftColor(TranslateColor(&HD58F54), -0.1)
    End Select
    ' --Region to paint button with backcolor
    lrgn = CreateRoundRectRgn(1, 1, lw + 1, lh - 1, 18, 18)
    PaintRegion lrgn, bColor
    ' --Highlight upper shining region
    hHlrgn = CreateRoundRectRgn(4, 1, lw - 4, 16, 16, 16)
    lColor = bColor
    ' --The for loop below is inspired from Candy Button
    ' --Thanks to Mario (Author of Candy button)
    For j = 0 To lh
        For i = 0 To lw \ 2
            ' --We are inside the button
            If PtInRegion(hHlrgn, i, j) Then
                ptColor = BlendColors(TranslateColor(vbWhite), ShiftColor(ShiftColor(bColor, 0.2), j / 2 * -0.07))
                DrawLineApi i, j, lw - i + 1, j, ptColor
                i = 0               'Starting pos
                j = j + 1           'Draw line downwards
            End If
        Next i
    Next j
    ' --More bottom shine for over and down states
    If vState = eStateNormal Then lColor = ShiftColor(lColor, 0.1)
    If vState = eStateOver Then lColor = ShiftColor(lColor, 0.18)
    If vState = eStateDown Then lColor = ShiftColor(lColor, 0.2)
    ' --Bottom shining layer
    DrawGradientEx 4, 10, lw - 7, lh - 15, bColor, ShiftColor(lColor, 0.08), gdVertical
    ' --Light bottom lines
    DrawLineApi 5, lh - 5, lw - 4, lh - 5, ShiftColor(lColor, 0.08)
    DrawLineApi 6, lh - 4, lw - 5, lh - 4, ShiftColor(lColor, 0.09)
    DrawLineApi 8, lh - 3, lw - 7, lh - 3, ShiftColor(lColor, 0.1)
    ' --Upper lines for some glassy effect
    If vState <> eStateNormal Then
        DrawLineApi 10, 1, lw - 8, 1, TranslateColor(&HFFC9C9)
    Else
        DrawLineApi 10, 1, lw - 8, 1, ShiftColor(lColor, 0.2)
    End If
    '  --Draw corners
    SetPixel hdc, 8, 1, ShiftColor(bColor, -0.1)        '--Upper left
    SetPixel hdc, lw - 8, 1, ShiftColor(bColor, -0.1)   '--Upper Right
    ' --Draw the four side lines
    If Not m_bEnabled Then bColor = ShiftColor(bColor, 0.2)
    DrawLineApi 8, 0, lw, 0, ShiftColor(bColor, -0.34)         'top
    DrawLineApi 8, lh - 2, lw - 7, lh - 2, ShiftColor(bColor, -0.34)    'botom
    DrawLineApi lw - 1, 9, lw - 1, lh - 9, ShiftColor(bColor, -0.34)  'right
    DrawLineApi 1, 9, 1, lh - 9, ShiftColor(bColor, -0.34)          'left
    DrawAquaCorners ShiftColor(bColor, -0.4)
    ' --Clean up
    DeleteObject lrgn
End Sub

Private Sub DrawAquaCorners(ByVal lColor As Long)
' --Draw the corners
' --I m working by pixel method. It is really tedious and not efficient!
' --There may be other alternatives but I have no idea!
' --sorry.... :(
    SetPixel hdc, 7, 1, lColor  '=======================Top Left
    SetPixel hdc, 6, 1, lColor: SetPixel hdc, 5, 2, lColor
    SetPixel hdc, 4, 3, lColor: SetPixel hdc, 4, 4, lColor
    SetPixel hdc, 3, 5, lColor: SetPixel hdc, 2, 6, lColor
    SetPixel hdc, 2, 7, lColor: SetPixel hdc, 1, 8, lColor
    SetPixel hdc, lw - 7, 1, lColor    '===================Top Right
    SetPixel hdc, lw - 6, 1, lColor: SetPixel hdc, lw - 5, 2, lColor
    SetPixel hdc, lw - 4, 3, lColor: SetPixel hdc, lw - 4, 4, lColor
    SetPixel hdc, lw - 3, 5, lColor: SetPixel hdc, lw - 2, 6, lColor
    SetPixel hdc, lw - 2, 7, lColor: SetPixel hdc, lw - 1, 8, lColor
    SetPixel hdc, 7, lh - 3, lColor    '-----------------------------Bottom Left
    SetPixel hdc, 6, lh - 3, lColor: SetPixel hdc, 5, lh - 4, lColor
    SetPixel hdc, 4, lh - 5, lColor: SetPixel hdc, 4, lh - 6, lColor
    SetPixel hdc, 3, lh - 7, lColor: SetPixel hdc, 2, lh - 8, lColor
    SetPixel hdc, 2, lh - 9, lColor: SetPixel hdc, 1, lh - 10, lColor
    SetPixel hdc, lw - 7, lh - 3, lColor    '------------------------Bottom Right
    SetPixel hdc, lw - 6, lh - 3, lColor: SetPixel hdc, lw - 5, lh - 4, lColor
    SetPixel hdc, lw - 4, lh - 5, lColor: SetPixel hdc, lw - 4, lh - 6, lColor
    SetPixel hdc, lw - 3, lh - 7, lColor: SetPixel hdc, lw - 2, lh - 8, lColor
    SetPixel hdc, lw - 2, lh - 9, lColor: SetPixel hdc, lw - 1, lh - 10, lColor
End Sub

Private Sub DrawOffice2003(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    Dim bColor As Long
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    SetRect m_ButtonRect, 0, 0, lw, lh
    If m_bCheckBoxMode And m_bValue Then
        If m_bMouseInCtl Then
            DrawGradientEx 0, 0, lw, lh, TranslateColor(&H4E91FE), TranslateColor(&H8ED3FF), gdVertical
        Else
            DrawGradientEx 0, 0, lw, lh, TranslateColor(&H8CD5FF), TranslateColor(&H55ADFF), gdVertical
        End If
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H800000)
        Exit Sub
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(bColor, 0.05), bColor, gdVertical
        DrawGradientEx 0, lh / 2, lw, lh / 2 + 1, ShiftColor(bColor, -0.01), ShiftColor(bColor, -0.13), gdVertical
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh, TranslateColor(&HCCF4FF), TranslateColor(&H91D0FF), gdVertical
    Case eStateDown
        DrawGradientEx 0, 0, lw, lh, TranslateColor(&H4E91FE), TranslateColor(&H8ED3FF), gdVertical
    End Select
    If m_Buttonstate <> eStateNormal Then
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H800000)
    End If
End Sub

Private Sub DrawSleekButton(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    Dim bColor As Long
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    If m_bCheckBoxMode And m_bValue Then
        vState = eStateDown
    End If
    Select Case vState
    Case eStateNormal
        CreateRegion
        DrawGradientEx 0, 0, lw, lh, bColor, ShiftColor(bColor, 0.08), gdHorizontal
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh, bColor, ShiftColor(bColor, 0.08), gdHorizontal
    Case eStateDown
        DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, -0.05), ShiftColor(bColor, 0.03), gdHorizontal
    End Select
    DrawRectangle 0, 0, lw, lh, TranslateColor(&H7D807E)
    ' --Button has focus  button is Default
    If m_bHasFocus Or m_bDefault Then
        If m_bShowFocus And Ambient.UserMode Then
            SetRect lpRect, 3, 3, lw - 3, lh - 3
            If m_bParentActive Then
                DrawFocusRect hdc, lpRect
            End If
        End If
    End If
End Sub

Private Sub PaintRegion(ByVal lrgn As Long, ByVal lColor As Long)
'Fills a specified region with specified color
    Dim hBrush As Long
    Dim hOldBrush As Long
    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(hdc, hBrush)
    FillRgn hdc, lrgn, hBrush
    SelectObject hdc, hOldBrush
    DeleteObject hBrush
End Sub

Private Sub PaintRect(ByVal lColor As Long, lpRect As RECT)
'Fills a region with specified color
    Dim hOldBrush As Long
    Dim hBrush As Long
    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(UserControl.hdc, hBrush)
    FillRect UserControl.hdc, lpRect, hBrush
    SelectObject UserControl.hdc, hOldBrush
    DeleteObject hBrush
End Sub

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long
'****************************************************************************
'* This routine shifts a color value specified by PercentInDecimal          *
'* Function inspired from DCbutton                                          *
'* All Credits goes to Noel Dacara                                          *
'* A Littlebit modified by me                                               *
'****************************************************************************
    Dim r As Long
    Dim g As Long
    Dim B As Long
    '  Add or remove a certain color quantity by how many percent.
    r = Color And 255
    g = (Color \ 256) And 255
    B = (Color \ 65536) And 255
    r = r + PercentInDecimal * 255       ' Percent should already
    g = g + PercentInDecimal * 255       ' be translated.
    B = B + PercentInDecimal * 255       ' Ex. 50% -> 50 / 100 = 0.5
    '  When overflow occurs, ....
    If (PercentInDecimal > 0) Then       ' RGB values must be between 0-255 only
        If (r > 255) Then r = 255
        If (g > 255) Then g = 255
        If (B > 255) Then B = 255
    Else
        If (r < 0) Then r = 0
        If (g < 0) Then g = 0
        If (B < 0) Then B = 0
    End If
    ShiftColor = r + 256& * g + 65536 * B    ' Return shifted color value
End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If m_bEnabled Then                           'Disabled?? get out!!
        If m_bIsSpaceBarDown Then
            m_bIsSpaceBarDown = False
            m_bIsDown = False
        End If
        If m_bCheckBoxMode Then                'Checkbox Mode?
            If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub    'Checkboxes dont repond to Enter/Escape'
            m_bValue = Not m_bValue             'Change Value (Checked/Unchecked)
            If Not m_bValue Then                'If value unchecked then
                m_Buttonstate = eStateNormal     'Normal State
            End If
            RedrawButton
        End If
        DoEvents                               'To remove focus from other button and Do events before click event
        RaiseEvent Click                       'Now Raiseevent
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    m_bDefault = Ambient.DisplayAsDefault
    If Not m_bEnabled Or m_bMouseInCtl Then Exit Sub
    If PropertyName = "DisplayAsDefault" Then
        RedrawButton
    End If
    If PropertyName = "BackColor" Then
        RedrawButton
    End If
End Sub

Private Sub UserControl_DblClick()
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If
    If m_lDownButton = 1 Then                    'React to only Left button
        SetCapture (hwnd)                         'Preserve Hwnd on DoubleClick
        If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        RedrawButton
        UserControl_MouseDown m_lDownButton, m_lDShift, m_lDX, m_lDY
        RaiseEvent DblClick
    End If
End Sub

Private Sub UserControl_GotFocus()
    m_bHasFocus = True
    If m_bMouseInCtl Then
        If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
    Else
        If Not m_bIsDown Then m_Buttonstate = eStateNormal
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim OS As OSVERSIONINFO
    ' --Get the operating system version for text drawing purposes.
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    m_WindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub

Private Sub UserControl_InitProperties()
'Initialize Properties for User Control
'Called on designtime everytime a control is added
    m_ButtonStyle = eStandard
    m_bShowFocus = True
    m_bEnabled = True
    m_Caption = Ambient.DisplayName
    UserControl.FontName = "Tahoma"
    m_PictureAlign = epLeftOfCaption
    m_bUseMaskColor = False
    m_lMaskColor = &HE0E0E0
    m_CaptionAlign = ecCenterAlign
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    SetThemeColors
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13                                'Enter Key
        RaiseEvent Click
    Case 37, 38                            'Left and Up Arrows
        SendKeys "+{TAB}"                  'Button should transfer focus to other ctl
    Case 39, 40                            'Right and Down Arrows
        SendKeys "{TAB}"                   'Button should transfer focus to other ctl
    Case 32                                'SpaceBar held down
        If Not m_bIsDown Then
            If Shift = 4 Then Exit Sub     'System Menu Should pop up
            m_bIsSpaceBarDown = True       'Set space bar as pressed
            If (m_bCheckBoxMode) Then      'Is CheckBoxMode??
                m_bValue = Not m_bValue    'Toggle Check Value
            Else
                If m_Buttonstate <> eStateDown Then
                    m_Buttonstate = eStateDown    'Button state should be down
                    RedrawButton
                End If
            End If
        End If
        If (Not GetCapture = UserControl.hwnd) Then
            ReleaseCapture
            SetCapture UserControl.hwnd    'No other processing until spacebar is released
        End If                          'Thanks to APIGuide
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If m_bMouseInCtl And m_bIsDown Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
            RedrawButton
        ElseIf m_bMouseInCtl And Not m_bIsDown Then   'If spacebar released over ctl
            If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver    'Draw Hover State
            RedrawButton
            RaiseEvent Click
        Else                                         'If Spacebar released outside ctl
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            RedrawButton
            RaiseEvent Click
        End If
        RaiseEvent KeyUp(KeyCode, Shift)
        m_bIsSpaceBarDown = False
        m_bIsDown = False
    End If
End Sub

Private Sub UserControl_LostFocus()
    m_bHasFocus = False                                 'No focus
    m_bIsDown = False                                   'No down state
    m_bIsSpaceBarDown = False                           'No spacebar held
    If Not m_bParentActive Then
        If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
    ElseIf m_bMouseInCtl Then
        If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
    Else
        If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
    End If
    RedrawButton
    If m_bDefault Then                                  'If default button,
        RedrawButton                                    'Show Focus
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    m_lDownButton = Button                       'Button pressed for Dblclick
    m_lDX = x
    m_lDY = Y
    m_lDShift = Shift
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If
    If Button = 1 Then
        m_bHasFocus = True
        m_bIsDown = True
        If m_bMouseInCtl Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        End If
        RedrawButton
        RaiseEvent MouseDown(Button, Shift, x, Y)
    End If
End Sub

Private Sub SetThemeColors()
'Sets a style colors to default colors when button initialized
'or whenever you change the style of Button
    With m_bColors
        Select Case m_ButtonStyle
        Case eStandard, eFlat, eVistaToolbar, e3DHover, eFlatHover
            .tBackColor = TranslateColor(vbButtonFace)
        Case eWindowsXP
            .tBackColor = TranslateColor(&HE7EBEC)
        Case eOutlook2007, eGelButton
            .tBackColor = TranslateColor(&HFFD1AD)
            .tForeColor = TranslateColor(&H8B4215)
        Case eXPToolbar
            .tBackColor = TranslateColor(&HECF1F1)
        Case eAOL
            .tBackColor = TranslateColor(&HAA6D00)
            .tForeColor = TranslateColor(vbWhite)
        Case eAqua
            .tBackColor = TranslateColor(&HD0D0D0)
        Case eVistaAero
            .tBackColor = ShiftColor(TranslateColor(&HD4D4D4), 0.06)
        Case eInstallShield
            .tBackColor = TranslateColor(&HE1D6D5)
        Case eVisualStudio
            .tBackColor = TranslateColor(vbButtonFace)
        Case eOffice2003
            .tBackColor = TranslateColor(&HFCE1CA)
        Case eSleek
            .tBackColor = TranslateColor(&HE6CAA7)
        End Select
        If m_ButtonStyle <> eAOL Then .tForeColor = TranslateColor(vbButtonText)
        If m_ButtonStyle = eFlat Or m_ButtonStyle = eSleek Or m_ButtonStyle = eInstallShield Or m_ButtonStyle = eStandard Or m_ButtonStyle = eAOL Then
            m_bShowFocus = True
        Else
            m_bShowFocus = False
        End If
    End With
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim p As POINT
    GetCursorPos p
    If (Not WindowFromPoint(p.x, p.Y) = UserControl.hwnd) Then
        m_bMouseInCtl = False
        RaiseEvent MouseLeave
    End If
    TrackMouseLeave UserControl.hwnd
    If m_bMouseInCtl Then
        If m_bHandPointer Then
            SetCursor m_lCursor
        End If
        If m_bIsDown Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        ElseIf Not m_bIsDown And Not m_bIsSpaceBarDown Then
            If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
        End If
        RedrawButton
    End If
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If
    If Button = vbLeftButton Then
        m_bIsDown = False
        If (x > 0 And Y > 0) And (x < ScaleWidth And Y < ScaleHeight) Then
            If m_bCheckBoxMode Then m_bValue = Not m_bValue
            RedrawButton
            RaiseEvent Click
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Resize()
' --At least, a checkbox will also need this much of size!!!!
    If Height < 220 Then Height = 220
    If Width < 220 Then Width = 220
    ' --On resize, create button region again
    CreateRegion
    RedrawButton
End Sub

Private Sub UserControl_Paint()
    RedrawButton
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_ButtonStyle = .ReadProperty("ButtonStyle", eFlat)
        m_bShowFocus = .ReadProperty("ShowFocusRect", False)    'for eFlat style only
        Set mFont = .ReadProperty("Font", Ambient.Font)
        Set UserControl.Font = mFont
        m_bColors.tBackColor = .ReadProperty("BackColor", TranslateColor(vbButtonFace))
        m_bEnabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", "jcbutton")
        m_bValue = .ReadProperty("Value", False)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)    'vbdefault
        m_bHandPointer = .ReadProperty("HandPointer", False)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_PicOver = .ReadProperty("PictureHover", Nothing)
        m_PicSize = .ReadProperty("PictureSize", epsNormal)
        m_lMaskColor = .ReadProperty("MaskColor", &HE0E0E0)
        m_bUseMaskColor = .ReadProperty("UseMaskCOlor", False)
        m_bCheckBoxMode = .ReadProperty("CheckBoxMode", False)
        m_PictureAlign = .ReadProperty("PictureAlign", epLeftOfCaption)
        m_CaptionAlign = .ReadProperty("CaptionAlign", ecCenterAlign)
        m_bColors.tForeColor = .ReadProperty("ForeColor", TranslateColor(vbButtonText))
        UserControl.ForeColor = m_bColors.tForeColor
        UserControl.Enabled = m_bEnabled
        SetAccessKey
        lh = UserControl.ScaleHeight
        lw = UserControl.ScaleWidth
        m_lParenthWnd = UserControl.Parent.hwnd
    End With
    UserControl_Resize
    If Ambient.UserMode Then                                                              'If we're not in design mode
        If m_bHandPointer Then
            m_lCursor = LoadCursor(0, IDC_HAND)     'Load System Hand pointer
            m_bHandPointer = (Not m_lCursor = 0)
        End If
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        If bTrack Then
            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Subclass_Start .hwnd
                Subclass_Start m_lParenthWnd
                Subclass_AddMsg .hwnd, WM_MOUSEMOVE, MSG_AFTER
                Subclass_AddMsg .hwnd, WM_MOUSELEAVE, MSG_AFTER
                On Error Resume Next
                If UserControl.Parent.MDIChild Then
                    Call Subclass_AddMsg(m_lParenthWnd, WM_NCACTIVATE, MSG_AFTER)
                Else
                    Call Subclass_AddMsg(m_lParenthWnd, WM_ACTIVATE, MSG_AFTER)
                End If
            End With
        End If
    End If
End Sub

'A nice place to stop subclasser
Private Sub UserControl_Terminate()
    On Error GoTo Crash:
    If m_lButtonRgn Then DeleteObject m_lButtonRgn
'    If Ambient.UserMode Then
       'Subclass_Stop m_lParenthWnd
        'Subclass_Stop UserControl.hwnd
       ' Subclass_StopAll                                               'Terminate all subclassing
   ' End If
Crash:
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ButtonStyle", m_ButtonStyle, eFlat
        .WriteProperty "ShowFocusRect", m_bShowFocus, False
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "Font", mFont, Ambient.Font
        .WriteProperty "BackColor", m_bColors.tBackColor, TranslateColor(vbButtonFace)
        .WriteProperty "Caption", m_Caption, "jcbutton1"
        .WriteProperty "ForeColor", m_bColors.tForeColor, TranslateColor(vbButtonText)
        .WriteProperty "CheckBoxMode", m_bCheckBoxMode, False
        .WriteProperty "Value", m_bValue, False
        .WriteProperty "MousePointer", UserControl.MousePointer, 0
        .WriteProperty "HandPointer", m_bHandPointer, False
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "Picture", m_Picture, Nothing
        .WriteProperty "PictureHover", m_PicOver, Nothing
        .WriteProperty "PictureAlign", m_PictureAlign, epLeftOfCaption
        .WriteProperty "pictureSize", m_PicSize, epsNormal
        .WriteProperty "UseMaskCOlor", m_bUseMaskColor, False
        .WriteProperty "MaskColor", m_lMaskColor, &HE0E0E0
        .WriteProperty "CaptionAlign", m_CaptionAlign, ecCenterAlign
    End With
End Sub

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hmod As Long
    Dim bLibLoaded As Boolean
    hmod = GetModuleHandleA(sModule)
    If hmod = 0 Then
        hmod = LoadLibraryA(sModule)
        If hmod Then
            bLibLoaded = True
        End If
    End If
    If hmod Then
        If GetProcAddress(hmod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    If bLibLoaded Then
        FreeLibrary hmod
    End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        If bTrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If
End Sub

'=========================================================================
'PUBLIC ROUTINES including subclassing & public button properties
' CREDITS: Paul Caton
'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'hWnd     - The window handle
'uMsg     - The message number
'wParam   - Message related data
'lParam   - Message related data
'Notes:
'If you really know what you're doing, it's possible to change the values of the
'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'values get passed to the default handler.. and optionaly, the 'after' callback
    Static bMoving As Boolean
    Select Case uMsg
    Case WM_MOUSEMOVE
        If Not m_bMouseInCtl Then
            m_bMouseInCtl = True
            TrackMouseLeave lng_hWnd
            If m_bMouseInCtl Then
                'If Not m_bIsSpaceBarDown Then m_Buttonstate = eStateOver
                If Not m_bIsSpaceBarDown Then
                    m_Buttonstate = eStateOver
                End If
            End If
            RedrawButton
            RaiseEvent MouseEnter
        End If
    Case WM_MOUSELEAVE
        m_bMouseInCtl = False
        If m_bIsSpaceBarDown Then Exit Sub
        If m_bEnabled Then
            m_Buttonstate = eStateNormal
        End If
        RedrawButton
        RaiseEvent MouseLeave
    Case WM_NCACTIVATE, WM_ACTIVATE
        If wParam Then
            m_bParentActive = True
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            If m_bDefault Then
                RedrawButton
            End If
            RedrawButton
        Else
            m_bIsDown = False
            m_bIsSpaceBarDown = False
            m_bHasFocus = False
            m_bParentActive = False
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            RedrawButton
        End If
    End Select
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    MsgBox "JCButton" & vbNewLine & _
           "A Multistyle Button Control" & vbNewLine & vbNewLine & _
           "Created by: Juned S. Chhipa", vbInformation + vbOKOnly, "About"
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_bColors.tBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_bColors.tBackColor = New_BackColor
    RedrawButton
    PropertyChanged "BackColor"
End Property

Public Property Get ButtonStyle() As enumButtonStlyes
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As enumButtonStlyes)
    m_ButtonStyle = New_ButtonStyle
    SetThemeColors          'Set colors
    CreateRegion            'Create Region Again
    RedrawButton            'Obviously, force redraw!!!
    PropertyChanged "ButtonStyle"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    SetAccessKey
    RedrawButton
    PropertyChanged "Caption"
End Property

Public Property Get CaptionAlign() As enumCaptionAlign
    CaptionAlign = m_CaptionAlign
End Property

Public Property Let CaptionAlign(ByVal New_CaptionAlign As enumCaptionAlign)
    m_CaptionAlign = New_CaptionAlign
    RedrawButton
    PropertyChanged "CaptionAlign"
End Property

Public Property Get CheckBoxMode() As Boolean
    CheckBoxMode = m_bCheckBoxMode
End Property

Public Property Let CheckBoxMode(ByVal New_CheckBoxMode As Boolean)
    m_bCheckBoxMode = New_CheckBoxMode
    'If Not m_bCheckBoxMode Then m_Buttonstate = eStateNormal
    If Not m_bCheckBoxMode Then
        m_Buttonstate = eStateNormal
    End If
    RedrawButton
    PropertyChanged "Value"
    PropertyChanged "CheckBoxMode"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
    'UserControl.Enabled = m_enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_bEnabled = New_Enabled
    UserControl.Enabled = m_bEnabled
    RedrawButton
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(ByVal New_Font As StdFont)
    Set mFont = New_Font
    Refresh
    RedrawButton
    PropertyChanged "Font"
    Call mFont_FontChanged("")
End Property

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
    Refresh
    RedrawButton
    PropertyChanged "Font"
End Sub

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_bColors.tForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_bColors.tForeColor = New_ForeColor
    UserControl.ForeColor = m_bColors.tForeColor
    UserControl_Resize
    PropertyChanged "ForeColor"
End Property

Public Property Get HandPointer() As Boolean
    HandPointer = m_bHandPointer
End Property

Public Property Let HandPointer(ByVal New_HandPointer As Boolean)
    m_bHandPointer = New_HandPointer
    PropertyChanged "HandPointer"
End Property

Public Property Get hwnd() As Long
' --Handle that uniquely identifies the control
    hwnd = UserControl.hwnd
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_lMaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_lMaskColor = New_MaskColor
    RedrawButton
    PropertyChanged "MaskColor"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_Icon As IPictureDisp)
    On Error Resume Next
    Set UserControl.MouseIcon = New_Icon
    If (New_Icon Is Nothing) Then
        UserControl.MousePointer = 0     ' vbDefault
    Else
        UserControl.MousePointer = 99     ' vbCustom
    End If
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_Cursor As MousePointerConstants)
    UserControl.MousePointer = New_Cursor
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    If Not New_Picture Is Nothing Then
        RedrawButton
        PropertyChanged "Picture"
    Else
        UserControl_Resize
    End If
End Property

Public Property Get PictureAlign() As enumPictureAlign
    PictureAlign = m_PictureAlign
End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As enumPictureAlign)
    m_PictureAlign = New_PictureAlign
    If Not m_Picture Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "PictureAlign"
End Property

Public Property Get PictureSize() As enumPictureSize
    PictureSize = m_PicSize
End Property

Public Property Let PictureSize(ByVal New_Size As enumPictureSize)
    Dim tmpPic As New StdPicture
    Set tmpPic = m_Picture
    m_PicSize = New_Size
    RedrawButton
    PropertyChanged "PictureSize"
End Property

Public Property Get PictureHover() As StdPicture
    Set PictureHover = m_PicOver
End Property

Public Property Set PictureHover(ByVal New_PictureHover As StdPicture)
    If m_Picture Is Nothing Then
        Set m_Picture = New_PictureHover        'Normal picture essential
        PropertyChanged "Picture"
    End If
    Set m_PicOver = New_PictureHover
    If Not New_PictureHover Is Nothing Then
        RedrawButton
        PropertyChanged "PictureHover"
    Else
        UserControl_Resize
    End If
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_bShowFocus
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_bShowFocus = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"
End Property

Public Property Get UseMaskColor() As Boolean
    UseMaskColor = m_bUseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
    m_bUseMaskColor = New_UseMaskColor
    If Not m_Picture Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "UseMaskColor"
End Property

Public Property Get Value() As Boolean
    Value = m_bValue
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    If m_bCheckBoxMode Then
        m_bValue = New_Value
        'If Not m_bValue Then m_Buttonstate = eStateNormal
        If Not m_bValue Then
            m_Buttonstate = eStateNormal
        End If
        RedrawButton
        PropertyChanged "Value"
    Else
        m_Buttonstate = eStateNormal
        RedrawButton
    End If
End Property

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            zAddMsg uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And eMsgWhen.MSG_AFTER Then
            zAddMsg uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            zDelMsg uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And eMsgWhen.MSG_AFTER Then
            zDelMsg uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
'lng_hWnd  - The handle of the window to be subclassed
'Returns;
'The sc_aSubData() index
    Const CODE_LEN As Long = 200                                         'Length of the machine code in bytes
    Const FUNC_CWP As String = "CallWindowProcA"                         'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM As String = "EbMode"                                  'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"                          'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER As String = "user32"                                  'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5 As String = "vba5"                                    'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"                                    'Location of the EbMode function if running VB6
    Const PATCH_01 As Long = 18                                          'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02 As Long = 68                                          'Address of the previous WndProc
    Const PATCH_03 As Long = 78                                          'Relative address of SetWindowsLong
    Const PATCH_06 As Long = 116                                         'Address of the previous WndProc
    Const PATCH_07 As Long = 121                                         'Relative address of CallWindowProc
    Const PATCH_0A As Long = 186                                         'Address of the owner object
    Static aBuf(1 To CODE_LEN) As Byte                                               'Static code buffer byte array
    Static pCWP As Long                                               'Address of the CallWindowsProc
    Static pEbMode As Long                                               'Address of the EbMode IDE break/stop/running function
    Static pSWL As Long                                               'Address of the SetWindowsLong function
    Dim i As Long                                               'Loop index
    Dim j As Long                                               'Loop index
    Dim nSubIdx As Long                                               'Subclass data index
    Dim sHex As String                                             'Hex code string
'If it's the first time through here..
    If aBuf(1) = 0 Then
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                                'Next pair of hex characters
        'Get API function addresses
        If Subclass_InIDE Then                                                              'If we're running in the VB IDE
            aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                               'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
            End If
        End If
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If
    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd                                                                    'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
        RtlMoveMemory ByVal .nAddrSub, aBuf(1), CODE_LEN                               'Copy the machine code from the static byte array to the code array in sc_aSubData
        zPatchRel .nAddrSub, PATCH_01, pEbMode                                         'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        zPatchVal .nAddrSub, PATCH_02, .nAddrOrig                                      'Original WndProc address for CallWindowProc, call the original WndProc
        zPatchRel .nAddrSub, PATCH_03, pSWL                                            'Patch the relative address of the SetWindowLongA api function
        zPatchVal .nAddrSub, PATCH_06, .nAddrOrig                                      'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        zPatchRel .nAddrSub, PATCH_07, pCWP                                            'Patch the relative address of the CallWindowProc api function
        zPatchVal .nAddrSub, PATCH_0A, ObjPtr(Me)                                      'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
    Do While i >= 0                                                                       'Iterate through each element
        With sc_aSubData(i)
            If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
                Subclass_Stop .hwnd                                                        'Subclass_Stop
            End If
        End With
        i = i - 1                                                                           'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        SetWindowLongA .hwnd, GWL_WNDPROC, .nAddrOrig                                  'Restore the original WndProc
        zPatchVal .nAddrSub, PATCH_05, 0                                               'Patch the Table B entry count to ensure no further 'before' callbacks
        zPatchVal .nAddrSub, PATCH_09, 0                                               'Patch the Table A entry count to ensure no further 'after' callbacks
        GlobalFree .nAddrSub                                                           'Release the machine code memory
        .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                                       'Clear the before table
        .nMsgCntA = 0                                                                       'Clear the after table
        Erase .aMsgTblB                                                                     'Erase the before table
        Erase .aMsgTblA                                                                     'Erase the after table
    End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.
'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long                                                               'Message table entry index
    Dim nOff1 As Long                                                               'Machine code buffer offset 1
    Dim nOff2 As Long                                                               'Machine code buffer offset 2
    If uMsg = ALL_MESSAGES Then                                                           'If all messages
        nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
    Else                                                                                  'Else a specific message number
        Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
                Exit Sub                                                                        'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
                Exit Sub                                                                        'Bail
            End If
        Loop                                                                                'Next entry
        nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
    End If
    If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
        nOff1 = PATCH_04                                                                    'Offset to the Before table
        nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
    Else                                                                                  'Else after
        nOff1 = PATCH_08                                                                    'Offset to the After table
        nOff2 = PATCH_09                                                                    'Offset to the After table entry count
    End If
    If uMsg <> ALL_MESSAGES Then
        zPatchVal nAddr, nOff1, VarPtr(aMsgTbl(1))                                     'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    zPatchVal nAddr, nOff2, nMsgCnt                                                  'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
        nMsgCnt = 0                                                                         'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
            nEntry = PATCH_05                                                                 'Patch the before table message count location
        Else                                                                                'Else after
            nEntry = PATCH_09                                                                 'Patch the after table message count location
        End If
        zPatchVal nAddr, nEntry, 0                                                     'Patch the table message count to zero
    Else                                                                                  'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                           'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
                Exit Do                                                                         'Bail
            End If
        Loop                                                                                'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                                'If we're searching not adding
                    Exit Function                                                                 'Found
                End If
            ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
                If bAdd Then                                                                    'If we're adding
                    Exit Function                                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                                     'Decrement the index
    Loop
    If Not bAdd Then
        Debug.Assert False                                                                  'hWnd not found, programmer error
    End If
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    RtlMoveMemory ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    RtlMoveMemory ByVal nAddr + nOffset, nValue, 4
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

'End of Subclassing routines
'---------------x---------------x--------------x--------------x-----------x---
' Oops! Control resulted Longer than expected!
' Lots of hours and lots of tedious work!   This is my first submission on PSC
' So if you want to vote for this, just do it ;)
' Comments are greatly appreciated...
'---------------x---------------x--------------x--------------x-----------x---
