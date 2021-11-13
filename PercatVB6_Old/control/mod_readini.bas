Attribute VB_Name = "Mod_Readini"
Option Explicit



Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public uKetquamoi As Integer
Public uLevel As Integer
Public uSokyTu As Integer
Public sPicGame As Integer
Public uBatAmThanh As Boolean
Public uBatThoiGian As Boolean
Public uBatXepHang As Boolean
Public uBatHieuUng As Boolean
Public uThoiGian As Integer
Public sTongVatPham As Byte

Public Function WriteIniStr(StandKey As String, keyName As String, keyValue As String, FileName As String) As Long
    Dim lenIniKey As Long
    Dim strkey As String * 255
    Dim lFileName As String
    lFileName = App.Path & "\" & FileName
    
    WriteIniStr = WritePrivateProfileString(StandKey, keyName, keyValue, lFileName)
End Function

Public Function GetIniStr(StandKey As String, keyName As String, Default As String, FileName As String) As String
    Dim lenIniKey As Long
    Dim strkey As String * 255
    Dim lFileName As String
    lFileName = App.Path & "\" & FileName
    
    lenIniKey = GetPrivateProfileString(StandKey, keyName, Default, strkey, Len(strkey), lFileName)
    If lenIniKey <> 0 Then GetIniStr = Left$(strkey, InStr(strkey, Chr(0)) - 1)
End Function
Public Sub LoadLevel(ByVal Level As Integer)
    Dim iGold As Byte
    If FileExists(App.Path & "\Data\Level\lv" & Level & ".ptq") = False Then
        LoadLevel (1)
        uLevel = 1
    Else
        DoDecrypt (App.Path & "\Data\Level\lv" & Level & ".ptq") 'giai ma map
        'doc map
        FrmMain.LblMucTieu = GetIniStr("Main", "MucTieu", "", "\Data\Level\lv" & Level & ".ptq")
        uThoiGian = GetIniStr("Main", "ThoiGian", "", "\Data\Level\lv" & Level & ".ptq")
        sPicGame = GetIniStr("Main", "Picture", "None", "\Data\Level\lv" & Level & ".ptq")
        sTongVatPham = GetIniStr("Main", "SoVatPham", 0, "Data\Level\lv" & Level & ".ptq")
        'If CStr(sPicGame) <> "None" Or FileExists(App.Path & "\Data\BG\" & sPicGame & ".ptq") = True Then
          '  FrmMain.PicGame.Picture = FrmPic.PicBG(sPicGame).Picture
        'End If
        Call TaoVatPham(0, Level)
        For iGold = 1 To sTongVatPham
            Load FrmMain.gold(iGold)
            Call TaoVatPham(iGold, Level)
        Next iGold
        DoEncrypt (App.Path & "\Data\Level\lv" & Level & ".ptq") 'ma hoa lai
    End If
End Sub
Sub TaoVatPham(ByVal ivatpham As Byte, iLevel As Integer)
    Dim BTarray() As Byte '
    FrmMain.gold(ivatpham).Top = GetIniStr(CStr(ivatpham), "Top", 0, "\Data\Level\lv" & iLevel & ".ptq")
    FrmMain.gold(ivatpham).Left = GetIniStr(CStr(ivatpham), "Left", 0, "\Data\Level\lv" & iLevel & ".ptq")
    FrmMain.gold(ivatpham).Tag = GetIniStr(CStr(ivatpham), "GiaTri", 0, "Data\Level\lv" & iLevel & ".ptq")
    BTarray() = LoadResData(116 + FrmMain.gold(ivatpham).Tag, "VATPHAM")
    FrmMain.gold(ivatpham).LoadImageFromStream BTarray
    FrmMain.gold(ivatpham).Visible = True
End Sub
Public Sub OpenSettings()
    If FileExists(App.Path & "\Settings.ini") = False Then
        WriteNormalSettings True
    Else
        uBatAmThanh = CBool(GetIniStr("Opitions", "BatAmThanh", True, "Settings.ini"))
        uBatThoiGian = CBool(GetIniStr("Opitions", "BatThoiGian", True, "Settings.ini"))
        uBatXepHang = CBool(GetIniStr("Opitions", "BatXepHang", False, "Settings.ini"))
        uBatHieuUng = CBool(GetIniStr("Opitions", "BatHieuUng", False, "Settings.ini"))
        
        uKetquamoi = CInt(GetIniStr("Opitions", "iOnc", 0, "Settings.ini"))
        uLevel = CInt(GetIniStr("Opitions", "Level", 0, "Settings.ini"))
    End If
End Sub
Private Function FileExists(FileName) As Boolean
    On Error GoTo ErrorHandler
    FileExists = (dir(FileName) <> "")
    Exit Function
ErrorHandler:
    FileExists = False
End Function
Public Sub SaveSetting(Optional Update As Boolean = False)
    On Error Resume Next
    If FileExists(App.Path & "\Settings.ini") = False Then
        WriteNormalSettings True
    Else
        WriteIniStr "Opitions", "BatAmThanh", CStr(uBatAmThanh), "Settings.ini"
        WriteIniStr "Opitions", "BatThoiGian", CStr(uBatThoiGian), "Settings.ini"
        WriteIniStr "Opitions", "BatXepHang", CStr(uBatXepHang), "Settings.ini"
        WriteIniStr "Opitions", "BatHieuUng", CStr(uBatHieuUng), "Settings.ini"
        
        WriteIniStr "Opitions", "iOnc", CStr(uKetquamoi), "Settings.ini"
        WriteIniStr "Opitions", "Level", CStr(uLevel), "Settings.ini"
    End If
    If Update Then OpenSettings
End Sub
Public Sub WriteNormalSettings(Optional Update As Boolean = False)
    On Error Resume Next
    If dir(App.Path & "\Settings.ini") <> "" Then Kill App.Path & "\Settings.ini"
    WriteIniStr "Opitions", "BatAmThanh", True, "Settings.ini"
    WriteIniStr "Opitions", "BatThoiGian", True, "Settings.ini"
    WriteIniStr "Opitions", "BatXepHang", False, "Settings.ini"
    WriteIniStr "Opitions", "BatHieuUng", False, "Settings.ini"
    
    WriteIniStr "Opitions", "iOnc", 0, "Settings.ini"
    WriteIniStr "Opitions", "Level", 0, "Settings.ini"
    If Update Then OpenSettings
End Sub
Public Sub DoEncrypt(sNamelang As String)
    ' encrypt file sub
    Dim csCrypt As New clsCrypto
    Dim strFile As String
    Dim lFileLength As Long
    
    ' get length of file to encrypt
    lFileLength = FileLen(sNamelang)
    ' allocate string to hold file
    strFile = String(lFileLength, vbNullChar)
    ' open file in binary
    Open sNamelang For Binary Access Read As #1
    
    Get 1, , strFile
    Close #1
    ' Get password
    csCrypt.password = "phanthequang4101987"
    csCrypt.InBuffer = strFile
    'generate hash of original file
    If Not csCrypt.HashFile Then Exit Sub
    ' generate password
    If Not csCrypt.GeneratePasswordKey Then Exit Sub
    ' encrypt message data
    If Not csCrypt.EncryptFileData Then Exit Sub
    ' destroy key
    csCrypt.DestroySessionKey
    ' check for valid data
    If csCrypt.OutBuffer <> "" Then
        ' delete current data file
        Kill sNamelang
        ' open new file for binary write
        Open sNamelang For Binary Access Write As #2
        ' write encrypted data to file
        Put 2, , csCrypt.OutBuffer
        Close #2    ' close open file
    End If
End Sub

Public Sub DoDecrypt(sNamelang As String)
    ' decrypt file sub
    Dim csCrypt As New clsCrypto
    Dim strFile As String
    Dim lFileLength As String
    
    ' get length of file
    lFileLength = FileLen(sNamelang)
    ' allocate string to hold file
    strFile = String(lFileLength, vbNullChar)
    ' open file in binary mode
    Open sNamelang For Binary Access Read As #1
    
    Get 1, , strFile
    Close #1
    ' set password
    csCrypt.password = "phanthequang4101987"
    csCrypt.InBuffer = strFile
    ' generate password
    If Not csCrypt.GeneratePasswordKey Then Exit Sub
    ' decrypt message
    If Not csCrypt.DecryptFileData Then Exit Sub
    csCrypt.DestroySessionKey
    
    ' check for valid data
    If csCrypt.OutBuffer <> "" Then
        ' delete current file
        Kill sNamelang
        ' creat new file
        Open sNamelang For Binary Access Write As #2
        Put 2, , csCrypt.OutBuffer
        Close #2
    End If
End Sub


