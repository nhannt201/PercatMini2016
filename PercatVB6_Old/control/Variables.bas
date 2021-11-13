Attribute VB_Name = "modgame"
Option Explicit
Public nul As Integer
Public z As Long
Public SoTien, iSoTien, iDiem As Integer
Public BatQuay, sBoomNo As Integer
Public DungBoom As Boolean, DungHoaMayMan As Boolean, DungNuocNumber1 As Boolean, DungNuocNumber2 As Boolean, DungDauNhot As Boolean
Public DungKienThucDa As Boolean, DungKienThucVang As Boolean, DungKienThucKimCuong As Boolean, DungKienThucXuong As Boolean
Public sDangKeo As Integer
Public sTocDoTang As Integer
Public LoaiMay As String
Public TocDo_Keo As Single
Public TocDo_KeoDay As Single
Public DangChay As Boolean
Public Sub Amthanh(Index As Long)
    On Error Resume Next
    Dim URLAmthanh As String
    If uBatAmThanh = True Then
        FrmMain.Wmp.Controls.stop
        Select Case Index
        Case 0: URLAmthanh = "add_score"
        Case 1: URLAmthanh = "alarm"
        Case 2: URLAmthanh = "click"
        Case 3: URLAmthanh = "click_button"
        Case 4: URLAmthanh = "cut-scene"
        Case 5: URLAmthanh = "cut-scene-2"
        Case 6: URLAmthanh = "dig"
        Case 7: URLAmthanh = "eat"
        Case 8: URLAmthanh = "explode_stone" 'pha da bang min
        Case 9: URLAmthanh = "explode_tnt"
        Case 10: URLAmthanh = "explosive"
        Case 11: URLAmthanh = "got_it"
        Case 12: URLAmthanh = "high-value" 'da nho
        Case 13: URLAmthanh = "hook_down" 'tha giay xuong
        Case 14: URLAmthanh = "hook_up" 'keo len
        Case 15: URLAmthanh = "large_gold" 'trung vang
        Case 16: URLAmthanh = "level_end" 'ket thuc lv
        Case 17: URLAmthanh = "level_start" 'bat dau lv moi
        Case 18: URLAmthanh = "low-value" 'cuc da
        Case 19: URLAmthanh = "no"
        Case 20: URLAmthanh = "normal-value" 'da thuong
        Case 21: URLAmthanh = "pull"
        Case 22: URLAmthanh = "pull-org"
        Case 23: URLAmthanh = "score1"
        Case 24: URLAmthanh = "score2"
        Case 25: URLAmthanh = "shoot"
        Case 26: URLAmthanh = "small_gold" 'vang nho
        Case 27: URLAmthanh = "sound 197"
        Case 28: URLAmthanh = "start"
        Case 29: URLAmthanh = "stone"
        End Select
        FrmMain.Wmp.URL = App.Path & "\Data\Sound\" & URLAmthanh & ".mp3"
        FrmMain.Wmp.Controls.play
    End If
End Sub
Public Sub ThongTinvatpham(Index As Integer)
    Dim Thongtin As String
    Select Case Index
    Case 0: Thongtin = "Boom: va65t pha63m du2ng d9e63 no3 va65t ma2 ba5n kho6ng muo61n ke1o le6n"
    Case 1: Thongtin = "Nu7o71c Ta8ng Lu75c Number 2: giu1p nha6n d9o6i su71c ma5nh khi ke1o"
    Case 2: Thongtin = "Nu7o71c Ta8ng Lu75c Number 1: giu1p nha6n bo61n su71c ma5nh khi ke1o"
    Case 3: Thongtin = "Hoa May Ma81n:nha6n d9o6i may ma1n khi mo73 bao may ma81n"
    Case 4: Thongtin = "Da62u nho71t: ta8ng to61c d9o65 cu3a ma1y 5%"
    Case 5: Thongtin = "So63 Tri Thu71c Ve62 D9a1: Ta8ng Ga61p Ba Gia1 Tri5 Khi Ke1o D9a1 Le6n"
    Case 6: Thongtin = "So63 Tri Thu71c Ve62 Xu7o7ng D9o65ng Va65t: Ta8ng Ga61p Na8m Gia1 Tri5 Khi Ke1o Xu7o7ng Le6n"
    Case 7: Thongtin = "So63 Tri Thu71c Ve62 D9a1 Qu1y: Ta8ng Ga61p D9o6i Gia1 Tri5 Khi Ke1o D9a1 Qu1y Le6n"
    Case 8: Thongtin = "So63 Tri Thu71c Ve62 Va2ng: Ta8ng Ga61p D9o6i Gia1 Tri5 Khi Ke1o Va2ng Le6n"
    Case 9: Thongtin = "CAP Vie65t Nam: la2 loa5i ca1p ke1o cao ca61p do Vie65t Nam sa3n xua61t ta8ng 20% to61c d9o65 ke1o"
    Case 10: Thongtin = "CAP Nha65t Ba3n: la2 loa5i ca1p ke1o cao ca61p nha65p kha63u do Nha65t Ba3n sa3n xua61t ta8ng 30% to61c d9o65 ke1o"
    Case 11: Thongtin = "Ma1y Thu7o72ng: Ma1y ke1o ba82ng tay, to61c d9o65 cha65m"
    Case 12: Thongtin = "Ma1y Da62u: Ma1y ke1o cha5y ba82ng da62u ho3a ,ta8ng to61c d9o65 10%"
    Case 13: Thongtin = "Ma1y D9ie65n: Ma1y ke1o cha5y ba82ng d9ie65n ,ta8ng to61c d9o65 20%"
    Case 14: Thongtin = "Ma1y Xa8ng: Ma1y ke1o cha5y ba82ng Xa8ng A95 ,ta8ng to61c d9o65 30%"
    End Select
    FrmMain.lblThongTinitem.Caption = Thongtin
End Sub
Public Sub MuaVatPham(Index As Integer)
    Select Case Index
    Case 0 'boom
        FrmMain.VatDaMua(Index).Caption = FrmMain.VatDaMua(Index).Caption + 1
        FrmMain.SoVatMua(Index).Caption = FrmMain.SoVatMua(Index).Caption + 1
    Case 1 'tang luc 2
        FrmMain.VatDaMua(Index).Caption = FrmMain.VatDaMua(Index).Caption + 1
        FrmMain.SoVatMua(Index).Caption = FrmMain.SoVatMua(Index).Caption + 1
    Case 2 'tangluc 1
        FrmMain.VatDaMua(Index).Caption = FrmMain.VatDaMua(Index).Caption + 1
        FrmMain.SoVatMua(Index).Caption = FrmMain.SoVatMua(Index).Caption + 1
    Case 3: DungHoaMayMan = True: FrmMain.BTMuavatpham(Index).Visible = False
    Case 4: DungDauNhot = True: FrmMain.BTMuavatpham(Index).Visible = False
    Case 5: DungKienThucDa = True: FrmMain.BTMuavatpham(Index).Visible = False
    Case 6: DungKienThucXuong = True: FrmMain.BTMuavatpham(Index).Visible = False
    Case 7: DungKienThucKimCuong = True: FrmMain.BTMuavatpham(Index).Visible = False
    Case 8: DungKienThucVang = True: FrmMain.BTMuavatpham(Index).Visible = False
    Case 9: FrmMain.DayXich.BorderColor = vbWhite 'cap viet
    Case 10: FrmMain.DayXich.BorderColor = vbYellow 'cap nhat
    Case 11: LoaiMay = "MAYTHUONG": sTocDoTang = 0
    Case 12: LoaiMay = "MAYXANG": sTocDoTang = 30
    Case 13: LoaiMay = "MAYDIEN": sTocDoTang = 20
    Case 14: LoaiMay = "MAYDAU": sTocDoTang = 10
    End Select
End Sub
Public Sub ResetShop()
    FrmMain.BTMuavatpham(3).Visible = True
    FrmMain.BTMuavatpham(4).Visible = True
    FrmMain.BTMuavatpham(5).Visible = True
    FrmMain.BTMuavatpham(6).Visible = True
    FrmMain.BTMuavatpham(7).Visible = True
    FrmMain.BTMuavatpham(8).Visible = True
    DungDauNhot = False
    DungKienThucDa = False
    DungKienThucXuong = False
    DungKienThucKimCuong = False
    DungKienThucVang = False
End Sub
Public Sub KiemDinhGiaTri(LayDuoc As Integer)
    Dim TocDoTang As Integer
    TocDoTang = 0
    Select Case FrmMain.gold(LayDuoc).Tag
    Case "16" 'da sieu to
        Amthanh 29
        TocDo_Keo = 1
        iSoTien = 20
        If DungKienThucDa = True Then iSoTien = iSoTien * 3
    Case "15" 'da to
        Amthanh 18
        TocDo_Keo = 2
        iSoTien = 15
        If DungKienThucDa = True Then iSoTien = iSoTien * 3
    Case "14" 'da vua
        Amthanh 20
        TocDo_Keo = 3
        iSoTien = 10
        If DungKienThucDa = True Then iSoTien = iSoTien * 3
    Case "13" 'da nho
        Amthanh 12
        TocDo_Keo = 4
        iSoTien = 5
        If DungKienThucDa = True Then iSoTien = iSoTien * 3
        '====================vang
    Case "12" 'Vang  To
        Amthanh 15
        TocDo_Keo = 3
        iSoTien = 200
        If DungKienThucVang = True Then iSoTien = iSoTien * 2
    Case "11" 'Vang  To
        Amthanh 15
        TocDo_Keo = 4
        iSoTien = 150
        If DungKienThucVang = True Then iSoTien = iSoTien * 2
    Case "10" 'Vang  vua
        Amthanh 15
        TocDo_Keo = 5
        iSoTien = 100
        If DungKienThucVang = True Then iSoTien = iSoTien * 2
    Case "9" 'Vang  nho
        Amthanh 26
        TocDo_Keo = 6
        iSoTien = 50
        If DungKienThucVang = True Then iSoTien = iSoTien * 2
        '----------------------
    Case "7", "5" 'kim cuong
        Amthanh 15
        TocDo_Keo = 7
        iSoTien = 200
        If DungKienThucKimCuong = True Then iSoTien = iSoTien * 2
    Case "8" 'kim cuong do
        Amthanh 15
        TocDo_Keo = 7
        iSoTien = 250
        If DungKienThucKimCuong = True Then iSoTien = iSoTien * 2
    Case "6" 'kim cuong giot nuoc
        Amthanh 15
        TocDo_Keo = 7
        iSoTien = 300
        If DungKienThucKimCuong = True Then iSoTien = iSoTien * 2
    Case "0" 'may man-------------------
        Amthanh 15
        TocDo_Keo = 7
        iSoTien = Int((Rnd() * 500) + 100)
        If DungHoaMayMan = True Then iSoTien = iSoTien * 2
    Case "1", "2", "3", "4" 'vat la xuong ca va mo neo
        Amthanh 12
        TocDo_Keo = 8
        If DungKienThucXuong = True Then
            iSoTien = 40
        Else
            iSoTien = 20
        End If
    End Select
    TocDo_KeoDay = 800 / TocDo_Keo
    If DungDauNhot = True Then TocDoTang = TocDo_KeoDay * 5 / 100
    TocDo_KeoDay = (TocDo_KeoDay - TocDoTang) - (TocDo_KeoDay * sTocDoTang / 100)
    
    If DungNuocNumber1 = True Then
        TocDo_KeoDay = TocDo_KeoDay / 4
        DungNuocNumber1 = False
    End If
    If DungNuocNumber2 = True Then
        TocDo_KeoDay = TocDo_KeoDay / 2
        DungNuocNumber2 = False
    End If
    FrmMain.Timer4.Interval = TocDo_KeoDay
End Sub
Public Sub ChoBoomNo()
    On Error Resume Next
    Dim ATarray() As Byte '
    'Amthanh 13
    FrmMain.imgBoom.Visible = True
    sBoomNo = sBoomNo + 1
    If sBoomNo = 1 Then
        ATarray() = LoadResData(104, "BOOM")
        sBoomNo = 2
    ElseIf sBoomNo = 2 Then
        ATarray() = LoadResData(105, "BOOM")
        sBoomNo = 3
    ElseIf sBoomNo = 3 Then
        ATarray() = LoadResData(106, "BOOM")
        sBoomNo = 4
    ElseIf sBoomNo = 4 Then
        ATarray() = LoadResData(107, "BOOM")
        sBoomNo = 5
    ElseIf sBoomNo = 5 Then
        ATarray() = LoadResData(108, "BOOM")
        sBoomNo = 6
    ElseIf sBoomNo = 6 Then
        ATarray() = LoadResData(109, "BOOM")
        sBoomNo = 7
    ElseIf sBoomNo = 7 Then
        ATarray() = LoadResData(110, "BOOM")
        sBoomNo = 8
    ElseIf sBoomNo = 8 Then
        ATarray() = LoadResData(111, "BOOM")
        sBoomNo = 9
    ElseIf sBoomNo = 9 Then
        ATarray() = LoadResData(112, "BOOM")
        sBoomNo = 10
    ElseIf sBoomNo = 10 Then
        ATarray() = LoadResData(113, "BOOM")
        sBoomNo = 11
    ElseIf sBoomNo = 11 Then
        ATarray() = LoadResData(114, "BOOM")
        sBoomNo = 12
    Else
        ATarray() = LoadResData(115, "BOOM")
        sBoomNo = 0
        FrmMain.Timer1.Interval = 50
        FrmMain.Timer2.Interval = 0
        FrmMain.Timer4.Enabled = 0
        FrmMain.gold(sDangKeo).Left = 2 * FrmMain.VungDatVang.Width
        FrmMain.gold(sDangKeo).Top = 4320
        FrmMain.Timer6.Enabled = False
        FrmMain.imgBoom.Visible = False
        nul = FrmMain.Bat_Dau()
        FrmMain.Timer1.Enabled = True
        FrmMain.Timer2.Enabled = True
    End If
    FrmMain.imgBoom.LoadImageFromStream ATarray
End Sub
Public Sub MayVeBanDau()
    On Error Resume Next
    Dim BTarray() As Byte '
    BTarray() = LoadResData(101, LoaiMay)
    FrmMain.NguoiDaoVang.LoadImageFromStream BTarray
End Sub
