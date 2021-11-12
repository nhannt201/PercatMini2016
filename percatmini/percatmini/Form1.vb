Imports System.IO
Imports System.Net
Imports Google.API.Search
Public Class percat
    Private Declare Function LockWorkStation Lib "user32.dll" () As Long
    Private Declare Function ShutDownDialog Lib "shell32" Alias "#60" (ByVal any As Long)
    Private Sub ThoátToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThoátToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ẨnPercatToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ẨnPercatToolStripMenuItem.Click
        Me.Hide()
    End Sub

    Private Sub percat_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '  wmp.URL = "http://translate.google.com/translate_tts?tl=en&q=" & Label1.Text
        Dim SAPI
        SAPI = CreateObject("SAPI.spvoice")

        SAPI.Speak(Label1.Text)
        Me.Size = New Size(714, 216)
    End Sub

    Private Sub HiệnPercatToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HiệnPercatToolStripMenuItem.Click
        Me.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            hm.Enabled = True
            TextBox1.Text = StrConv(TextBox1.Text, 2)
            If TextBox1.Text = "end" Then
                Me.Close()
            ElseIf TextBox1.Text = "thoát" Then
                Me.Close()
            ElseIf TextBox1.Text = "thoát percat" Then
                Me.Close()
            ElseIf TextBox1.Text = "máy bay" Then
                Me.Close()
            ElseIf TextBox1.Text = "tắt" Then
                Me.Close()
            ElseIf TextBox1.Text = "ẩn" Then
                Me.Hide()
            ElseIf TextBox1.Text = "ẩn percat" Then
                Me.Hide()
            End If
            scann.Items.Clear()
            Dim client As New GwebSearchClient("http://www.google.com.vn")
            Dim results As IList(Of IWebResult) = client.Search(TextBox1.Text, 10)
            For Each result As IWebResult In results

                scann.Items.Add(result.Title)
            Next
            Me.Size = New Size(714, 470)
        Catch
            'MsgBox("Lỗi BT1!")
        End Try
    End Sub
    Public Function getUrlSource(url As String) As String
        Dim Request As HttpWebRequest = HttpWebRequest.Create(url)
        Dim Response As HttpWebResponse = Request.GetResponse()
        Dim reader As StreamReader = New StreamReader(Response.GetResponseStream)
        Dim httpContent As String
        httpContent = reader.ReadToEnd
        Return httpContent
    End Function

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            scann.Items.Clear()
            If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
                hm.Enabled = True
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Dim client As New GwebSearchClient("http://www.google.com.vn")
                Dim results As IList(Of IWebResult) = client.Search(TextBox1.Text, 10)
                For Each result As IWebResult In results

                    scann.Items.Add(result.Title)
                Next

                Me.Size = New Size(714, 470)
            End If
        Catch
            'MsgBox("Có kí tự lạ khi nhập  & lỗi không xác định!")
        End Try
    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick
        TextBox1.Text = ""
    End Sub

   
  

    'The color and the width of the border.
  

    Dim drag As Boolean

    Dim mousex As Integer

    Dim mousey As Integer



    Private Sub percat_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown

        drag = True 'Sets the variable drag to true.

        mousex = Windows.Forms.Cursor.Position.X - Me.Left 'Sets variable mousex

        mousey = Windows.Forms.Cursor.Position.Y - Me.Top 'Sets variable mousey

    End Sub



    Private Sub percat_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove

        'If drag is set to true then move the form accordingly.

        If drag Then

            Me.Top = Windows.Forms.Cursor.Position.Y - mousey

            Me.Left = Windows.Forms.Cursor.Position.X - mousex

        End If

    End Sub

    Private Sub percat_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp

        drag = False 'Sets drag to false, so the form does not move according to the code in MouseMove

    End Sub

 


    
    Private Sub Text1_Click(sender As Object, e As EventArgs) Handles Text1.Click
       
    End Sub

    Private Sub Text1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text1.KeyPress
        Try
            If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
                hm.Enabled = True
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Text1.Text = ""

            End If
        Catch
            MsgBox("Có lỗi không xác định xảy ra!")
        End Try

    End Sub

    Private Sub check_CheckedChanged(sender As Object, e As EventArgs) Handles check.CheckedChanged
        Try
            If check.Checked = True Then
                My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue(Application.ProductName, Application.ExecutablePath)
                'Sets the location of the form w.r.t the position
            ElseIf check.Checked = False Then
                My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(Application.ProductName)
                'Sets the location of the form w.r.t the position
            End If

        Catch
            MsgBox("Có lỗi xảy ra không thể thêm khởi động cùng Windows!")
        End Try
    End Sub

    Private Sub en_CheckedChanged(sender As Object, e As EventArgs) Handles en.CheckedChanged
        If en.Checked = True Then
            check.Text = "Run when Windows starts"
            Label1.Text = "Hello !"
            Button1.Text = "&Send"
            sd.Text = "Sound =))"
            sk.Text = "&Speak"
            sk.Enabled = True
            vi.Checked = False
        End If
    End Sub

    Private Sub vi_CheckedChanged(sender As Object, e As EventArgs) Handles vi.CheckedChanged
        If vi.Checked = True Then
            check.Text = "Khởi động cùng Windows"
            Label1.Text = "Xin chào !"
            Button1.Text = "&Gửi"
            '   sd.Text = "Không có âm thanh =X"
            sk.Text = "&Nói"
            ' sk.Enabled = False
            en.Checked = False
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            If Label1.Text = "Sorry, I do not find, I will open google to assist you!" Then
                Dim doi78 As String
                web.Show()
                t2.Text = TextBox1.Text
                doi78 = RemoveSign4VietnameseString(t2.Text)
                web.wb.Navigate("https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & doi78)
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Dim SAPI
                SAPI = CreateObject("SAPI.spvoice")

                SAPI.Speak(Label1.Text)
                Me.Size = New Size(714, 216)
                Timer1.Enabled = False
            ElseIf InStr(Label1.Text, "web") > 0 Then
                Label1.Text = Replace(Label1.Text, "web", "")
                Dim do1i As String
                web.Show()
                do1i = Label1.Text
                web.wb.Navigate(do1i)
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Timer1.Enabled = False
            ElseIf Label1.Text = "error" Then
                Label1.Text = "Xin lỗi, tôi không tìm thấy, tôi sẽ mở google để hỗ trợ bạn"
                Dim doi78 As String
                web.Show()

                wmp.URL = "http://percat.esy.es/php.php?text=" & Label1.Text
                t2.Text = TextBox1.Text
                doi78 = RemoveSign4VietnameseString(t2.Text)

                web.wb.Navigate("https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & doi78)
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Timer1.Enabled = False
            End If
            Timer1.Enabled = False
        Catch
            MsgBox("Lỗi kết nối máy chủ '-tm1'")
        End Try

    End Sub
    Private Shared ReadOnly VietnameseSigns As String() = New String() {"aAeEoOuUiIdDyY", "áàạảãâấầậẩẫăắằặẳẵ", "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ", "éèẹẻẽêếềệểễ", "ÉÈẸẺẼÊẾỀỆỂỄ", "óòọỏõôốồộổỗơớờợởỡ", _
                                                                            "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ", "úùụủũưứừựửữ", "ÚÙỤỦŨƯỨỪỰỬỮ", "íìịỉĩ", "ÍÌỊỈĨ", "đ", _
                                                                            "Đ", "ýỳỵỷỹ", "ÝỲỴỶỸ"}

    Public Shared Function RemoveSign4VietnameseString(ByVal str As String) As String
        For i As Integer = 1 To VietnameseSigns.Length - 1
            For j As Integer = 0 To VietnameseSigns(i).Length - 1
                str = str.Replace(VietnameseSigns(i)(j), VietnameseSigns(0)(i - 1))
            Next
        Next
        Return str
    End Function
    Private Sub google_Click(sender As Object, e As EventArgs) Handles google.Click
        web.Show()

        web.wb.Navigate("https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & TextBox1.Text)
    End Sub

    Private Sub yahoo_Click(sender As Object, e As EventArgs) Handles yahoo.Click
        web.Show()

        web.wb.Navigate("https://vn.search.yahoo.com/search?q=" & TextBox1.Text)
    End Sub

    Private Sub bing_Click(sender As Object, e As EventArgs) Handles bing.Click
        web.Show()

        web.wb.Navigate("http://www.bing.com/search?q=" & TextBox1.Text)
    End Sub

    Private Sub sk_Click(sender As Object, e As EventArgs) Handles sk.Click
        

        If sk.Text = "&Nói" Then
            '  TextBox1.Text = TextBox1.Text.Replace("+", " ")

            wmp.URL = "http://percat.esy.es/php.php?text=" & TextBox1.Text
        Else
            Dim SAPI
            SAPI = CreateObject("SAPI.spvoice")

            SAPI.Speak(TextBox1.Text)
            ' wmp.URL = "http://translate.google.com/translate_tts?tl=en&q=" & TextBox1.Text
        End If
    End Sub

    Private Sub hm_Tick(sender As Object, e As EventArgs) Handles hm.Tick
        so.Text = Int(Rnd() * 2)
        If so.Text = "1" Then
            aq.Enabled = True
            hm.Enabled = False
        ElseIf so.Text = "2" Then
            aw.Enabled = True
            hm.Enabled = False
        End If
        hm.Enabled = False
    End Sub

    Private Sub aq_Tick(sender As Object, e As EventArgs) Handles aq.Tick
        Try
            If en.Checked = True Then
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Timer1.Enabled = True
                Label1.Text = getUrlSource("http://percat.esy.es/sv/1.php?chat=" & TextBox1.Text)
                ' wmp.URL = "http://translate.google.com/translate_tts?tl=en&q=" & Label1.Text
                Dim SAPI
                SAPI = CreateObject("SAPI.spvoice")

                SAPI.Speak(Label1.Text)
                aq.Enabled = False
            ElseIf vi.Checked = True Then
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Timer1.Enabled = True
                Label1.Text = getUrlSource("http://percat.esy.es/sv/vi.php?chat=" & TextBox1.Text)
                wmp.URL = "http://percat.esy.es/php.php?text=" & Label1.Text
                aq.Enabled = False
            Else
                Label1.Text = "Xin lỗi, tôi không tìm thấy, tôi sẽ mở google để hỗ trợ bạn!"
                wmp.URL = "http://percat.esy.es/php.php?text=" & Label1.Text
                web.Show()

                web.wb.Navigate("https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & TextBox1.Text)
                aq.Enabled = False
            End If
        Catch
            MsgBox("Đã xảy ra lỗi khi kết nối máy chủ! '-aq'", vbInformation, "Báo lỗi")
        End Try
    End Sub

    Private Sub aw_Tick(sender As Object, e As EventArgs) Handles aw.Tick
        Try
            If en.Checked = True Then
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Timer1.Enabled = True
                Label1.Text = getUrlSource("http://percat.esy.es/sv/2.php?chat=" & TextBox1.Text)
                '    wmp.URL = "http://translate.google.com/translate_tts?tl=en&q=" & Label1.Text
                Dim SAPI
                SAPI = CreateObject("SAPI.spvoice")

                SAPI.Speak(Label1.Text)
                aw.Enabled = False
            ElseIf vi.Checked = True Then
                TextBox1.Text = StrConv(TextBox1.Text, 2)
                Timer1.Enabled = True
                Label1.Text = getUrlSource("http://percat.esy.es/sv/vi.php?chat=" & TextBox1.Text)
                wmp.URL = "http://percat.esy.es/php.php?text=" & Label1.Text
                aw.Enabled = False
            Else
                Label1.Text = "Xin lỗi, tôi không tìm thấy, tôi sẽ mở google để hỗ trợ bạn!"
                wmp.URL = "http://percat.esy.es/php.php?text=" & Label1.Text
                web.Show()

                web.wb.Navigate("https://www.google.com.vn/search?ie=UTF-8&hl=vi&q=" & TextBox1.Text)
                aq.Enabled = False
            End If
        Catch
            MsgBox("Đã xảy ra lỗi khi kết nối máy chủ! '-aw' ", vbInformation, "Báo lỗi")
        End Try
    End Sub

    Private Sub Text1_TextChanged(sender As Object, e As EventArgs) Handles Text1.TextChanged
        If Text1.Text = "tắt máy" Then
            Shell("Shutdown -s")
        ElseIf Text1.Text = "shutdown" Then
            Shell("Shutdown -s")
        ElseIf Text1.Text = "khởi động lại" Then
            Shell("Shutdown -r")
        ElseIf Text1.Text = "restart" Then
            Shell("Shutdown -r")
        ElseIf Text1.Text = "log off" Then
            Shell("Shutdown -l")
        ElseIf Text1.Text = "đăng xuất" Then
            Shell("Shutdown -l")
        End If
    End Sub

    Private Sub hidescan_Click(sender As Object, e As EventArgs) Handles hidescan.Click
        Me.Size = New Size(714, 216)
    End Sub

    Private Sub scann_SelectedIndexChanged(sender As Object, e As EventArgs) Handles scann.SelectedIndexChanged
        Try
            Dim client As New GwebSearchClient("http://www.google.com.vn")
            Dim results As IList(Of IWebResult) = client.Search(scann.Text, 1)
            For Each result As IWebResult In results
                web.wb.Navigate(result.Url)

                web.Show()
            Next
        Catch
            MsgBox("Xảy ra lỗi không xác định '-scann' ", vbInformation, "Báo lỗi")
        End Try
    End Sub

 
    Private Sub TácGiảToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TácGiảToolStripMenuItem.Click
        MsgBox("Tác giả: Nguyễn Trung Nhẫn", vbInformation, "Thông Tin")
    End Sub

    Private Sub LiênHệToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LiênHệToolStripMenuItem.Click
        MsgBox("Địa chỉ liên hệ Email : trungnhan21.12@gmail.com", vbInformation, "Thông Tin")
    End Sub

    Private Sub PhiênBảnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PhiênBảnToolStripMenuItem.Click
        MsgBox("Phiên bản Mini 2.1", vbInformation, "Thông Tin")
    End Sub

    Private Sub ĐịaChỉToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ĐịaChỉToolStripMenuItem.Click
        MsgBox("Địa chỉ : Trường THCS - THPT Nguyễn Văn Khải", vbInformation, "Thông Tin")
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
