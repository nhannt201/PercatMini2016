Imports System.IO
Imports System.Net
Public Class load

    Private Sub load_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim a As String
            Dim u As String

            a = getUrlSource("http://percat.esy.es/ip.php")
            u = getUrlSource("http://percat.esy.es/user/" & a & ".txt")

            If u = "no" Then
                Label1.Text = "Bạn tên là gì ?"
                ten.Visible = True
                ok.Visible = True
                so.Text = "1"
            Else
                so.Text = "1"
                n1.Enabled = True
            End If
        Catch
            Label1.Text = "Bạn tên là gì ?"
            ten.Visible = True
            ok.Visible = True
            so.Text = "1"
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

    Private Sub n1_Tick(sender As Object, e As EventArgs) Handles n1.Tick
        Dim u As String
        Dim a1 As String
        Dim er As String

        er = getUrlSource("http://percat.esy.es/wel.txt")
       

        a1 = getUrlSource("http://percat.esy.es/ip.php")
       
        u = getUrlSource("http://percat.esy.es/user/" & a1 & ".txt")
       

        percat.Label1.Text = "Hello, " & u
        n1.Enabled = False
        percat.Show()
        Me.Hide()
    End Sub

    Private Sub ok_Click(sender As Object, e As EventArgs) Handles ok.Click
        Dim bd As String
        percat.Label1.Text = "Hello, " & ten.Text
        bd = getUrlSource("http://percat.esy.es/get.php?name=" & ten.Text)
        ten.Visible = False
        ok.Visible = False
        percat.Show()
        Me.Hide()
    End Sub
End Class