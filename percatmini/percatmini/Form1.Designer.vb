<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class percat
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(percat))
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.tray = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Text1 = New System.Windows.Forms.ToolStripTextBox()
        Me.hidescan = New System.Windows.Forms.ToolStripMenuItem()
        Me.HiệnPercatToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ẨnPercatToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ThongTinMN = New System.Windows.Forms.ToolStripMenuItem()
        Me.TácGiảToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LiênHệToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PhiênBảnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ĐịaChỉToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ThoátToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.check = New System.Windows.Forms.CheckBox()
        Me.vi = New System.Windows.Forms.CheckBox()
        Me.en = New System.Windows.Forms.CheckBox()
        Me.sd = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.t2 = New System.Windows.Forms.TextBox()
        Me.google = New System.Windows.Forms.Label()
        Me.yahoo = New System.Windows.Forms.Label()
        Me.bing = New System.Windows.Forms.Label()
        Me.sk = New System.Windows.Forms.Button()
        Me.aq = New System.Windows.Forms.Timer(Me.components)
        Me.aw = New System.Windows.Forms.Timer(Me.components)
        Me.so = New System.Windows.Forms.Label()
        Me.hm = New System.Windows.Forms.Timer(Me.components)
        Me.scann = New System.Windows.Forms.ListBox()
        Me.wmp = New AxWMPLib.AxWindowsMediaPlayer()
        Me.tray.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.wmp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenuStrip = Me.tray
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Percat Mini"
        Me.NotifyIcon1.Visible = True
        '
        'tray
        '
        Me.tray.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Text1, Me.hidescan, Me.HiệnPercatToolStripMenuItem, Me.ẨnPercatToolStripMenuItem, Me.ThongTinMN, Me.ThoátToolStripMenuItem})
        Me.tray.Name = "tray"
        Me.tray.Size = New System.Drawing.Size(161, 139)
        '
        'Text1
        '
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(100, 23)
        '
        'hidescan
        '
        Me.hidescan.Name = "hidescan"
        Me.hidescan.Size = New System.Drawing.Size(160, 22)
        Me.hidescan.Text = "Ẩn Tìm Kiếm"
        '
        'HiệnPercatToolStripMenuItem
        '
        Me.HiệnPercatToolStripMenuItem.Name = "HiệnPercatToolStripMenuItem"
        Me.HiệnPercatToolStripMenuItem.ShortcutKeyDisplayString = "P"
        Me.HiệnPercatToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.HiệnPercatToolStripMenuItem.Text = "Hiện Percat"
        '
        'ẨnPercatToolStripMenuItem
        '
        Me.ẨnPercatToolStripMenuItem.Name = "ẨnPercatToolStripMenuItem"
        Me.ẨnPercatToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.ẨnPercatToolStripMenuItem.Text = "Ẩn Percat"
        '
        'ThongTinMN
        '
        Me.ThongTinMN.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TácGiảToolStripMenuItem, Me.LiênHệToolStripMenuItem, Me.PhiênBảnToolStripMenuItem, Me.ĐịaChỉToolStripMenuItem})
        Me.ThongTinMN.Name = "ThongTinMN"
        Me.ThongTinMN.Size = New System.Drawing.Size(160, 22)
        Me.ThongTinMN.Text = "Thông Tin"
        '
        'TácGiảToolStripMenuItem
        '
        Me.TácGiảToolStripMenuItem.Name = "TácGiảToolStripMenuItem"
        Me.TácGiảToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.TácGiảToolStripMenuItem.Text = "Tác giả"
        '
        'LiênHệToolStripMenuItem
        '
        Me.LiênHệToolStripMenuItem.Name = "LiênHệToolStripMenuItem"
        Me.LiênHệToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.LiênHệToolStripMenuItem.Text = "Liên hệ"
        '
        'PhiênBảnToolStripMenuItem
        '
        Me.PhiênBảnToolStripMenuItem.Name = "PhiênBảnToolStripMenuItem"
        Me.PhiênBảnToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.PhiênBảnToolStripMenuItem.Text = "Phiên bản"
        '
        'ĐịaChỉToolStripMenuItem
        '
        Me.ĐịaChỉToolStripMenuItem.Name = "ĐịaChỉToolStripMenuItem"
        Me.ĐịaChỉToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.ĐịaChỉToolStripMenuItem.Text = "Địa chỉ"
        '
        'ThoátToolStripMenuItem
        '
        Me.ThoátToolStripMenuItem.Name = "ThoátToolStripMenuItem"
        Me.ThoátToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.ThoátToolStripMenuItem.Text = "Thoát"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(628, 129)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(79, 46)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&Gửi"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 9)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(224, 181)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(257, 9)
        Me.Label1.MinimumSize = New System.Drawing.Size(450, 110)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(450, 119)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Xin chào !"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(257, 131)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(365, 44)
        Me.TextBox1.TabIndex = 4
        '
        'check
        '
        Me.check.AutoSize = True
        Me.check.Location = New System.Drawing.Point(295, 181)
        Me.check.Name = "check"
        Me.check.Size = New System.Drawing.Size(149, 17)
        Me.check.TabIndex = 5
        Me.check.Text = "Khởi động cùng Windows"
        Me.check.UseVisualStyleBackColor = True
        '
        'vi
        '
        Me.vi.AutoSize = True
        Me.vi.Location = New System.Drawing.Point(12, 196)
        Me.vi.Name = "vi"
        Me.vi.Size = New System.Drawing.Size(74, 17)
        Me.vi.TabIndex = 7
        Me.vi.Text = "Tiếng Việt"
        Me.vi.UseVisualStyleBackColor = True
        '
        'en
        '
        Me.en.AutoSize = True
        Me.en.Checked = True
        Me.en.CheckState = System.Windows.Forms.CheckState.Checked
        Me.en.Location = New System.Drawing.Point(92, 196)
        Me.en.Name = "en"
        Me.en.Size = New System.Drawing.Size(60, 17)
        Me.en.TabIndex = 8
        Me.en.Text = "English"
        Me.en.UseVisualStyleBackColor = True
        '
        'sd
        '
        Me.sd.AutoSize = True
        Me.sd.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sd.Location = New System.Drawing.Point(169, 195)
        Me.sd.Name = "sd"
        Me.sd.Size = New System.Drawing.Size(67, 16)
        Me.sd.TabIndex = 9
        Me.sd.Text = "Sound =))"
        '
        'Timer1
        '
        Me.Timer1.Interval = 200
        '
        't2
        '
        Me.t2.Location = New System.Drawing.Point(761, 196)
        Me.t2.Name = "t2"
        Me.t2.Size = New System.Drawing.Size(39, 20)
        Me.t2.TabIndex = 10
        '
        'google
        '
        Me.google.AutoSize = True
        Me.google.ForeColor = System.Drawing.Color.Blue
        Me.google.Location = New System.Drawing.Point(482, 188)
        Me.google.Name = "google"
        Me.google.Size = New System.Drawing.Size(41, 13)
        Me.google.TabIndex = 11
        Me.google.Text = "&Google"
        '
        'yahoo
        '
        Me.yahoo.AutoSize = True
        Me.yahoo.ForeColor = System.Drawing.Color.Purple
        Me.yahoo.Location = New System.Drawing.Point(528, 188)
        Me.yahoo.Name = "yahoo"
        Me.yahoo.Size = New System.Drawing.Size(38, 13)
        Me.yahoo.TabIndex = 12
        Me.yahoo.Text = "&Yahoo"
        '
        'bing
        '
        Me.bing.AutoSize = True
        Me.bing.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.bing.Location = New System.Drawing.Point(572, 188)
        Me.bing.Name = "bing"
        Me.bing.Size = New System.Drawing.Size(28, 13)
        Me.bing.TabIndex = 13
        Me.bing.Text = "&Bing"
        '
        'sk
        '
        Me.sk.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sk.Location = New System.Drawing.Point(627, 178)
        Me.sk.Name = "sk"
        Me.sk.Size = New System.Drawing.Size(79, 33)
        Me.sk.TabIndex = 14
        Me.sk.Text = "&Speak"
        Me.sk.UseVisualStyleBackColor = True
        '
        'aq
        '
        '
        'aw
        '
        '
        'so
        '
        Me.so.AutoSize = True
        Me.so.Location = New System.Drawing.Point(738, 100)
        Me.so.Name = "so"
        Me.so.Size = New System.Drawing.Size(13, 13)
        Me.so.TabIndex = 15
        Me.so.Text = "0"
        '
        'hm
        '
        '
        'scann
        '
        Me.scann.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.scann.FormattingEnabled = True
        Me.scann.ItemHeight = 18
        Me.scann.Location = New System.Drawing.Point(12, 219)
        Me.scann.Name = "scann"
        Me.scann.Size = New System.Drawing.Size(690, 238)
        Me.scann.TabIndex = 16
        '
        'wmp
        '
        Me.wmp.Enabled = True
        Me.wmp.Location = New System.Drawing.Point(738, 188)
        Me.wmp.Name = "wmp"
        Me.wmp.OcxState = CType(resources.GetObject("wmp.OcxState"), System.Windows.Forms.AxHost.State)
        Me.wmp.Size = New System.Drawing.Size(10, 10)
        Me.wmp.TabIndex = 6
        '
        'percat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(714, 470)
        Me.ContextMenuStrip = Me.tray
        Me.Controls.Add(Me.scann)
        Me.Controls.Add(Me.so)
        Me.Controls.Add(Me.sk)
        Me.Controls.Add(Me.bing)
        Me.Controls.Add(Me.yahoo)
        Me.Controls.Add(Me.google)
        Me.Controls.Add(Me.t2)
        Me.Controls.Add(Me.sd)
        Me.Controls.Add(Me.en)
        Me.Controls.Add(Me.vi)
        Me.Controls.Add(Me.wmp)
        Me.Controls.Add(Me.check)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "percat"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Percat Mini 2.2"
        Me.tray.ResumeLayout(False)
        Me.tray.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.wmp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents tray As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents HiệnPercatToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ẨnPercatToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ThoátToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Text1 As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents check As System.Windows.Forms.CheckBox
    Friend WithEvents vi As System.Windows.Forms.CheckBox
    Friend WithEvents en As System.Windows.Forms.CheckBox
    Friend WithEvents sd As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents t2 As System.Windows.Forms.TextBox
    Friend WithEvents google As System.Windows.Forms.Label
    Friend WithEvents yahoo As System.Windows.Forms.Label
    Friend WithEvents bing As System.Windows.Forms.Label
    Friend WithEvents sk As System.Windows.Forms.Button
    Friend WithEvents aq As System.Windows.Forms.Timer
    Friend WithEvents aw As System.Windows.Forms.Timer
    Friend WithEvents so As System.Windows.Forms.Label
    Friend WithEvents hm As System.Windows.Forms.Timer
    Friend WithEvents wmp As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents scann As System.Windows.Forms.ListBox
    Friend WithEvents hidescan As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ThongTinMN As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TácGiảToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LiênHệToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PhiênBảnToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ĐịaChỉToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
