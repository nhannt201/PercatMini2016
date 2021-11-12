<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class load
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(load))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.so = New System.Windows.Forms.Label()
        Me.n1 = New System.Windows.Forms.Timer(Me.components)
        Me.ten = New System.Windows.Forms.TextBox()
        Me.ok = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'PictureBox1
        '
        resources.ApplyResources(Me.PictureBox1, "PictureBox1")
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.TabStop = False
        '
        'so
        '
        resources.ApplyResources(Me.so, "so")
        Me.so.Name = "so"
        '
        'n1
        '
        Me.n1.Interval = 2000
        '
        'ten
        '
        resources.ApplyResources(Me.ten, "ten")
        Me.ten.Name = "ten"
        '
        'ok
        '
        resources.ApplyResources(Me.ok, "ok")
        Me.ok.Name = "ok"
        Me.ok.UseVisualStyleBackColor = True
        '
        'load
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.Controls.Add(Me.ok)
        Me.Controls.Add(Me.ten)
        Me.Controls.Add(Me.so)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "load"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents so As System.Windows.Forms.Label
    Friend WithEvents n1 As System.Windows.Forms.Timer
    Friend WithEvents ten As System.Windows.Forms.TextBox
    Friend WithEvents ok As System.Windows.Forms.Button
End Class
