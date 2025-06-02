<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTestFace
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnFormat = New System.Windows.Forms.Button()
        Me.btnVideo = New System.Windows.Forms.Button()
        Me.btnTakePicture = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(347, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "Last Picture"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Live Feed"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(350, 26)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(320, 240)
        Me.PictureBox1.TabIndex = 42
        Me.PictureBox1.TabStop = False
        '
        'btnFormat
        '
        Me.btnFormat.Location = New System.Drawing.Point(92, 277)
        Me.btnFormat.Name = "btnFormat"
        Me.btnFormat.Size = New System.Drawing.Size(64, 24)
        Me.btnFormat.TabIndex = 47
        Me.btnFormat.Text = "Format"
        '
        'btnVideo
        '
        Me.btnVideo.Location = New System.Drawing.Point(12, 277)
        Me.btnVideo.Name = "btnVideo"
        Me.btnVideo.Size = New System.Drawing.Size(72, 24)
        Me.btnVideo.TabIndex = 46
        Me.btnVideo.Text = "Video"
        '
        'btnTakePicture
        '
        Me.btnTakePicture.Location = New System.Drawing.Point(224, 277)
        Me.btnTakePicture.Name = "btnTakePicture"
        Me.btnTakePicture.Size = New System.Drawing.Size(108, 24)
        Me.btnTakePicture.TabIndex = 45
        Me.btnTakePicture.Text = "Capture/Start Timer"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(573, 277)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 24)
        Me.Button3.TabIndex = 44
        Me.Button3.Text = "Save"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(471, 277)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 24)
        Me.Button2.TabIndex = 43
        Me.Button2.Text = "Clear"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(490, 311)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(74, 17)
        Me.CheckBox1.TabIndex = 49
        Me.CheckBox1.Text = "Use Timer"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Enabled = False
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"1 second", "3 seconds", "5 seconds"})
        Me.ComboBox1.Location = New System.Drawing.Point(580, 307)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(90, 21)
        Me.ComboBox1.TabIndex = 48
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Location = New System.Drawing.Point(12, 26)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(320, 240)
        Me.PictureBox2.TabIndex = 50
        Me.PictureBox2.TabStop = False
        '
        'frmTestFace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(684, 335)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.btnFormat)
        Me.Controls.Add(Me.btnVideo)
        Me.Controls.Add(Me.btnTakePicture)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.MaximizeBox = False
        Me.Name = "frmTestFace"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Input Test Face"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnFormat As System.Windows.Forms.Button
    Friend WithEvents btnVideo As System.Windows.Forms.Button
    Friend WithEvents btnTakePicture As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents PictureBox2 As PictureBox
End Class
