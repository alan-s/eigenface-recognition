''''''''''''''''''''''''''''''''' 
'     Eigenface Recognition     '
'        Alan Suleiman          '
'          April 2007           '
'     King's College London     '        
'   alan.suleiman@kcl.ac.uk     '
'''''''''''''''''''''''''''''''''

Imports System.Drawing.Drawing2D

Public Class frmHistogram
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents picHisto2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents picHisto As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.picHisto = New System.Windows.Forms.PictureBox
        Me.picHisto2 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        CType(Me.picHisto, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picHisto2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picHisto
        '
        Me.picHisto.BackColor = System.Drawing.Color.White
        Me.picHisto.Location = New System.Drawing.Point(54, 25)
        Me.picHisto.Name = "picHisto"
        Me.picHisto.Size = New System.Drawing.Size(392, 192)
        Me.picHisto.TabIndex = 0
        Me.picHisto.TabStop = False
        '
        'picHisto2
        '
        Me.picHisto2.BackColor = System.Drawing.Color.White
        Me.picHisto2.Location = New System.Drawing.Point(54, 281)
        Me.picHisto2.Name = "picHisto2"
        Me.picHisto2.Size = New System.Drawing.Size(392, 192)
        Me.picHisto2.TabIndex = 1
        Me.picHisto2.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(54, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Before Equalisation"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(54, 265)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "After Equalisation"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(54, 219)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(14, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "0"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(427, 219)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(28, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "255"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(118, 219)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(21, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "50"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(190, 219)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(28, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "100"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(269, 219)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(28, 13)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "150"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(348, 219)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(28, 13)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "200"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(9, 201)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(14, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "0"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(9, 113)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(14, 13)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "0"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(9, 69)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(14, 13)
        Me.Label12.TabIndex = 13
        Me.Label12.Text = "0"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(9, 157)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(14, 13)
        Me.Label10.TabIndex = 14
        Me.Label10.Text = "0"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(9, 25)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(14, 13)
        Me.Label13.TabIndex = 15
        Me.Label13.Text = "0"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(348, 476)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(28, 13)
        Me.Label14.TabIndex = 21
        Me.Label14.Text = "200"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(269, 476)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(28, 13)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "150"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(190, 476)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(28, 13)
        Me.Label16.TabIndex = 19
        Me.Label16.Text = "100"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(118, 476)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(21, 13)
        Me.Label17.TabIndex = 18
        Me.Label17.Text = "50"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(427, 476)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(28, 13)
        Me.Label18.TabIndex = 17
        Me.Label18.Text = "255"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(54, 476)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(14, 13)
        Me.Label19.TabIndex = 16
        Me.Label19.Text = "0"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(9, 281)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(14, 13)
        Me.Label20.TabIndex = 26
        Me.Label20.Text = "0"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(9, 413)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(14, 13)
        Me.Label21.TabIndex = 25
        Me.Label21.Text = "0"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(9, 325)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(14, 13)
        Me.Label22.TabIndex = 24
        Me.Label22.Text = "0"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(9, 369)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(14, 13)
        Me.Label23.TabIndex = 23
        Me.Label23.Text = "0"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(9, 457)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(14, 13)
        Me.Label24.TabIndex = 22
        Me.Label24.Text = "0"
        '
        'PictureBox6
        '
        Me.PictureBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox6.Location = New System.Drawing.Point(471, 350)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(83, 120)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox6.TabIndex = 49
        Me.PictureBox6.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox5.Location = New System.Drawing.Point(471, 97)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(83, 120)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 48
        Me.PictureBox5.TabStop = False
        '
        'frmHistogram
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(570, 502)
        Me.Controls.Add(Me.PictureBox6)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.picHisto2)
        Me.Controls.Add(Me.picHisto)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(578, 529)
        Me.MinimumSize = New System.Drawing.Size(578, 529)
        Me.Name = "frmHistogram"
        Me.Text = "Histogram Equalisation"
        CType(Me.picHisto, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picHisto2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    'Declarations
    Dim max As Integer
    Dim index As Integer = 0

    Dim freqBeforeIn As List(Of Integer)
    Dim freqAfterIn As List(Of Integer)

    Private m_DataValues(255) As Single
    Private m_DataValues2(255) As Single

    'Get original and new frequency arrays when form loads
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        freqBeforeIn = ImageProcessor.getFrequencyBefore()
        freqAfterIn = ImageProcessor.getFrequencyAfter()

        max = m_DataValues(0)
        For i As Integer = 0 To 255
            m_DataValues(i) = freqBeforeIn.Item(i)
            m_DataValues2(i) = freqAfterIn.Item(i)

            If m_DataValues(i) > max Then
                max = m_DataValues(i)
            End If
        Next
        max = max + 100


    End Sub

    'Draw histogram of original image
    Private Sub picHisto_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picHisto.Paint
        e.Graphics.Clear(picHisto.BackColor)

        Dim value As Integer

        value = max * 0.25
        Label10.Text = value
        Label21.Text = value

        value = max * 0.5
        Label11.Text = value
        Label23.Text = value

        value = max * 0.75
        Label12.Text = value
        Label22.Text = value

        Label13.Text = max
        Label20.Text = max

        ' Flip vertically and scale to fit.
        Dim scalex As Single = picHisto.Width / (m_DataValues.GetUpperBound(0) + 1)
        Dim scaley As Single = -picHisto.Height / max
        e.Graphics.ScaleTransform( _
            scalex, scaley, MatrixOrder.Append)

        ' Translate so (0, MAX_VALUE) maps to the origin.
        e.Graphics.TranslateTransform(0, picHisto.Height, MatrixOrder.Append)

        ' Draw the histogram.
        For i As Integer = 0 To m_DataValues.GetUpperBound(0)
            Dim rect As New Rectangle(i, 0, 1, m_DataValues(i))
            Dim clr As Color = Color.FromArgb(&HFF000000 Or QBColor(i Mod 16))
            Dim the_brush As New SolidBrush(Color.Firebrick)
            Dim the_pen As New Pen(Color.Black, 0)
            e.Graphics.FillRectangle(the_brush, rect)
            e.Graphics.DrawRectangle(the_pen, rect)
            the_brush.Dispose()
            the_pen.Dispose()
        Next i

        e.Graphics.ResetTransform()
        e.Graphics.DrawRectangle(Pens.Black, 0, 0, picHisto.Width - 1, picHisto.Height - 1)
    End Sub

    'Draw histogram of equalised image
    Private Sub picHisto2_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picHisto2.Paint
        e.Graphics.Clear(picHisto.BackColor)

        ' Flip vertically and scale to fit.
        Dim scalex As Single = picHisto2.Width / (m_DataValues.GetUpperBound(0) + 1)
        Dim scaley As Single = -picHisto2.Height / max
        e.Graphics.ScaleTransform( _
            scalex, scaley, MatrixOrder.Append)

        ' Translate so (0, MAX_VALUE) maps to the origin.
        e.Graphics.TranslateTransform(0, picHisto.Height, MatrixOrder.Append)

        ' Draw the histogram.
        For i As Integer = 0 To m_DataValues2.GetUpperBound(0)
            Dim rect As New Rectangle(i, 0, 1, m_DataValues2(i))
            Dim clr As Color = Color.FromArgb(&HFF000000 Or QBColor(i Mod 16))
            Dim the_brush As New SolidBrush(Color.Firebrick)
            Dim the_pen As New Pen(Color.Black, 0)
            e.Graphics.FillRectangle(the_brush, rect)
            e.Graphics.DrawRectangle(the_pen, rect)
            the_brush.Dispose()
            the_pen.Dispose()
        Next i

        e.Graphics.ResetTransform()
        e.Graphics.DrawRectangle(Pens.Black, 0, 0, picHisto.Width - 1, picHisto.Height - 1)
    End Sub

    'Upon mouse click, return frequency displayed by specific bar(original chart)
    Private Sub picHisto_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picHisto.MouseDown
        ' Determine which data value was clicked.
        Dim bar_wid As Single = picHisto.Width / (m_DataValues.GetUpperBound(0) + 1)
        Dim i As Integer = Int(e.X / bar_wid)
        MessageBox.Show(i & ": " & m_DataValues(i), "Value", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    'Upon mouse click, return frequency displayed by specific bar(equalised chart)
    Private Sub picHisto2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picHisto2.MouseDown
        ' Determine which data value was clicked.
        Dim bar_wid As Single = picHisto.Width / (m_DataValues2.GetUpperBound(0) + 1)
        Dim i As Integer = Int(e.X / bar_wid)
        MessageBox.Show(i & ": " & m_DataValues2(i), "Value", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class