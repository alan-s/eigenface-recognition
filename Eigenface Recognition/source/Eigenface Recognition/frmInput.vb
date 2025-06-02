''''''''''''''''''''''''''''''''' 
'     Eigenface Recognition     '
'        Alan Suleiman          '
'          April 2007           '
'     King's College London     '        
'   alan.suleiman@kcl.ac.uk     '
'''''''''''''''''''''''''''''''''

Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.IO

Public Class frmInput
    Inherits System.Windows.Forms.Form

    Dim Counter As Integer = 1
    Dim saved As Boolean = False
    Dim everBeenSaved As Boolean = False

    Shared Now As DateTime
    Shared appPath As String = Application.StartupPath()
    Shared fileName As String
    Shared userName As String

    Private videoDevices As FilterInfoCollection
    Private videoSource As VideoCaptureDevice

    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents PictureBox7 As PictureBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox


#Region "Windows form designer generated code"
    Public Sub New()
        MyBase.New()

        'Initialize
        InitializeComponent()

    End Sub

    'Dispose/cleanup method
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnFormat As System.Windows.Forms.Button
    Friend WithEvents btnVideo As System.Windows.Forms.Button
    Friend WithEvents btnTakePicture As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Private WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Private WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Private WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Private WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Private WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.btnFormat = New System.Windows.Forms.Button()
        Me.btnVideo = New System.Windows.Forms.Button()
        Me.btnTakePicture = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(471, 277)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 24)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Clear"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(573, 277)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 24)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "Save"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "BMP"
        Me.SaveFileDialog1.Filter = "BMP(*.BMP)|*.BMP"
        Me.SaveFileDialog1.Title = "Save Picture"
        '
        'btnFormat
        '
        Me.btnFormat.Location = New System.Drawing.Point(92, 277)
        Me.btnFormat.Name = "btnFormat"
        Me.btnFormat.Size = New System.Drawing.Size(64, 24)
        Me.btnFormat.TabIndex = 2
        Me.btnFormat.Text = "Format"
        '
        'btnVideo
        '
        Me.btnVideo.Location = New System.Drawing.Point(12, 277)
        Me.btnVideo.Name = "btnVideo"
        Me.btnVideo.Size = New System.Drawing.Size(72, 24)
        Me.btnVideo.TabIndex = 1
        Me.btnVideo.Text = "Video"
        '
        'btnTakePicture
        '
        Me.btnTakePicture.Location = New System.Drawing.Point(224, 277)
        Me.btnTakePicture.Name = "btnTakePicture"
        Me.btnTakePicture.Size = New System.Drawing.Size(108, 24)
        Me.btnTakePicture.TabIndex = 3
        Me.btnTakePicture.Text = "Capture/Start Timer"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(350, 26)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(320, 240)
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox3.Location = New System.Drawing.Point(145, 344)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(124, 180)
        Me.PictureBox3.TabIndex = 30
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Location = New System.Drawing.Point(12, 344)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(124, 180)
        Me.PictureBox2.TabIndex = 31
        Me.PictureBox2.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox4.Location = New System.Drawing.Point(278, 344)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(124, 180)
        Me.PictureBox4.TabIndex = 32
        Me.PictureBox4.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox5.Location = New System.Drawing.Point(411, 344)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(124, 180)
        Me.PictureBox5.TabIndex = 33
        Me.PictureBox5.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox6.Location = New System.Drawing.Point(544, 344)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(124, 180)
        Me.PictureBox6.TabIndex = 34
        Me.PictureBox6.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(61, 314)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(173, 20)
        Me.TextBox1.TabIndex = 6
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(12, 317)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(43, 13)
        Me.label1.TabIndex = 36
        Me.label1.Text = "Name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Live Feed"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(347, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Last Picture"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Enabled = False
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"1 second", "3 seconds", "5 seconds"})
        Me.ComboBox1.Location = New System.Drawing.Point(580, 309)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(90, 21)
        Me.ComboBox1.TabIndex = 39
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(490, 313)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(74, 17)
        Me.CheckBox1.TabIndex = 40
        Me.CheckBox1.Text = "Use Timer"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'PictureBox7
        '
        Me.PictureBox7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox7.Location = New System.Drawing.Point(12, 26)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(320, 240)
        Me.PictureBox7.TabIndex = 41
        Me.PictureBox7.TabStop = False
        '
        'frmInput
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(684, 531)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.PictureBox6)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnFormat)
        Me.Controls.Add(Me.btnVideo)
        Me.Controls.Add(Me.btnTakePicture)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmInput"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Input Face Database"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    'Set's focus for text box for name input
    Private Sub frmInput_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        TextBox1.Focus()
    End Sub

    'Show and set frmOverlay to display red face box upon frmInput opening
    Private Sub frmInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        frmOverlay.Location = New Point(Me.Location.X + 16, Me.Location.Y + 49)
        frmOverlay.Show()
        ComboBox1.SelectedIndex = 0

        videoDevices = New FilterInfoCollection(FilterCategory.VideoInputDevice)

        If videoDevices.Count = 0 Then
            MessageBox.Show("No video sources found")
            Return
        End If

        videoSource = New VideoCaptureDevice(videoDevices(0).MonikerString)

        Dim videoCapabilities = videoSource.VideoCapabilities
        For Each cap In videoCapabilities
            If cap.FrameSize.Width = 320 AndAlso cap.FrameSize.Height = 240 Then
                videoSource.VideoResolution = cap
                Exit For
            End If
        Next

        AddHandler videoSource.NewFrame, AddressOf videoSource_NewFrame

        videoSource.Start()

    End Sub

    Private Sub videoSource_NewFrame(sender As Object, eventArgs As NewFrameEventArgs)
        Try
            If PictureBox7.IsHandleCreated Then
                Dim bitmap = CType(eventArgs.Frame.Clone(), Bitmap)

                PictureBox7.Invoke(Sub()
                                       If PictureBox7.Image IsNot Nothing Then
                                           PictureBox7.Image.Dispose()
                                       End If
                                       PictureBox7.Image = bitmap
                                   End Sub)
            End If
        Catch ex As ObjectDisposedException
        Catch ex As InvalidOperationException
        End Try
    End Sub

    'Set new location of frmOverlay to display red face box upon frmInput moving
    Private Sub frmInput_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move

        frmOverlay.Location = New Point(Me.Location.X + 16, Me.Location.Y + 49)

    End Sub

    'Close frmOverlay upon frmInput closing
    Private Sub frmInput_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        frmOverlay.Close()

        If everBeenSaved = True Then
            frmMain.Label7.Text = "Face database inputted"
        Else
            frmMain.Label7.Text = ""
        End If

    End Sub

    'Draws red rectangle on 'last image' picturebox
    Private Sub PictureBox1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint

        Dim rect As Rectangle
        rect = New Rectangle(98, 30, 124, 180)
        e.Graphics.DrawRectangle(Pens.Red, rect)

    End Sub

    'Bring-up video source dialog
    Private Sub btnVideo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVideo.Click

        Dim form As New VideoCaptureDeviceForm()

        If form.ShowDialog() = DialogResult.OK Then
            If videoSource IsNot Nothing Then
                videoSource.SignalToStop()
            End If

            videoSource = form.VideoDevice

            Dim videoCapabilities = videoSource.VideoCapabilities
            For Each cap In videoCapabilities
                If cap.FrameSize.Width = 320 AndAlso cap.FrameSize.Height = 240 Then
                    videoSource.VideoResolution = cap
                    Exit For
                End If
            Next


            ' Hook up event and start
            AddHandler videoSource.NewFrame, AddressOf videoSource_NewFrame

            videoSource.Start()
        End If

    End Sub

    'Bring-up video format dialog
    Private Sub btnFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormat.Click
    End Sub

    'Handles change in checkbox status to turn on and off timer 
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If ComboBox1.Enabled = False Then
            ComboBox1.Enabled = True

        Else
            Timer1.Stop()
            ComboBox1.Enabled = False
            Button2.Enabled = True

        End If

    End Sub

    'Handles activation of campture button.  Decides whether single image capture needed or multiple using timer
    Private Sub btnTakePicture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTakePicture.Click

        If CheckBox1.Checked = True Then

            Select Case ComboBox1.SelectedIndex

                Case 0

                    Button2.Enabled = False
                    Timer1.Interval = 1000
                    Timer1.Start()

                Case 1

                    Button2.Enabled = False
                    Timer1.Interval = 3000
                    Timer1.Start()

                Case 2

                    Button2.Enabled = False
                    Timer1.Interval = 5000
                    Timer1.Start()

            End Select

        Else
            capturePic()
        End If

    End Sub

    'Timer function for timed capture of images
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If Counter < 6 Then

            PlayWav(appPath + "\shutter.wav")
            capturePic()

        Else
            Timer1.Stop()
            Button2.Enabled = True
            ComboBox1.Enabled = True
        End If

    End Sub

    'WAV file playing function
    Shared Function PlayWav(ByVal fileFullPath As String) _
    As Boolean

        'return true if successful, false if otherwise
        Dim iRet As Integer = 0

        Try

            iRet = PlaySound(fileFullPath, 0, SND_FILENAME)

        Catch

        End Try

        Return iRet

    End Function

    Private Declare Auto Function PlaySound Lib "winmm.dll" _
  (ByVal lpszSoundName As String, ByVal hModule As Integer,
   ByVal dwFlags As Integer) As Integer

    Private Const SND_FILENAME As Integer = &H20000


    'Capture method. Captures images to pictureboxes and 'last image' picturebox
    Private Function capturePic()

        Picture = Nothing

        If PictureBox7.Image IsNot Nothing Then
            Picture = CType(PictureBox7.Image.Clone(), Bitmap)
            PictureBox1.Image = Picture
            saved = False

            Select Case Counter

                Case Is = 1
                    Picture = ImageProcessor.Crop(Picture, 98, 30, 124, 180)
                    PictureBox2.Image = Picture
                    Counter = Counter + 1
                Case Is = 2
                    Picture = ImageProcessor.Crop(Picture, 98, 30, 124, 180)
                    PictureBox3.Image = Picture
                    Counter = Counter + 1
                Case Is = 3
                    Picture = ImageProcessor.Crop(Picture, 98, 30, 124, 180)
                    PictureBox4.Image = Picture
                    Counter = Counter + 1
                Case Is = 4
                    Picture = ImageProcessor.Crop(Picture, 98, 30, 124, 180)
                    PictureBox5.Image = Picture
                    Counter = Counter + 1
                Case Is = 5
                    Picture = ImageProcessor.Crop(Picture, 98, 30, 124, 180)
                    PictureBox6.Image = Picture
                    Counter = Counter + 1
                Case Is > 5
                    MsgBox("Five Images Already Inputted", MsgBoxStyle.Information, "Eigenface Recognition")

            End Select

            RaiseEvent AfterPhoto(Picture)

        End If

    End Function

    Private Picture As Image

    Event AfterPhoto(ByVal pic As Image)

    'Save method. Only saves if current images have not already been saved and there are images to save. Crops and grayscales images before saving
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Select Case saved

            Case False

                If Not (PictureBox2.Image Is Nothing) And Not (PictureBox3.Image Is Nothing) And Not (PictureBox4.Image Is Nothing) And Not (PictureBox5.Image Is Nothing) And Not (PictureBox6.Image Is Nothing) Then

                    If TextBox1.Text = "" Then
                        MsgBox("A name is required", MsgBoxStyle.Exclamation, "Eigenface Recognition")
                        Exit Sub
                    End If

                    Picture = ImageProcessor.Resize(PictureBox2.Image.Clone, 83, 120)
                    'Picture = ImageProcessor.Greyscale(Picture)
                    Picture.Save(generateFileName(1), System.Drawing.Imaging.ImageFormat.Jpeg)

                    Picture = ImageProcessor.Resize(PictureBox3.Image.Clone, 83, 120)
                    'Picture = ImageProcessor.Greyscale(Picture)
                    Picture.Save(generateFileName(2), System.Drawing.Imaging.ImageFormat.Jpeg)

                    Picture = ImageProcessor.Resize(PictureBox4.Image.Clone, 83, 120)
                    'Picture = ImageProcessor.Greyscale(Picture)
                    Picture.Save(generateFileName(3), System.Drawing.Imaging.ImageFormat.Jpeg)

                    Picture = ImageProcessor.Resize(PictureBox5.Image.Clone, 83, 120)
                    'Picture = ImageProcessor.Greyscale(Picture)
                    Picture.Save(generateFileName(4), System.Drawing.Imaging.ImageFormat.Jpeg)

                    Picture = ImageProcessor.Resize(PictureBox6.Image.Clone, 83, 120)
                    'Picture = ImageProcessor.Greyscale(Picture)
                    Picture.Save(generateFileName(5), System.Drawing.Imaging.ImageFormat.Jpeg)

                    saved = True
                    everBeenSaved = True

                    MsgBox("Images saved to '" + appPath + "'", MsgBoxStyle.Information, "Eigenface Recognition")

                Else

                    MsgBox("Five images are required", MsgBoxStyle.Exclamation, "Eigenface Recognition")

                End If

            Case True

                MsgBox("Current images already saved", MsgBoxStyle.Exclamation, "Eigenface Recognition")

        End Select

    End Sub

    Shared Sub EnsureFolderExists(folderPath As String)
        If Not Directory.Exists(folderPath) Then
            Directory.CreateDirectory(folderPath)
        End If
    End Sub


    'Generates unique file names for image saving based on user inputted text and system date and time
    Shared Function generateFileName(ByVal i As Integer) As String

        Now = DateTime.Now
        userName = frmInput.TextBox1.Text

        EnsureFolderExists(appPath + "\face DB\")

        fileName = appPath + "\face DB\" + userName + "_" + i.ToString + "_" + Now.Day.ToString + "_" + Now.Month.ToString + "_" + Now.Year.ToString + "(" + Now.Hour.ToString + "." + Now.Minute.ToString + "." + Now.Second.ToString + ")" + ".jpg"

        Return fileName

    End Function

    'Clear method. Clears all pictureboxes and resets variables
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        PictureBox1.Image = Nothing

        PictureBox2.Image = Nothing
        PictureBox3.Image = Nothing
        PictureBox4.Image = Nothing
        PictureBox5.Image = Nothing
        PictureBox6.Image = Nothing

        TextBox1.Text = ""
        Counter = 1
        saved = False

    End Sub

    Private Sub frmInput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If videoSource IsNot Nothing Then
            videoSource.SignalToStop()
        End If
    End Sub
End Class
