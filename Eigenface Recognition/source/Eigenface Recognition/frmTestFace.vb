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

Public Class frmTestFace

    Dim saved As Boolean = False
    Dim everBeenSaved As Boolean = False
    Dim counter = 1
    Dim Picture2 As Bitmap
    Shared appPath As String = Application.StartupPath()
    Shared fileName As String

    Private videoDevices As FilterInfoCollection
    Private videoSource As VideoCaptureDevice


    'Upon loading, display and set frmOverlay
    Private Sub frmTestFace_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ComboBox1.SelectedIndex = 0
        frmOverlay.Location = New Point(Me.Location.X + 16, Me.Location.Y + 49)
        frmOverlay.Show()

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
            If PictureBox2.IsHandleCreated Then
                Dim bitmap = CType(eventArgs.Frame.Clone(), Bitmap)

                PictureBox2.Invoke(Sub()

                                       If PictureBox2.Image IsNot Nothing Then
                                           PictureBox2.Image.Dispose()
                                       End If
                                       PictureBox2.Image = bitmap
                                   End Sub)
            End If
        Catch ex As ObjectDisposedException
        Catch ex As InvalidOperationException
        End Try
    End Sub

    'Upon moving, reset location of frmOverlay
    Private Sub frmTestFace_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move

        frmOverlay.Location = New Point(Me.Location.X + 16, Me.Location.Y + 49)

    End Sub

    'Upon closing, close frmOverlay
    Private Sub frmTestFace_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        frmOverlay.Close()

        If everBeenSaved = True Then
            frmMain.Label7.Text = "Test face inputted"
        Else
            frmMain.Label7.Text = ""
        End If

    End Sub

    'Draw red rectangle on 'last image' picturebox
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

            AddHandler videoSource.NewFrame, AddressOf videoSource_NewFrame

            videoSource.Start()
        End If
    End Sub

    'Bring-up video format dialog
    Private Sub btnFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormat.Click

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If counter < 2 Then

            frmInput.PlayWav(appPath + "\shutter.wav")
            capturePic()
            counter = counter + 1

        Else
            Timer1.Stop()
            Button2.Enabled = True
        End If

    End Sub

    'Handles enabling and disabling of checkbox to use timer
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If ComboBox1.Enabled = False Then
            ComboBox1.Enabled = True

        Else
            Timer1.Stop()
            Button2.Enabled = True
            ComboBox1.Enabled = False

        End If

    End Sub

    'Handles what should happen when take picture button is pressed i.e. use timer or snap single photo
    Private Sub btnTakePicture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTakePicture.Click

        If CheckBox1.Checked = True Then

            Select Case ComboBox1.SelectedIndex

                Case 0

                    counter = 1
                    Button2.Enabled = False
                    Timer1.Interval = 1000
                    Timer1.Start()

                Case 1

                    counter = 1
                    Button2.Enabled = False
                    Timer1.Interval = 3000
                    Timer1.Start()

                Case 2

                    counter = 1
                    Button2.Enabled = False
                    Timer1.Interval = 5000
                    Timer1.Start()

            End Select

        Else
            capturePic()
        End If

    End Sub

    Private Picture As Image
    Event AfterPhoto(ByVal pic As Image)

    'Capture method to capture images to picturebox
    Private Function capturePic()

        Picture = Nothing

        If PictureBox2.Image IsNot Nothing Then

            Picture = CType(PictureBox2.Image.Clone(), Bitmap)
            PictureBox1.Image = Picture
            saved = False

            RaiseEvent AfterPhoto(Picture)
        End If

    End Function

    Shared Sub EnsureFolderExists(folderPath As String)
        If Not Directory.Exists(folderPath) Then
            Directory.CreateDirectory(folderPath)
        End If
    End Sub

    'Save method. Only saves if current image has not already been saved and there is an image to save.  Crops and grayscales image before saving.
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If saved = False Then

            If Not (PictureBox1.Image Is Nothing) Then

                Picture = ImageProcessor.Crop(Picture, 98, 30, 124, 180)
                Picture = ImageProcessor.Resize(Picture, 83, 120)
                'Picture = ImageProcessor.Greyscale(Picture)
                Picture2 = Picture
                fileName = appPath + "\face DB\test face\"
                EnsureFolderExists(fileName)


                Picture2.Save(fileName + "testFace.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                frmMain.pictureBox2.Image = Picture.Clone
                Picture.Dispose()
                Picture2.Dispose()

                saved = True
                everBeenSaved = True

                MsgBox("Images saved to '" + fileName + "testFace.jpg'", MsgBoxStyle.Information, "Eigenface Recognition")

            Else

                    MsgBox("An image is required", MsgBoxStyle.Exclamation, "Eigenface Recognition")

            End If
        Else
            MsgBox("Current image already saved", MsgBoxStyle.Exclamation, "Eigenface Recognition")
        End If


    End Sub

    'Clear method
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        PictureBox1.Image = Nothing
        saved = False
        counter = 1
    End Sub

    Private Sub frmTestFace_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If videoSource IsNot Nothing Then
            videoSource.SignalToStop()
        End If
    End Sub
End Class