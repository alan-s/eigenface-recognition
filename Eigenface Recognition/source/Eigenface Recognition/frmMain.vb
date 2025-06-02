''''''''''''''''''''''''''''''''' 
'     Eigenface Recognition     '
'        Alan Suleiman          '
'          April 2007           '
'     King's College London     '        
'   alan.suleiman@kcl.ac.uk     '
'''''''''''''''''''''''''''''''''

Imports System.Windows.Forms
Imports System.IO
Imports Mapack
Imports System.Math
Imports System.Diagnostics

Public Class frmMain

    'Declarations
    Dim di As DirectoryInfo
    Dim fi As FileInfo
    Shared appPath As String = Application.StartupPath()

    Dim targetBitmap As Bitmap
    Dim initialWidth As Integer = 0
    Dim initialHeight As Integer = 0
    Dim heightOfImage As Integer = 0
    Dim widthOfImage As Integer = 0

    Dim dbEqualised = False

    Dim y As Integer = 0
    Dim x As Integer = 0

    Dim i As Integer = 0
    Dim counter As Integer = 0

    Dim trainDir As String
    Dim testFace As String

    Dim noOfImages As Integer
    Dim imgLength As Integer = 0

    Dim imgAsInteger As List(Of Integer)
    Dim bitMapStore As List(Of List(Of Integer))

    Dim averageFace As List(Of Integer)

    Dim value As Integer

    Dim matrix(,) As Double
    Dim m As Integer = 0
    Dim n As Integer = 0

    Dim eigenvectors As ArrayList
    Dim eigenface As List(Of Double)
    Dim eigenfaceStore As List(Of List(Of Double))

    Dim coefficent As List(Of Double)
    Dim coefficentStore As List(Of List(Of Double))

    Dim tmpValue As Double = 0
    Dim finalPixel As Integer = 0

    Dim testFaceAsInteger As List(Of Integer)
    Dim testDiff As List(Of Integer)

    Dim coefficentTest As List(Of Double)
    Dim distance As ArrayList

    Dim smallestPos As Integer = 0

    Dim distanceSort As ArrayList

    Dim fileAttArray()
    Dim fileAtt As System.IO.FileSystemInfo



    'Part of GUI. Draws rectangles around the form
    Private Sub frmMain_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Dim rect As Rectangle
        Dim rect2 As Rectangle
        Dim rect3 As Rectangle
        Dim rect4 As Rectangle
        Dim rect5 As Rectangle
        Dim pen As Pen = New Pen(Color.Black, 1)

        rect = New Rectangle(10, 15, 400, 110)
        e.Graphics.DrawRectangle(pen, rect)

        rect2 = New Rectangle(423, 15, 250, 110)
        e.Graphics.DrawRectangle(pen, rect2)

        rect3 = New Rectangle(10, 135, 663, 218)
        e.Graphics.DrawRectangle(pen, rect3)

        rect4 = New Rectangle(683, 15, 150, 338)
        e.Graphics.DrawRectangle(pen, rect4)

        rect5 = New Rectangle(10, 359, 663, 60)
        e.Graphics.DrawRectangle(pen, rect5)

    End Sub

    'Handles Histogram Equalisation ON, OFF
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            Label7.Text = "Histogram Equalisation ON"
            CheckBox2.Enabled = True
        Else
            Label7.Text = "Histogram Equalisation OFF"
            CheckBox2.Enabled = False
        End If

    End Sub

    'Set default image resize
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 3
        Label7.Text = "Histogram Equalisation ON"
        CheckBox2.Enabled = True
    End Sub

    'Select directory for database
    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        Dim trainingDir As New FolderBrowserDialog
        Label7.Text = "Selecting face database directory..."
        trainingDir.Description = "Face Database:"
        'trainingDir.RootFolder = Environment.SpecialFolder.Desktop ' Start at the desktop   
        trainingDir.SelectedPath = appPath + "\face DB"

        If trainingDir.ShowDialog() = Windows.Forms.DialogResult.OK Then

            trainDir = trainingDir.SelectedPath

            textBox1.Text = trainDir
            di = New DirectoryInfo(trainDir)
            button3.Enabled = True
            Label7.Text = "Face database directory selected"

        Else

            If button3.Enabled = True Then
                Label7.Text = "Face database directory selected"
            Else
                Label7.Text = ""
            End If

        End If

    End Sub

    'Enrol user selected database
    Private Sub button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button3.Click

        button2.Enabled = True
        Button9.Enabled = False
        enrol(trainDir)

    End Sub

    'Select test face file
    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click

        Label7.Text = "Selecting test face..."
        dlgTestFace.InitialDirectory = appPath + "\face DB\test face\"
        dlgTestFace.Filter = "Image files (.jpg, .jpeg, .bmp, .gif)|*.jpg;*.jpeg;*.bmp;*.gif"

        If dlgTestFace.ShowDialog() = Windows.Forms.DialogResult.OK Then

            testFace = dlgTestFace.FileName.ToString
            TextBox2.Text = testFace

            pictureBox2.ImageLocation = testFace
            button4.Enabled = True
            Label7.Text = "Test face selected"

        Else

            If button4.Enabled = True Then
                Label7.Text = "Test face selected"
            Else
                Label7.Text = ""
            End If

        End If

    End Sub

    'Match with test face from file
    Private Sub button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button4.Click

        match(testFace)

    End Sub

    'Inputting face datbase using webcam
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Label7.Text = "Inputting face database..."
        frmInput.Show()

    End Sub

    'Enrol webcam database
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        testFace = ""
        TextBox2.Text = ""
        button2.Enabled = False
        button4.Enabled = False
        Button9.Enabled = True
        enrol(appPath + "\face DB\")

    End Sub

    'Inputting test face via webcam
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Label7.Text = "Inputting test face..."
        frmTestFace.Show()

    End Sub

    'Match using test face from webcam
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        fi = New FileInfo(appPath + "\face DB\test face\testFace.jpg")

        If fi.Exists = True Then
            match(appPath + "\face DB\test face\testFace.jpg")
        Else
            MsgBox("No compatible image file(testFace) found in (test face) subdirectory", MsgBoxStyle.Exclamation, "Eigenface Recognition")
        End If

    End Sub

    'Enrol
    Private Function enrol(ByVal dirIn As String)

        Label7.Text = "Enrolling face database..."

        trainDir = dirIn
        di = New DirectoryInfo(trainDir)
        Dim jpgFiles = di.GetFiles("*.jpg")

        imgAsInteger = New List(Of Integer)
        bitMapStore = New List(Of List(Of Integer))

        averageFace = New List(Of Integer)

        'Check for compatable files and limit number of images enrolled
        If jpgFiles.Length = 0 Then

            MsgBox("No compatible image files found in directory", MsgBoxStyle.Exclamation, "Eigenface Recognition")
            Label7.Text = "Enrollment failed"
            Exit Function

        ElseIf jpgFiles.Length > 1000 Then
            MsgBox("Maximum number of images allowed in face database is 1000", MsgBoxStyle.Exclamation, "Eigenface Recognition")
            Exit Function
        End If

        'Checking to make sure all input images are initially the same size
        i = 0
        For Each fi In jpgFiles

            targetBitmap = New Bitmap(fi.FullName)

            If i = 0 Then

                i = i + 1
                widthOfImage = targetBitmap.Width
                heightOfImage = targetBitmap.Height

            End If

            If Not targetBitmap.Width.Equals(widthOfImage) Or Not targetBitmap.Height.Equals(heightOfImage) Then

                MsgBox("Width and Height of images must all be the same", MsgBoxStyle.Critical, "Eigenface Recognition")
                Label7.Text = "Enrollment failed"
                Exit Function

            End If

        Next
        initialHeight = targetBitmap.Height
        initialWidth = targetBitmap.Width

        'Checking to see if image needs to be resized and whether already in greyscale
        dbEqualised = False
        For Each fi In di.GetFiles("*.jpg")

            targetBitmap = New Bitmap(fi.FullName)

            Select Case ComboBox1.SelectedIndex
                    Case 0
                        targetBitmap = ImageProcessor.Resize(targetBitmap, 83, 120)
                    Case 1
                        targetBitmap = ImageProcessor.Resize(targetBitmap, 53, 90)
                    Case 2
                        targetBitmap = ImageProcessor.Resize(targetBitmap, 33, 70)
                    Case 3
                        targetBitmap = ImageProcessor.Resize(targetBitmap, 23, 60)
                End Select
            'End If

            Dim px1 = targetBitmap.GetPixel(0, 0)
            Dim px2 = targetBitmap.GetPixel(targetBitmap.Width \ 2, targetBitmap.Height \ 2)

            If Not (px1.R = px1.G And px1.G = px1.B And px2.R = px2.G And px2.G = px2.B) Then
                targetBitmap = ImageProcessor.Greyscale(targetBitmap)
            End If

            'Image into vector
            For y = 0 To targetBitmap.Height - 1
                For x = 0 To targetBitmap.Width - 1
                    imgAsInteger.Add(targetBitmap.GetPixel(x, y).R.ToString())
                Next
            Next

            'Applying histogram equalisation if checkbox ticked
            If CheckBox1.Checked = True Then

                imgAsInteger = ImageProcessor.HistogramEqualisation(imgAsInteger)
                dbEqualised = True

            End If
            bitMapStore.Add(imgAsInteger)
            imgAsInteger = New List(Of Integer)
        Next

        noOfImages = bitMapStore.Count
        Label15.Text = noOfImages
        heightOfImage = targetBitmap.Height
        widthOfImage = targetBitmap.Width
        imgLength = targetBitmap.Height * targetBitmap.Width
        targetBitmap.Dispose()

        'Calculating average face.  Uses each pixel value from each image to do so.
        For counter = 0 To imgLength - 1
            Dim sum As Integer = 0
            For i = 0 To noOfImages - 1
                sum += bitMapStore(i)(counter)
            Next
            averageFace.Add(sum \ noOfImages)
        Next

        'Display average face
        targetBitmap = New Bitmap(widthOfImage, heightOfImage)
        counter = 0

        For y = 0 To heightOfImage - 1

            For x = 0 To widthOfImage - 1
                value = averageFace.Item(counter)
                targetBitmap.SetPixel(x, y, Color.FromArgb(255, value, value, value))
                counter = counter + 1
            Next
        Next
        pictureBox1.Image = targetBitmap.Clone


        'calculating difference of each face from average and inserting into matrix
        Dim matrix(noOfImages - 1, imgLength - 1)

        For m = 0 To noOfImages - 1
            Dim imgVec = bitMapStore(m)
            For n = 0 To imgLength - 1
                matrix(m, n) = imgVec(n) - averageFace(n)
            Next
        Next

        'Calculate Covariance Matrix
        Dim covMatrix As Matrix
        covMatrix = New Matrix(noOfImages, noOfImages)

        value = 0

        For n = 0 To noOfImages - 1
            For m = n To noOfImages - 1
                Dim value As Double = 0
                For i = 0 To imgLength - 1
                    value += matrix(m, i) * matrix(n, i)
                Next
                covMatrix(n, m) = value
                covMatrix(m, n) = value
            Next
        Next

        'Get Eigenvectors
        Dim eigen As Mapack.EigenvalueDecomposition
        eigen = New Mapack.EigenvalueDecomposition(covMatrix)

        'Copy eigenvectors into eigenvectors array
        eigenvectors = New ArrayList

        For m = 0 To noOfImages - 1

            For n = 0 To noOfImages - 1

                eigenvectors.Add(eigen.EigenvectorMatrix.Item(n, m))
            Next
        Next

        'Calculate eigenfaces
        eigenface = New List(Of Double)
        eigenfaceStore = New List(Of List(Of Double))


        For m = 0 To noOfImages - 1
            For n = 0 To imgLength - 1
                i = m
                tmpValue = 0
                While i < eigenvectors.Count
                    tmpValue += matrix(m, n) * eigenvectors.Item(i)
                    i += noOfImages
                End While
                eigenface.Add(tmpValue)
            Next
            eigenfaceStore.Add(eigenface)
            eigenface = New List(Of Double)
        Next

        eigenvectors.Clear()

        'outputting(eigenfaces)
        For counter = eigenfaceStore.Count - 1 To eigenfaceStore.Count - 1

            imgAsInteger = New List(Of Integer)
            eigenface = New List(Of Double)

            eigenface = eigenfaceStore.Item(counter)

            For i = 0 To eigenface.Count - 1
                value = eigenface.Item(i)

                If value < 0 Then

                    value = 0

                ElseIf value > 255 Then

                    value = 255

                End If

                imgAsInteger.Add(value)

            Next

            imgAsInteger = ImageProcessor.HistogramEqualisation(imgAsInteger)

            m = 0
            For y = 0 To heightOfImage - 1

                For x = 0 To widthOfImage - 1
                    value = imgAsInteger.Item(m)
                    targetBitmap.SetPixel(x, y, Color.FromArgb(255, value, value, value))
                    m = m + 1
                Next
            Next
            PictureBox4.Image = targetBitmap.Clone

        Next

        'Calculating coefficents between difference vectors and eigenfaces 
        coefficent = New List(Of Double)
        coefficentStore = New List(Of List(Of Double))

        tmpValue = 0
        For i = 0 To noOfImages - 1
            Dim coeff As New List(Of Double)
            For f = 0 To eigenfaceStore.Count - 1
                Dim projection As Double = 0
                Dim face = eigenfaceStore(f)
                For j = 0 To imgLength - 1
                    projection += matrix(i, j) * face(j)
                Next
                coeff.Add(projection)
            Next
            coefficentStore.Add(coeff)
        Next

        Label7.Text = "Face database enrolled"

    End Function

    'Match
    Private Function match(ByVal dirIn As String)

        testFaceAsInteger = New List(Of Integer)(widthOfImage * heightOfImage)
        testDiff = New List(Of Integer)(widthOfImage * heightOfImage)

        coefficentTest = New List(Of Double)(eigenfaceStore.Count)
        coefficentTest = New List(Of Double)(eigenfaceStore.Count)
        distance = New ArrayList()

        Label7.Text = "Matching Facing..."

        ' Load and validate image
        targetBitmap = New Bitmap(dirIn)
        If targetBitmap.Width <> initialWidth OrElse targetBitmap.Height <> initialHeight Then
            MsgBox("Width and Height of test image must be the same as enrolled images", MsgBoxStyle.Critical, "Eigenface Recognition")
            Label7.Text = "Enrollment failed"
            Exit Function
        End If
        pictureBox2.Image = CType(targetBitmap.Clone(), Bitmap)

        ' Resize image
        Select Case ComboBox1.SelectedIndex
            Case 0
                targetBitmap = ImageProcessor.Resize(targetBitmap, 83, 120)
            Case 1
                targetBitmap = ImageProcessor.Resize(targetBitmap, 53, 90)
            Case 2
                targetBitmap = ImageProcessor.Resize(targetBitmap, 33, 70)
            Case 3
                targetBitmap = ImageProcessor.Resize(targetBitmap, 23, 60)
        End Select

        ' Convert to greyscale if necessary
        Dim c1 = targetBitmap.GetPixel(0, 0)
        Dim c2 = targetBitmap.GetPixel(targetBitmap.Width \ 2, targetBitmap.Height \ 2)
        If c1.R <> c1.G OrElse c1.R <> c1.B OrElse c2.R <> c2.G OrElse c2.R <> c2.B Then
            targetBitmap = ImageProcessor.Greyscale(targetBitmap)
        End If

        ' Flatten image into vector
        For y = 0 To heightOfImage - 1
            For x = 0 To widthOfImage - 1
                testFaceAsInteger.Add(targetBitmap.GetPixel(x, y).R)
            Next
        Next

        ' Histogram Equalization
        If dbEqualised Then
            If CheckBox1.Checked Then
                testFaceAsInteger = ImageProcessor.HistogramEqualisation(testFaceAsInteger)
                frmHistogram.PictureBox5.Image = CType(targetBitmap.Clone(), Bitmap)

                If CheckBox2.Checked Then
                    Dim bmpData = targetBitmap.Clone()
                    Dim counter = 0
                    For y = 0 To heightOfImage - 1
                        For x = 0 To widthOfImage - 1
                            value = testFaceAsInteger(counter)
                            bmpData.SetPixel(x, y, Color.FromArgb(255, value, value, value))
                            counter += 1
                        Next
                    Next
                    frmHistogram.PictureBox6.Image = bmpData
                    frmHistogram.Show()
                End If
            Else
                MsgBox("Face database WAS enrolled with Histogram Equalisation.  Make sure checkbox IS ticked before matching.", MsgBoxStyle.Critical, "Eigenface Recognition")
                Label7.Text = "Matching failed"
                Exit Function
            End If
        Else
            If CheckBox1.Checked Then
                MsgBox("Face database was NOT enrolled with Histogram Equalisation.  Make sure checkbox is NOT ticked before matching.", MsgBoxStyle.Critical, "Eigenface Recognition")
                Label7.Text = "Matching failed"
                Exit Function
            End If
        End If
        targetBitmap.Dispose()

        ' Subtract average face
        For i = 0 To testFaceAsInteger.Count - 1
            testDiff.Add(testFaceAsInteger(i) - averageFace(i))
        Next

        ' Project onto eigenfaces
        For Each face In eigenfaceStore
            tmpValue = 0
            For n = 0 To testDiff.Count - 1
                tmpValue += testDiff(n) * face(n)
            Next
            coefficentTest.Add(tmpValue)
        Next

        ' Compute Euclidean distances
        For Each coeff In coefficentStore
            tmpValue = 0
            For m = 0 To coeff.Count - 1
                tmpValue += Math.Abs(coefficentTest(m) - coeff(m))
            Next
            distance.Add(tmpValue)
        Next

        ' Sorting distances
        distanceSort = CType(distance.Clone(), ArrayList)
        distanceSort.Sort()

        fileAttArray = di.GetFileSystemInfos("*.jpg")

        Dim topIndexes As New List(Of Integer)
        For i = 0 To Math.Min(noOfImages - 1, 2)
            topIndexes.Add(distance.IndexOf(distanceSort(i)))
        Next

        ' Display matched images
        If topIndexes.Count > 0 Then
            TextBox4.Text = distanceSort(0)
            pictureBox3.ImageLocation = fileAttArray(topIndexes(0)).FullName
        End If
        If topIndexes.Count > 1 Then PictureBox5.ImageLocation = fileAttArray(topIndexes(1)).FullName
        If topIndexes.Count > 2 Then PictureBox6.ImageLocation = fileAttArray(topIndexes(2)).FullName

        ' Extract name if applicable
        smallestPos = topIndexes(0)
        fileAtt = fileAttArray(smallestPos)

        If Button9.Enabled AndAlso fileAtt.Name.Contains("_") Then
            i = fileAtt.Name.IndexOf("_")
            TextBox3.Text = fileAtt.Name.Substring(0, i)
        Else
            TextBox3.Text = "Not available"
        End If

        Label7.Text = "Face matched"


    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        dlgAbout.ShowDialog()
    End Sub
End Class
