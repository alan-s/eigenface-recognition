''''''''''''''''''''''''''''''''' 
'     Eigenface Recognition     '
'        Alan Suleiman          '
'          April 2007           '
'     King's College London     '        
'   alan.suleiman@kcl.ac.uk     '
'''''''''''''''''''''''''''''''''

Public Class ImageProcessor

    Shared freqArray As List(Of Integer)
    Shared freqArray2 As List(Of Integer)

    'Crop function based on graphics class in .NET  Crops image starting at (x,y) point using N x M rectangle
    Shared Function Crop(ByVal Source As Bitmap, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32) As Bitmap

        Dim cropped As New Bitmap(width, height)
        Dim g As Graphics = Graphics.FromImage(cropped)

        g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic

        Dim rect As New Rectangle(0, 0, width, height)
        g.DrawImage(Source, rect, x, y, width, height, GraphicsUnit.Pixel)

        g.Dispose()
        Return cropped
        cropped.Dispose()

    End Function

    'Resize function based on graphics class in .NET  Resizes according to new width and height using Bicubic interpolation mode
    Shared Function Resize(ByVal Source As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap

        Dim thumb As New Bitmap(width, height)
        Dim g As Graphics = Graphics.FromImage(thumb)

        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(Source, New Rectangle(0, 0, width, height), New Rectangle(0, 0, Source.Width, Source.Height), GraphicsUnit.Pixel)

        g.Dispose()
        Return thumb
        thumb.Dispose()

    End Function

    'Grayscale method works by averaging the Red, Green and Blue compnent of a pixel. Alpha is untouched at 255
    '(R + G  + B) / 3
    Shared Function Greyscale(ByVal image As Image) As Bitmap

        Dim ia As Imaging.ImageAttributes = New Imaging.ImageAttributes
        Dim cm As Imaging.ColorMatrix = New Imaging.ColorMatrix
        Dim g As Graphics = Graphics.FromImage(image)

        cm.Matrix00 = 1 / 3.0F
        cm.Matrix01 = 1 / 3.0F
        cm.Matrix02 = 1 / 3.0F
        cm.Matrix10 = 1 / 3.0F
        cm.Matrix11 = 1 / 3.0F
        cm.Matrix12 = 1 / 3.0F
        cm.Matrix20 = 1 / 3.0F
        cm.Matrix21 = 1 / 3.0F
        cm.Matrix22 = 1 / 3.0F

        ia.SetColorMatrix(cm)
        g.DrawImage(image, New Rectangle(0, 0, image.Width, image.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, ia)

        Return image
        ia.Dispose()
        g.Dispose()
        image.Dispose()

    End Function

    'Histogram Equalization
    Shared Function HistogramEqualisation(ByVal source As List(Of Integer)) As List(Of Integer)


        Dim cumuFreqArray As List(Of Integer)
        Dim imgAsInteger As List(Of Integer)

        freqArray = New List(Of Integer)
        freqArray2 = New List(Of Integer)
        cumuFreqArray = New List(Of Integer)
        imgAsInteger = New List(Of Integer)

        imgAsInteger = source


        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim i As Integer = 0
        Dim pixValue As Integer = 0
        Dim valueIn As Integer = 0


        'populate frequency before array
        For i = 0 To 255
            freqArray.Add(0)
        Next

        'populate frequency after array
        For i = 0 To 255
            freqArray2.Add(0)
        Next

        'calculate frequency of gray levels 0 - 255
        For i = 0 To imgAsInteger.Count - 1
            pixValue = imgAsInteger.Item(i)
            freqArray.Item(pixValue) = freqArray.Item(pixValue) + 1

        Next

        'calculate cumulative frequency of gray levels
        cumuFreqArray.Add(freqArray.Item(0))

        For i = 1 To 255
            cumuFreqArray.Add(cumuFreqArray.Item(i - 1) + freqArray.Item(i))
        Next

        'calculate new pixel values according to cumulative shift
        Dim alpha As Double = 255 / cumuFreqArray.Item(255)

        For i = 0 To imgAsInteger.Count - 1

            valueIn = imgAsInteger.Item(i)
            pixValue = cumuFreqArray.Item(valueIn) * alpha
            imgAsInteger.Item(i) = pixValue
        Next

        'calculate new frequency after histogram equalisation
        For i = 0 To imgAsInteger.Count - 1
            pixValue = imgAsInteger.Item(i)
            freqArray2.Item(pixValue) = freqArray2.Item(pixValue) + 1

        Next

        cumuFreqArray.Clear()
        Return imgAsInteger
        imgAsInteger.Clear()

    End Function

    Public Shared Function getFrequencyBefore() As List(Of Integer)

        Return freqArray

    End Function

    Public Shared Function getFrequencyAfter() As List(Of Integer)

        Return freqArray2

    End Function

End Class
