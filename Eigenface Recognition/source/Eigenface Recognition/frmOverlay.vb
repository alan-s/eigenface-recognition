''''''''''''''''''''''''''''''''' 
'     Eigenface Recognition     '
'        Alan Suleiman          '
'          April 2007           '
'     King's College London     '        
'   alan.suleiman@kcl.ac.uk     '
'''''''''''''''''''''''''''''''''

Public Class frmOverlay

    'Draws a red rectangle on the semi-transparent form to overlay over live feeds
    Private Sub frmOverlay_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Dim rect As Rectangle
        Dim pen As Pen = New Pen(Color.Red, 2)

        rect = New Rectangle(98, 30, 124, 180)
        e.Graphics.DrawRectangle(pen, rect)

    End Sub

End Class