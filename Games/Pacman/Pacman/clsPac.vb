Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class Pac

    ':: Game sprites from resource bitmaps ::

    Private PacLeft(3) As System.Drawing.Bitmap
    Private PacRight(3) As System.Drawing.Bitmap
    Private PacUP(3) As System.Drawing.Bitmap
    Private PacDown(3) As System.Drawing.Bitmap

    ':: Properties for this class ::

    Private X As Integer
    Private Y As Integer
    Private Image As Drawing.Bitmap
    Private ImageAttributes As New Imaging.ImageAttributes

    Public Sub New()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Constructor for creating a sprite instance 
        '------------------------------------------------------------------------------------------------------------------

        ':: Load resource image(s) ::

        ' Pacman left
        PacLeft(0) = My.Resources.PacNeutral()
        PacLeft(1) = My.Resources.PacLeft01()
        PacLeft(2) = My.Resources.PacLeft02()
        PacLeft(3) = My.Resources.PacLeft03()

        ' Pac right
        PacRight(0) = My.Resources.PacNeutral()
        PacRight(1) = My.Resources.PacRight01()
        PacRight(2) = My.Resources.PacRight02()
        PacRight(3) = My.Resources.PacRight03()

        ' Pac up
        PacUP(0) = My.Resources.PacNeutral()
        PacUP(1) = My.Resources.PacUP01()
        PacUP(2) = My.Resources.PacUP02()
        PacUP(3) = My.Resources.PacUP03()

        ' Pac down
        PacDown(0) = My.Resources.PacNeutral()
        PacDown(1) = My.Resources.PacDown01()
        PacDown(2) = My.Resources.PacDown02()
        PacDown(3) = My.Resources.PacDown03()

        ':: Remove backgrounds ::

        Dim i As Integer

        For i = 0 To 3
            PacLeft(i).MakeTransparent(Color.Black)
            PacRight(i).MakeTransparent(Color.Black)
            PacUP(i).MakeTransparent(Color.Black)
            PacDown(i).MakeTransparent(Color.Black)
        Next

    End Sub

    Public Sub DrawSprite(ByVal Destination As Graphics, ByVal X As Integer, ByVal Y As Integer, ByVal Ani As Integer, ByVal Dir As Integer)
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Draws the image to the destination
        '------------------------------------------------------------------------------------------------------------------

        Select Case Dir

            Case 1
                Destination.DrawImage(PacRight(Ani), New Rectangle(X, Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
            Case 2
                Destination.DrawImage(PacLeft(Ani), New Rectangle(X, Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
            Case 3
                Destination.DrawImage(PacUP(Ani), New Rectangle(X, Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
            Case 4
                Destination.DrawImage(PacDown(Ani), New Rectangle(X, Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
        End Select

    End Sub

End Class
