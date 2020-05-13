Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class Ghost

    ':: Game sprites from resource bitmaps ::

    Private GhostLeft(3) As System.Drawing.Bitmap
    Private GhostRight(3) As System.Drawing.Bitmap
    Private GhostUP(3) As System.Drawing.Bitmap
    Private GhostDown(3) As System.Drawing.Bitmap
    Private GhostBlue As System.Drawing.Bitmap

    ':: Properties for this class ::

    Private X As Integer
    Private Y As Integer
    Private MapX As Integer
    Private MapY As Integer
    Private GhostDir As Integer
    Private Chase As Integer
    Private Active As Boolean
    Private ShowPoints As Boolean
    Private ShowPointsTimer As Integer
    Private ShowPointsX As Integer
    Private ShowPointsY As Integer

    Private Image As Drawing.Bitmap
    Private ImageAttributes As New Imaging.ImageAttributes

    Public Sub New(ByVal X As Integer, ByVal Y As Integer, ByVal MapX As Integer, ByVal MapY As Integer, _
    ByVal Chase As Boolean, ByVal Active As Boolean)
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Constructor for creating a sprite instance 
        '------------------------------------------------------------------------------------------------------------------

        ':: Properties ::

        ' X & Y axis for sprite plus X & Y for location in map 
        Me.X = X
        Me.Y = Y
        Me.MapX = MapX
        Me.MapY = MapY

        ' Denotes if the ghost will chase pac 
        Me.Chase = Chase

        ' Set true if the ghost is moving around the maze
        Me.Active = Active

        ':: Load resource image(s) ::

        GhostUP(0) = My.Resources.GhostOrangeUP()
        GhostUP(1) = My.Resources.GhostRedUP()
        GhostUP(2) = My.Resources.GhostCyanUP()
        GhostUP(3) = My.Resources.GhostPinkUP()

        GhostDown(0) = My.Resources.GhostOrangeDown()
        GhostDown(1) = My.Resources.GhostRedDown()
        GhostDown(2) = My.Resources.GhostCyanDown()
        GhostDown(3) = My.Resources.GhostPinkDown()

        GhostLeft(0) = My.Resources.GhostOrangeLeft()
        GhostLeft(1) = My.Resources.GhostRedLeft()
        GhostLeft(2) = My.Resources.GhostCyanLeft()
        GhostLeft(3) = My.Resources.GhostPinkLeft()

        GhostRight(0) = My.Resources.GhostOrangeRight()
        GhostRight(1) = My.Resources.GhostRedRight()
        GhostRight(2) = My.Resources.GhostCyanRight()
        GhostRight(3) = My.Resources.GhostPinkRight()

        GhostBlue = My.Resources.GhostBlue()

        ':: Remove backgrounds ::

        Dim i As Integer

        For i = 0 To 3
            GhostUP(0).MakeTransparent(Color.Black)
            GhostUP(1).MakeTransparent(Color.Black)
            GhostUP(2).MakeTransparent(Color.Black)
            GhostUP(3).MakeTransparent(Color.Black)

            GhostDown(0).MakeTransparent(Color.Black)
            GhostDown(1).MakeTransparent(Color.Black)
            GhostDown(2).MakeTransparent(Color.Black)
            GhostDown(3).MakeTransparent(Color.Black)

            GhostLeft(0).MakeTransparent(Color.Black)
            GhostLeft(1).MakeTransparent(Color.Black)
            GhostLeft(2).MakeTransparent(Color.Black)
            GhostLeft(3).MakeTransparent(Color.Black)

            GhostRight(0).MakeTransparent(Color.Black)
            GhostRight(1).MakeTransparent(Color.Black)
            GhostRight(2).MakeTransparent(Color.Black)
            GhostRight(3).MakeTransparent(Color.Black)
        Next

        GhostBlue.MakeTransparent(Color.Black)

    End Sub

    Public Sub DrawSprite(ByVal Destination As Graphics, ByVal Dir As Integer, ByVal Col As Integer)
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Draws the image to the destination
        '------------------------------------------------------------------------------------------------------------------

        If Col < 4 Then
            Select Case Dir
                Case 0
                    Destination.DrawImage(GhostRight(Col), New Rectangle(Me.X, Me.Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
                Case 1
                    Destination.DrawImage(GhostLeft(Col), New Rectangle(Me.X, Me.Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
                Case 2
                    Destination.DrawImage(GhostUP(Col), New Rectangle(Me.X, Me.Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
                Case 3
                    Destination.DrawImage(GhostDown(Col), New Rectangle(Me.X, Me.Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
            End Select
        Else
            Destination.DrawImage(GhostBlue, New Rectangle(Me.X, Me.Y, 25, 25), 0, 0, 25, 25, GraphicsUnit.Pixel, ImageAttributes)
        End If

    End Sub

    Public Sub SetMoveDir(ByVal GhostDir As Integer)

        Me.GhostDir = GhostDir

    End Sub

    Public Function GetMoveDir() As Integer

        Return Me.GhostDir

    End Function

    Public Sub SetX(ByVal X As Integer)

        Me.X = X

    End Sub

    Public Sub SetY(ByVal Y As Integer)

        Me.Y = Y

    End Sub

    Public Function GetX() As Integer

        Return Me.X

    End Function

    Public Function GetY() As Integer

        Return Me.Y

    End Function

    Public Sub SetMapX(ByVal X As Integer)

        Me.MapX = X

    End Sub

    Public Sub SetMapY(ByVal Y As Integer)

        Me.MapY = Y

    End Sub

    Public Function GetMapX() As Integer

        Return Me.MapX

    End Function

    Public Function GetMapY() As Integer

        Return Me.MapY

    End Function

    Public Function IsChaser() As Boolean

        Return Me.Chase

    End Function

    Public Function IsActive() As Boolean

        Return Me.Active

    End Function

    Public Sub SetActive(ByVal Active As Boolean)

        Me.Active = Active

    End Sub

    Public Sub SetShowPoints(ByVal ShowPoints As Boolean)

        Me.ShowPoints = ShowPoints

    End Sub

    Public Function IsShowingPoints() As Boolean

        Return Me.ShowPoints

    End Function

    Public Sub SetPointsTimer(ByVal Val As Integer)

        Me.ShowPointsTimer = Val

    End Sub

    Public Function GetPointsTimer() As Integer

        Return Me.ShowPointsTimer

    End Function

    Public Sub SetShowPointsX(ByVal x As Integer)

        Me.ShowPointsX = X

    End Sub

    Public Function GetPointsX() As Integer

        Return Me.ShowPointsX

    End Function

    Public Sub SetShowPointsY(ByVal Y As Integer)

        Me.ShowPointsY = Y

    End Sub

    Public Function GetPointsY() As Integer

        Return Me.ShowPointsY

    End Function

End Class
