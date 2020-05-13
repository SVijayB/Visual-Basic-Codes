
' ______________________________________________________________________________
'         ______      ___       ______   ___  ___       ___       __   __ 
'        |   _  \    /   \     /      | |   \/   |     /   \     |  \ |  | 
'        |  |_)  |  /  ^  \   |  ,----' |  \  /  |    /  ^  \    |   \|  | 
'        |   ___/  /  /_\  \  |  |      |  |\/|  |   /  /_\  \   |       | 
'        |  |     /  _____  \ |  `----  |  |  |  |  /  _____  \  |  |\   | 
'        |__|    /__/     \__\ \______| |__|  |__| /__/     \__\ |__| \__|
'
'                Copyright ©2008 Trent Jackson All Rights Reserved  
'                      Email: trentjackson888@bigpond.com.au
' ______________________________________________________________________________

' Product:  Pacman Video game
' Release:  Beta     
' Version:  0.02
' Platform: Windows XP 
' Date:     March 09, 2009

' TERMS OF USE:
' This program is free software. Redistribution in the form of source code only, 
' strictly for non-profit purposes, with or without modification is permitted. 
' The author accepts no liability for anything that may result from the usage of 
' this product. This notice must not be removed or altered in any way whatsoever. 
' ______________________________________________________________________________


Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Media

Public Class frmGame

    ' APIs
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Integer) As Short
    Private Declare Function timeGetTime Lib "winmm.dll" () As Integer
    Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Integer) As Integer
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    ':: Game variables ::

    ' Integers
    Private PacX As Integer                 ' Pac's current X location on the screen
    Private PacY As Integer                 ' Pac's current Y location on the screen
    Private PacMapX As Integer              ' Pac's current location X in the maze map 
    Private PacMapY As Integer              ' Pac's current location Y in the maze map 
    Private PacDir As Integer               ' Pac's current direction: left, right, up or down
    Private NewDirection As Integer         ' The next direction that the player wants to move pac
    Private Sync As Integer                 ' Synchronization between many procedures 
    Private TimeScalerA As Integer          ' Scale base delay 4:1
    Private TimeScalerB As Integer          ' Scale base delay 10:1
    Private TimeScalerC As Integer          ' Scale base delay 25:1
    Private ProcessStartmS As Integer       ' Regulated delay usage: sample of time in mS
    Private AnimatePac As Integer           ' Alternate bewteen pac images, thus animating
    Private ToggleEatSnd As Integer         ' When pac eats a pill alternate between two sounds
    Private Top5PlayerName(5) As String     ' Contains top 5 player names 
    Private Top5PlayerScore(5) As Integer   ' Top 5 player scores
    Private Score As Integer                ' Player's score ...
    Private FlashPowerPill As Integer       ' Variable used to toggle power pills on / off 
    Private PowerPillTimer As Integer       ' When pac gets a power pill it last for ~5 secs
    Private PillsEaten As Integer           ' Total num of pills that pac has eaten  
    Private IntroTimer As Integer           ' The game begins with an intro that lasts a few secs
    Private ReleaseGhostsTimer As Integer   ' Ghosts are released from the box every few secs 
    Private Lives As Integer                ' Total num of lives pac has remaining ...
    Private PacDeadTimer As Integer         ' When pac gets caught by a ghost pause for a few secs
    Private ShowGameOverTimer As Integer    ' When the game has ended pause for a few secs 
    Private GhostEatenSndTimer As Integer   ' When pac eats a ghots (n) time is allowed for a sound

    ' Strings
    Private MapRow(30) As String            ' Used to store a map of the maze in the game ...
    Public Shared PlayerName As String      ' Used to retrieve player's name from dialog-based prompt 

    ' Boolean flags 
    Private MoveLeft As Boolean             ' Keyboard usage: set when cursor key left is pressed 
    Private MoveRight As Boolean            ' Keyboard usage: set when cursor key right is pressed
    Private MoveUP As Boolean               ' Keyboard usage: set when cursor key up is pressed  
    Private MoveDown As Boolean             ' Keyboard usage: set when cursor key down is pressed
    Private GameOver As Boolean = True      ' Set when game no game is playing ...
    Private NewFood As Boolean              ' Set when a new apple should be administered
    Private ShowGameOver As Boolean         ' Text message: set when showing player "game over"  
    Private AppRunning As Boolean = True    ' Set when entire application is active
    Private NewGame As Boolean              ' Set when a new game has started, reset variables etc
    Private PacMoving As Boolean            ' Set when pac is in motion
    Private GotPowerPill As Boolean         ' Set when pac has a power pill
    Private DoIntro As Boolean              ' Set when the intro is current
    Private StartIntro As Boolean           ' Set when the game first begins
    Private ReleaseGhost As Boolean         ' Set when a ghots is released from the box
    Private PacDead As Boolean              ' Set if pac gets caught by a ghost 
    Private PlayGhostEatenSnd As Boolean    ' Set when the ghost eaten sound is played

    ':: Object references / instances ::

    ' Random num obj
    Private objRandom As New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer))

    ' Imaging attributes 
    Private ImageAttributes As New Imaging.ImageAttributes

    ' Primary screen buffer references
    Private ScrnBufferBmp As Bitmap
    Private GraphicsBuffer As Graphics

    ' Game sprites from resource bitmaps 
    Private Pill As System.Drawing.Bitmap
    Private PowerPill As System.Drawing.Bitmap
    Private Maze As System.Drawing.Bitmap
    Private PacTitle As System.Drawing.Bitmap

    ' Instance of pacman  
    Private Pac As New Pac()

    ' Ref of ghosts  
    Private Ghosts(4) As Ghost

    ' Sound effects channel B 
    Private Intro As New SndChanB("Intro.wav")
    Private Eat1 As New SndChanB("PelletEat1.wav")
    Private Eat2 As New SndChanB("PelletEat2.wav")
    Private GotPac As New SndChanB("PacEaten.wav")
    Private GotGhost As New SndChanB("GhostEaten.wav")

    ' Sound effects Channel A 
    Private Invincible As New SoundPlayer(My.Resources.Invincible)
    Private Siren As New SoundPlayer(My.Resources.Siren)

    ' Collision detection usage
    Private GhostRect As New Rectangle
    Private PacRect As New Rectangle
    Private Collision As New Rectangle

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Load resource images
        PowerPill = My.Resources.PowerPill()
        Pill = My.Resources.Pill()
        Maze = My.Resources.Maze()
        PacTitle = My.Resources.PacTitle()

        ' Remove background from image(s) -- thus producing a sprite 
        Pill.MakeTransparent(Color.Black)
        PowerPill.MakeTransparent(Color.Black)
        PacTitle.MakeTransparent(Color.Black)

        ' Create a new graphics buffer (this handles the entire screen)
        ScrnBufferBmp = New Bitmap(500, 600, Me.CreateGraphics())
        GraphicsBuffer = Graphics.FromImage(ScrnBufferBmp)

        ' Enable double buffering ensuring minimal flicker
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)

        ' Set collision detect rects width & height ...
        PacRect.Width = 16
        PacRect.Height = 16
        GhostRect.Width = 16
        GhostRect.Height = 16

        ' Make the game visible 
        Me.Show()

        ' Initialize delay
        Call InitDelay()

        ' Load high scores from disk
        Call LoadScoreTable()

        ' Merge to main loop
        Call MainLoop()

    End Sub

    Private Sub MainLoop()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Managing the game by continuously cycle through the game procedures
        '------------------------------------------------------------------------------------------------------------------

        Do While AppRunning

            ' Time prescaler 3:1 
            TimeScalerA += 1
            TimeScalerA = TimeScalerA Mod 3

            ' Prescaler 10:1
            TimeScalerB += 1
            TimeScalerB = TimeScalerB Mod 10

            ' Prescaler 25:1
            TimeScalerC += 1
            TimeScalerC = TimeScalerC Mod 25

            ' Fetch user input
            Call ScanKeyboard()

            ' Start new game
            If NewGame Then Call InitGame()

            ' Run title screen if no game is playing
            If GameOver Then
                If Not ShowGameOver Then
                    Call WipeGraphicsBuffer()
                    Call ShowTitle()
                    Call ShowScoreTable()

                Else ':: Game over, inform the player, flash "game over" for a few secs

                    ShowGameOverTimer += 1
                    If ShowGameOverTimer = 200 Then
                        ShowGameOver = False
                        ShowGameOverTimer = 0
                    End If

                    ' Main calls while displaying game over msg ... 
                    Call MainCallsAtIdle()

                    ' Flash "game over"
                    If TimeScalerC < 12 Then
                        DrawText(GraphicsBuffer, "GAME", 189, 260, 16, FontStyle.Bold, Brushes.Red)
                        DrawText(GraphicsBuffer, "OVER", 190, 280, 16, FontStyle.Bold, Brushes.Red)
                    End If
                End If

            Else ':: The game is being played ::

                ' Play intro music at the start of a new game 
                If DoIntro Then
                    If StartIntro Then
                        Intro.Open()
                        Intro.PlaySnd()
                        StartIntro = False
                    End If

                    ' Main calls during intro ...
                    Call MainCallsAtIdle()

                    ' Flash "get ready"
                    If TimeScalerC < 12 Then
                        DrawText(GraphicsBuffer, "Get Ready!", 168, 220, 14, FontStyle.Bold + FontStyle.Italic, Brushes.Red)
                    End If

                    ' Intro runs for approx 5 sec, begin game sound when finished
                    IntroTimer += 1
                    If IntroTimer = 150 Then
                        Siren.PlayLooping()
                        DoIntro = False
                        IntroTimer = 0
                    End If

                Else ':: Game now in progress :: 

                    If TimeScalerA = 2 Then

                        ' Inc synchronize variable
                        Sync += 1
                        Sync = Sync Mod 2

                        If PlayGhostEatenSnd Then
                            GhostEatenSndTimer += 1
                            If GhostEatenSndTimer = 10 Then
                                GhostEatenSndTimer = 0
                                PlayGhostEatenSnd = False
                            End If
                        End If

                        If Not PacDead Then

                            ' Main calls during game play
                            Call WipeScore()
                            Call Wipelives()
                            Call ProcessInput()
                            Call DoPacman()
                            Call DoGhosts()
                            Call RenderMaze()
                            Call RenderPills()
                            Call RenderPac()
                            Call RenderGhosts()
                            Call CheckCollision()
                            Call ShowScore()
                            Call ShowHighScore()
                            Call ShowLives()

                        Else ':: Pac dead -- pause for a moment, subtract a life and restart the game ::

                            PacDeadTimer += 1
                            If PacDeadTimer = 25 Then
                                PacDead = False
                                Lives -= 1
                                PacDeadTimer = 0

                                If Lives >= 0 Then
                                    NewGame = True
                                    Siren.PlayLooping()
                                Else
                                    GameOver = True
                                    ShowGameOver = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            ' Draw screen buffer to the actual screen (straight to the form in this case)
            Me.Invalidate(New Rectangle(0, 0, 500, 600))

            ' RegulatedDelay 25mS
            Call RegulatedDelay(25)

        Loop

        ' Loop exited, terminate app ... 
        Me.Close()

    End Sub

    Private Sub InitGame()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Initialize game objects / variables with starting values
        '------------------------------------------------------------------------------------------------------------------

        ' Load game map (the maze)
        ' . = Pill
        ' o = Power Pill
        ' W = Wall

        MapRow(0) = "WWWWWWWWWWWWWWWWWWWWWWWWWWWW"
        MapRow(1) = "W............WW............W"
        MapRow(2) = "W.WWWW.WWWWW.WW.WWWWW.WWWW.W"
        MapRow(3) = "WoWWWW.WWWWW.WW.WWWWW.WWWWoW"
        MapRow(4) = "W.WWWW.WWWWW.WW.WWWWW.WWWW.W"
        MapRow(5) = "W..........................W"
        MapRow(6) = "W.WWWW.WW.WWWWWWWW.WW.WWWW.W"
        MapRow(7) = "W.WWWW.WW.WWWWWWWW.WW.WWWW.W"
        MapRow(8) = "W......WW....WW....WW......W"
        MapRow(9) = "WWWWWW.WWWWW WW WWWWW.WWWWWW"
        MapRow(10) = "     W.WWWWW WW WWWWW.W     "
        MapRow(11) = "     W.WW          WW.W     "
        MapRow(12) = "     W.WW WWWWWWWW WW.W     "
        MapRow(13) = "WWWWWW.WW W      W WW.WWWWWW"
        MapRow(14) = "      .   W      W   .      "
        MapRow(15) = "WWWWWW.WW W      W WW.WWWWWW"
        MapRow(16) = "     W.WW WWWWWWWW WW.W     "
        MapRow(17) = "     W.WW          WW.W     "
        MapRow(18) = "     W.WW WWWWWWWW WW.W     "
        MapRow(19) = "WWWWWW.WW WWWWWWWW WW.WWWWWW"
        MapRow(20) = "W............WW............W"
        MapRow(21) = "W.WWWW.WWWWW.WW.WWWWW.WWWW.W"
        MapRow(22) = "W.WWWW.WWWWW.WW.WWWWW.WWWW.W"
        MapRow(23) = "Wo..WW.......  .......WW..oW"
        MapRow(24) = "WWW.WW.WW.WWWWWWWW.WW.WW.WWW"
        MapRow(25) = "WWW.WW.WW.WWWWWWWW.WW.WW.WWW"
        MapRow(26) = "W......WW....WW....WW......W"
        MapRow(27) = "W.WWWWWWWWWW.WW.WWWWWWWWWW.W"
        MapRow(28) = "W.WWWWWWWWWW.WW.WWWWWWWWWW.W"
        MapRow(29) = "W..........................W"
        MapRow(30) = "WWWWWWWWWWWWWWWWWWWWWWWWWWWW"

        ' Create ghost instances
        Ghosts(0) = New Ghost(212, 268, 14, 14, True, False)
        Ghosts(1) = New Ghost(212, 268, 14, 14, False, False)
        Ghosts(2) = New Ghost(212, 268, 14, 14, False, False)
        Ghosts(3) = New Ghost(212, 268, 14, 14, True, False)

        ' Reset game vars ...
        If GameOver Then
            Lives = 3
            Score = 0
        End If

        PacX = 212
        PacY = 414
        PacMapX = 14
        PacMapY = 23
        NewDirection = 1
        PacDir = 1
        AnimatePac = 3
        Sync = 0
        PillsEaten = 0
        ReleaseGhostsTimer = 0
        GameOver = False
        DoIntro = True
        StartIntro = True
        GotPowerPill = False
        NewGame = False
        Siren.Stop()

    End Sub

    Private Sub ProcessInput()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Decide what to do with the player's keystrokes
        '------------------------------------------------------------------------------------------------------------------

        Dim temp As String

        If PacMapX < 28 Then
            If PacMapX > 1 Then
                If MoveRight Then
                    temp = MapRow(PacMapY).Substring(PacMapX, 1)
                    If temp <> "W" Then
                        NewDirection = 1
                    Else
                        NewDirection = PacDir
                    End If

                ElseIf MoveLeft Then
                    temp = MapRow(PacMapY).Substring(PacMapX - 2, 1)
                    If temp <> "W" Then
                        NewDirection = 2
                    Else
                        NewDirection = PacDir
                    End If

                ElseIf MoveUP Then
                    temp = MapRow(PacMapY - 1).Substring(PacMapX - 1, 1)
                    If temp <> "W" Then
                        NewDirection = 3
                    Else
                        NewDirection = PacDir
                    End If

                ElseIf MoveDown Then
                    temp = MapRow(PacMapY + 1).Substring(PacMapX - 1, 1)
                    If temp <> "W" Then
                        NewDirection = 4
                    Else
                        NewDirection = PacDir
                    End If
                End If
            End If
        End If

        MoveUP = False
        MoveDown = False
        MoveLeft = False
        MoveRight = False

    End Sub

    Private Sub DoGhosts()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Handle all the ghosts in the game
        '------------------------------------------------------------------------------------------------------------------

        Dim GhostMapX As Integer
        Dim GhostMapY As Integer
        Dim GhostX As Integer
        Dim GhostY As Integer
        Dim GhostDir As Integer
        Dim temp As String
        Dim i As Integer

        ReleaseGhostsTimer += 1
        ReleaseGhostsTimer = ReleaseGhostsTimer Mod 50

        For i = 0 To 3

            ':: Fetch coords ::

            GhostMapX = Ghosts(i).GetMapX
            GhostMapY = Ghosts(i).GetMapY
            GhostX = Ghosts(i).GetX
            GhostY = Ghosts(i).GetY
            GhostDir = Ghosts(i).GetMoveDir()

            ':: Release the ghosts one at a time ::

            If Not Ghosts(i).IsActive Then
                If ReleaseGhostsTimer = 49 Then
                    ReleaseGhost = True
                End If
                If ReleaseGhost Then
                    If Sync = 0 Then
                        GhostX = 204
                        GhostY = 222
                        GhostMapX = 14
                        GhostMapY = 11
                        Ghosts(i).SetMoveDir(Rnd() * 3)
                        Ghosts(i).SetActive(True)
                        ReleaseGhostsTimer = 0
                        ReleaseGhost = False
                    End If
                End If
            End If

            If Sync = 0 Then
                If Not GotPowerPill Then
                    If Ghosts(i).IsActive Then
                        If Ghosts(i).IsChaser Then

                            ' Determine which axis needs the most attention: X or Y ...
                            ' which one is the most out of alignment from pac's coordinates
                            Dim DiffX As Integer
                            Dim DiffY As Integer

                            If PacX > GhostX Then
                                DiffX = PacX - GhostX
                            Else
                                DiffX = GhostX - PacX
                            End If
                            If PacY > GhostY Then
                                DiffY = PacY - GhostY
                            Else
                                DiffY = GhostY - PacY
                            End If

                            ' The axis that is further apart takes priority ...
                            If DiffX > DiffY Then
                                If GhostX > PacX Then
                                    GhostDir = 1
                                Else
                                    GhostDir = 0
                                End If
                            Else
                                If GhostY < PacY Then
                                    GhostDir = 3
                                Else
                                    GhostDir = 2
                                End If
                            End If
                        End If
                    End If

                Else ':: Abort chase, move in the opposite direction -- pac has a power pill ::

                    Select Case PacDir
                        Case 1
                            GhostDir = 1
                        Case 2
                            GhostDir = 0
                        Case 3
                            GhostDir = 3
                        Case 4
                            GhostDir = 2
                    End Select
                End If

                ':: Validate new moves, make sure theres not a wall at the coordinates ::

                temp = "W"
                Dim Attemps As Integer

                Do Until temp <> "W"
                    Select Case GhostDir
                        Case 0
                            temp = MapRow(GhostMapY).Substring(GhostMapX, 1)
                            If Ghosts(i).IsActive Then
                                If GhostMapY = 14 Then temp = "W"
                            End If
                        Case 1
                            temp = MapRow(GhostMapY).Substring(GhostMapX - 2, 1)
                            If Ghosts(i).IsActive Then
                                If GhostMapY = 14 Then temp = "W"
                            End If
                        Case 2
                            temp = MapRow(GhostMapY - 1).Substring(GhostMapX - 1, 1)
                        Case 3
                            temp = MapRow(GhostMapY + 1).Substring(GhostMapX - 1, 1)
                    End Select

                    ' If there's a wall then fetch new moves ...
                    If temp = "W" Then
                        If Not Ghosts(i).IsChaser Then
                            GhostDir = Rnd() * 4
                            GhostDir = GhostDir Mod 4
                        Else
                            If Not GotPowerPill Then
                                If Attemps < 1 Then
                                    Select Case GhostDir
                                        Case 0, 1
                                            If GhostY > PacY Then
                                                GhostDir = 2
                                            Else
                                                GhostDir = 3
                                            End If

                                        Case 2, 3
                                            If GhostX > PacX Then
                                                GhostDir = 1
                                            Else
                                                GhostDir = 0
                                            End If
                                    End Select
                                Else
                                    GhostDir = Rnd() * 4
                                    GhostDir = GhostDir Mod 4
                                End If

                                Attemps += 1

                            Else
                                GhostDir = Rnd() * 4
                                GhostDir = GhostDir Mod 4
                            End If
                        End If
                    End If
                Loop
            End If

            ':: Apply the moves ::

            Select Case GhostDir
                Case 0 ':: Move right ::
                    GhostX += 8
                    If Sync = 1 Then
                        GhostMapX += 1
                    End If
                Case 1 ':: Left ::
                    GhostX -= 8
                    If Sync = 1 Then
                        GhostMapX -= 1
                    End If
                Case 2 ' :: UP ::
                    GhostY -= 8
                    If Sync = 1 Then
                        GhostMapY -= 1
                    End If
                Case 3 ' :: Down ::
                    GhostY += 8
                    If Sync = 1 Then
                        GhostMapY += 1
                    End If
            End Select

            ' Apply new values to class properties 
            Ghosts(i).SetMapX(GhostMapX)
            Ghosts(i).SetMapY(GhostMapY)
            Ghosts(i).SetX(GhostX)
            Ghosts(i).SetY(GhostY)
            Ghosts(i).SetMoveDir(GhostDir)
        Next

    End Sub

    Private Sub CheckCollision()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Handle collisions between pac and the ghosts
        '------------------------------------------------------------------------------------------------------------------

        Dim Overlap As New Rectangle
        Dim i As Integer

        For i = 0 To 3

            If Ghosts(i).IsActive Then

                GhostRect.X = Ghosts(i).GetX + 8
                GhostRect.Y = Ghosts(i).GetY + 8

                Overlap = Rectangle.Intersect(PacRect, GhostRect)

                If Overlap.X <> 0 Then
                    If Sync = 1 Then

                        If GotPowerPill Then

                            GotGhost.Open()
                            GotGhost.PlaySnd()
                            PlayGhostEatenSnd = True

                            ' 500 pts for eating a ghost
                            Score += 500
                            Ghosts(i).SetShowPoints(True)
                            Ghosts(i).SetPointsTimer(0)
                            Ghosts(i).SetShowPointsX(Ghosts(i).GetX)
                            Ghosts(i).SetShowPointsY(Ghosts(i).GetY + 3)

                            Ghosts(i).SetActive(False)
                            Ghosts(i).SetX(204)
                            Ghosts(i).SetY(268)
                            Ghosts(i).SetMapX(14)
                            Ghosts(i).SetMapY(14)
                            ReleaseGhostsTimer = 0

                        Else

                            PacDead = True
                            Siren.Stop()
                            GotPac.Open()
                            GotPac.PlaySnd()
                        End If
                    End If
                End If
            End If
        Next

    End Sub

    Private Sub RenderGhosts()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Render all ghost sprites to the screen
        '------------------------------------------------------------------------------------------------------------------

        Dim i As Integer
        Dim Col As Integer
        Dim ShowPointsTimer As Integer
        Static FlashGhost As Integer

        FlashGhost += 1
        FlashGhost = FlashGhost Mod 2

        For i = 0 To 3

            ' Set col to blue if player has a power pill ...
            If GotPowerPill Then Col = 4 Else Col = i

            ' Alternate col between blue and the actual ghost col when time is running out
            If PowerPillTimer > 50 Then
                If FlashGhost Then Col = i Else Col = 4
            End If

            ' Render ghosts ...
            Ghosts(i).DrawSprite(GraphicsBuffer, Ghosts(i).GetMoveDir, Col)

            ' Show points for a few secs after ghost has been eaten 
            If Ghosts(i).IsShowingPoints Then
                ShowPointsTimer = Ghosts(i).GetPointsTimer

                ShowPointsTimer += 1
                Ghosts(i).SetPointsTimer(ShowPointsTimer)

                If ShowPointsTimer < 25 Then
                    DrawText(GraphicsBuffer, "500", Ghosts(i).GetPointsX, Ghosts(i).GetPointsY, 10, FontStyle.Bold, Brushes.Red)
                Else
                    Ghosts(i).SetShowPoints(False)
                End If
            End If
        Next

    End Sub

    Private Sub DoPacman()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Handle pacman through the maze
        '------------------------------------------------------------------------------------------------------------------

        Dim temp As String

        If Sync = 0 Then PacDir = NewDirection

        If PacMoving Then
            AnimatePac += 1
            AnimatePac = AnimatePac Mod 4
        Else
            AnimatePac = 3
        End If

        Select Case PacDir
            Case 1 ':: Move Right ::

                If PacMapY = 14 And PacMapX = 28 Then
                    PacMapX = 1
                    PacX = 4

                Else

                    temp = MapRow(PacMapY).Substring(PacMapX - 1, 1)

                    If temp = "." Then
                        MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                        + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                        Score += 100
                        PillsEaten += 1
                        Call PlayEatenPillSnds()

                    ElseIf temp = "o" Then
                        MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                        + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                        GotPowerPill = True
                        PowerPillTimer = 0
                        Invincible.PlayLooping()
                    End If

                    temp = MapRow(PacMapY).Substring(PacMapX, 1)

                    If temp <> "W" Then
                        PacX += 8
                        PacMoving = True
                        If Sync = 1 Then
                            PacMapX += 1
                        End If
                    Else
                        PacMoving = False
                    End If
                    End If

            Case 2 ':: Move Left ::

                If PacMapY = 14 And PacMapX = 1 Then
                    PacMapX = 28
                    PacX = 420

                Else

                    temp = MapRow(PacMapY).Substring(PacMapX - 1, 1)

                    If temp = "." Then
                        MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                        + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                        Score += 100
                        PillsEaten += 1
                        Call PlayEatenPillSnds()

                    ElseIf temp = "o" Then
                        MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                        + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                        GotPowerPill = True
                        PowerPillTimer = 0
                        Invincible.PlayLooping()
                    End If

                    temp = MapRow(PacMapY).Substring(PacMapX - 2, 1)

                    If temp <> "W" Then
                        PacX -= 8
                        PacMoving = True
                        If Sync = 1 Then
                            PacMapX -= 1
                        End If
                    Else
                        PacMoving = False
                    End If
                End If

            Case 3 ':: Move UP ::

                temp = MapRow(PacMapY).Substring(PacMapX - 1, 1)

                If temp = "." Then
                    MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                    + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                    Score += 100
                    PillsEaten += 1
                    Call PlayEatenPillSnds()

                ElseIf temp = "o" Then
                    MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                    + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                    GotPowerPill = True
                    PowerPillTimer = 0
                    Invincible.PlayLooping()
                End If

                temp = MapRow(PacMapY - 1).Substring(PacMapX - 1, 1)

                If temp <> "W" Then
                    PacY -= 8
                    PacMoving = True
                    If Sync = 1 Then
                        PacMapY -= 1
                    End If
                Else
                    PacMoving = False
                End If

            Case 4 ':: Move Down ::

                temp = MapRow(PacMapY).Substring(PacMapX - 1, 1)

                If temp = "." Then
                    MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                    + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                    Score += 100
                    PillsEaten += 1
                    Call PlayEatenPillSnds()

                ElseIf temp = "o" Then
                    MapRow(PacMapY) = MapRow(PacMapY).Substring(0, PacMapX - 1) + "x" _
                    + MapRow(PacMapY).Substring(PacMapX, MapRow(PacMapY).Length - PacMapX)

                    GotPowerPill = True
                    PowerPillTimer = 0
                    Invincible.PlayLooping()
                End If

                temp = MapRow(PacMapY + 1).Substring(PacMapX - 1, 1)

                If temp <> "W" Then
                    PacY += 8
                    PacMoving = True
                    If Sync = 1 Then
                        PacMapY += 1
                    End If
                Else
                    PacMoving = False
                End If
        End Select

        If GotPowerPill Then
            PowerPillTimer += 1
            If PowerPillTimer = 85 Then
                GotPowerPill = False
                PowerPillTimer = 0
                Siren.PlayLooping()
            End If
        End If

        PacRect.X = PacX + 8
        PacRect.Y = PacY + 8

        If PillsEaten = 240 Then NewGame = True

    End Sub

    Private Sub PlayEatenPillSnds()

        ToggleEatSnd += 1
        ToggleEatSnd = ToggleEatSnd Mod 2

        If Not PlayGhostEatenSnd Then
            If ToggleEatSnd Then
                Eat1.Open()
                Eat1.PlaySnd()
            Else
                Eat2.Open()
                Eat2.PlaySnd()
            End If
        End If

    End Sub

    Private Sub RenderPills()

        Dim i As Integer
        Dim j As Integer
        Dim temp As String

        FlashPowerPill += 1
        FlashPowerPill = FlashPowerPill Mod 8

        ' Power pills not to be flahed during the intro or while displaying game over msg
        If DoIntro Or ShowGameOver Then
            FlashPowerPill = 0
        End If

        For j = 1 To 29
            For i = 1 To 28
                temp = Mid$(MapRow(j), i, 1)
                If temp = "." Then
                    GraphicsBuffer.DrawImage(Pill, New Rectangle((i * 16) - 10, (j * 16) + 55, 6, 6), 0, 0, 6, 6, GraphicsUnit.Pixel, ImageAttributes)

                ElseIf temp = "o" Then
                    If FlashPowerPill < 4 Then
                        GraphicsBuffer.DrawImage(PowerPill, New Rectangle((i * 16) - 14, (j * 16) + 51, 14, 14), 0, 0, 14, 14, GraphicsUnit.Pixel, ImageAttributes)
                    End If
                End If
            Next
        Next

    End Sub

    Private Sub RenderMaze()

        GraphicsBuffer.DrawImage(Maze, New Rectangle(0, 50, 448, 496), 0, 0, 448, 496, GraphicsUnit.Pixel, ImageAttributes)

    End Sub

    Private Sub RenderPac()

        Pac.DrawSprite(GraphicsBuffer, PacX, PacY, AnimatePac, PacDir)

    End Sub

    Private Sub ShowTitle()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Render game title and copyright notice 
        '------------------------------------------------------------------------------------------------------------------

        Dim i As Integer

        GraphicsBuffer.DrawImage(PacTitle, New Rectangle(172, 101, 50, 50), 0, 0, 50, 50, GraphicsUnit.Pixel, ImageAttributes)

        For i = 0 To 3
            DrawText(GraphicsBuffer, "PA   MAN", 70 + i, 88 + i, 48, FontStyle.Bold, IIf(i, Brushes.Yellow, Brushes.Green))
        Next

        For i = 0 To 1
            DrawText(GraphicsBuffer, "-------:: By Trent Jackson Copyright ©2008 ::-------", 80 + i, 158 + i, 10, FontStyle.Regular, IIf(i, Brushes.White, Brushes.Black))
        Next

        Dim oPen As System.Drawing.Pen

        oPen = New System.Drawing.Pen(Color.Cyan)
        GraphicsBuffer.DrawRectangle(oPen, 75, 93, 300, 90)
        oPen.Dispose()

        If TimeScalerB < 5 Then
            DrawText(GraphicsBuffer, "Press F2 To Play", 143, 245, 14, FontStyle.Bold, Brushes.Red)
        Else
            DrawText(GraphicsBuffer, "Press F2 To Play", 143, 245, 14, FontStyle.Bold, Brushes.Gold)
        End If

    End Sub

    Private Sub ShowScoreTable()

        Dim i As Integer
        Dim j As Integer

        ' Purpose: Render sorted high score table (top 5 players)
        For i = 0 To 1
            DrawText(GraphicsBuffer, "TOP 5 PLAYERS", 95 + i, 325 + i, 23, FontStyle.Bold, IIf(i, Brushes.Gold, Brushes.Yellow))
        Next

        For j = 0 To 1
            For i = 0 To 4
                DrawText(GraphicsBuffer, Top5PlayerName(i), 100 + j, 375 + (i * 30) + j, 16, FontStyle.Regular, IIf(j, Brushes.LightGreen, Brushes.Green))
                DrawText(GraphicsBuffer, Top5PlayerScore(i), 285 + j, 375 + (i * 30) + j, 16, FontStyle.Regular, IIf(j, Brushes.LightGreen, Brushes.Green))
            Next
        Next

    End Sub

    Private Sub ShowLives()

        Dim i As Integer

        For i = 0 To Lives - 1
            GraphicsBuffer.DrawImage(PacTitle, New Rectangle((i * 35) + 5, 550, 25, 25), 0, 0, 50, 50, GraphicsUnit.Pixel, ImageAttributes)
        Next

    End Sub

    Private Sub MainCallsAtIdle()

        Call WipeScore()
        Call Wipelives()
        Call RenderMaze()
        Call RenderPills()
        Call RenderPac()
        Call ShowScore()
        Call ShowHighScore()
        Call ShowLives()

    End Sub

    Private Sub CheckForNewHighScore()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Check to see if player's score has beaten any of the top 5 scores, prompt player for their name if so
        '------------------------------------------------------------------------------------------------------------------

        Dim i As Integer
        Dim j As Integer

        For i = 0 To 4
            If Score > Top5PlayerScore(i) Then

                ' Prompt player for their name
                If frmNameEntry.ShowDialog() = Windows.Forms.DialogResult.OK Then

                    ' Sort table
                    For j = 4 To (i + 1) Step -1
                        Top5PlayerName(i) = Top5PlayerName(j - 1)
                        Top5PlayerScore(j) = Top5PlayerScore(j - 1)
                    Next

                    ' Add new player
                    If PlayerName = vbNullString Then
                        PlayerName = "Anonymous"
                    End If

                    Top5PlayerName(i) = PlayerName
                    Top5PlayerScore(i) = Score

                    Exit Sub
                End If
            End If
        Next

    End Sub

    Private Sub LoadScoreTable()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Load saved high score from file, assign default values if no high scores exist
        '------------------------------------------------------------------------------------------------------------------

        Dim i As Integer

        Dim FILENAME As String = (CurDir() & "\high scores.txt")

        ' Get a StreamReader class that can be used to read the file  
        Dim objStreamReader As StreamReader

        Try
            ' OPen file
            objStreamReader = File.OpenText(FILENAME)

            ' Read player names & scores from file
            For i = 0 To 4
                Top5PlayerName(i) = objStreamReader.ReadLine()
                Top5PlayerScore(i) = objStreamReader.ReadLine()
            Next

            ' Close the stream  
            objStreamReader.Close()

        Catch

            ':: Load defaults, file dosen't exist ::

            For i = 0 To 4
                Top5PlayerName(i) = "Pacman Champ"
            Next
            For i = 0 To 4
                Top5PlayerScore(i) = (95000 - (i * 15000))
            Next

        End Try

    End Sub

    Private Sub SaveHighScores()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Save high scores and player's names to disk
        '------------------------------------------------------------------------------------------------------------------

        Dim FILENAME As String = (CurDir() & "\high scores.txt")

        ' Get a StreamReader class that can be used to read the file  
        Dim objFileStream As FileStream
        objFileStream = File.Create(FILENAME)

        Dim encoder As New System.Text.ASCIIEncoding()
        Dim Str As String
        Dim buffer() As Byte
        Dim i As Integer

        ' Write high scores to file
        For i = 0 To 4

            Str = Top5PlayerName(i)
            Str += vbNewLine

            ReDim buffer(Str.Length - 1)
            encoder.GetBytes(Str, 0, Str.Length, buffer, 0)
            objFileStream.Write(buffer, 0, buffer.Length)

            Str = Top5PlayerScore(i)
            Str += vbNewLine

            ReDim buffer(Str.Length - 1)
            encoder.GetBytes(Str, 0, Str.Length, buffer, 0)
            objFileStream.Write(buffer, 0, buffer.Length)

        Next

        ' Close the stream  
        objFileStream.Close()

    End Sub

    Private Sub ShowScore()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Render player's score 
        '------------------------------------------------------------------------------------------------------------------

        Dim i As Integer

        For i = 0 To 1
            DrawText(GraphicsBuffer, "SCORE: " & Score, 275 + i, 10, 14, FontStyle.Regular, Brushes.White)
        Next

    End Sub

    Private Sub ShowHighScore()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Render highest score 
        '------------------------------------------------------------------------------------------------------------------

        Dim i As Integer

        For i = 0 To 1
            DrawText(GraphicsBuffer, "HIGH SCORE: " & Top5PlayerScore(0), 40 + i, 10, 14, FontStyle.Regular, Brushes.Red)
        Next

    End Sub

    Private Sub DrawText(ByVal Destination As Graphics, ByVal Text As String, ByVal x As Integer, ByVal y As Integer, ByVal FontSize As Integer, ByVal Style As FontStyle, ByVal Col As Brush)
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Draw text to graphics buffer using specified attributes and text string
        '------------------------------------------------------------------------------------------------------------------

        ' Font and specified attributes
        Dim DrawFont As New Font("MS Sans Serif", FontSize, Style)
        Dim DrawFormat As New StringFormat()
        DrawFormat.FormatFlags = StringFormatFlags.NoFontFallback

        ' Draw the string to graphics buffer
        Destination.DrawString(Text, DrawFont, Col, x, y, DrawFormat)

    End Sub

    Private Sub ScanKeyboard()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Fetch keyboard strokes from user ...
        '------------------------------------------------------------------------------------------------------------------

        If GetAsyncKeyState(Keys.Right) Then
            MoveRight = True
        ElseIf GetAsyncKeyState(Keys.Left) Then
            MoveLeft = True
        ElseIf GetAsyncKeyState(Keys.Up) Then
            MoveUP = True
        ElseIf GetAsyncKeyState(Keys.Down) Then
            MoveDown = True

        ElseIf GetAsyncKeyState(Keys.F2) Then
            If Not DoIntro Then
                If Not ShowGameOver Then
                    NewGame = True
                    GameOver = True
                End If
            End If
        End If

        AppRunning = GetAsyncKeyState(Keys.Escape)
        AppRunning = Not AppRunning

    End Sub

    Private Sub InitDelay()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Set 1uS resolution for timegettime  
        '------------------------------------------------------------------------------------------------------------------

        timeBeginPeriod(1)

    End Sub

    Private Sub RegulatedDelay(ByVal mS As Integer)
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Provides a precision, regulated delay: 1mS resolution
        '------------------------------------------------------------------------------------------------------------------

        Dim StartmS As Integer
        Dim ElapsedmS As Integer

        StartmS = timeGetTime

        If ProcessStartmS <> 0 Then
            mS = mS - (StartmS - ProcessStartmS)
            Do While (ElapsedmS < mS)
                ElapsedmS = (timeGetTime - StartmS)
                Application.DoEvents()
            Loop
        End If

        ProcessStartmS = timeGetTime()

    End Sub

    Private Sub WipeScore()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Completely erases score area top of screen ...
        '------------------------------------------------------------------------------------------------------------------

        GraphicsBuffer.FillRectangle(New SolidBrush(Color.Black), 0, 0, 500, 50)

    End Sub

    Private Sub Wipelives()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Completely erases lower area of screen constaing lives remaining
        '------------------------------------------------------------------------------------------------------------------

        GraphicsBuffer.FillRectangle(New SolidBrush(Color.Black), 0, 550, 500, 50)

    End Sub

    Private Sub WipeGraphicsBuffer()
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Completely erases primary graphics buffer
        '------------------------------------------------------------------------------------------------------------------

        GraphicsBuffer.FillRectangle(New SolidBrush(Color.Black), 0, 0, 500, 600)

    End Sub

    Private Sub form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Render graphics buffer to the form
        '------------------------------------------------------------------------------------------------------------------

        e.Graphics.DrawImage(ScrnBufferBmp, 0, 0)

    End Sub

    Private Sub frmGame_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Prepare to terminate app, save high scores and config files and restore screen res if changed
        '------------------------------------------------------------------------------------------------------------------

        If AppRunning Then
            AppRunning = False
            e.Cancel = True
            Call SaveHighScores()
        End If

    End Sub

    Private Sub frmGame_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '------------------------------------------------------------------------------------------------------------------
        ' Purpose: Terminate application -- clean up, dispose of resources
        '------------------------------------------------------------------------------------------------------------------

        ScrnBufferBmp.Dispose()
        GraphicsBuffer.Dispose()
        Me.Dispose()
        End

    End Sub
End Class
