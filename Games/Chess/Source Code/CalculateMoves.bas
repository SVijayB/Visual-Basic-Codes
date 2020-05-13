Attribute VB_Name = "CalculateMoves"

Public Function FilterMoves(UnitID As Integer, Scan As Boolean, Row As Integer, Column As Integer) As Integer
Dim StoreOriginalRow As Integer
Dim StoreOriginalColumn As Integer
Dim BreakSwitch As Boolean
Dim CounterX1 As Integer
Dim CounterX2 As Integer
Dim unitType As String

Erase RawLegalMoves
ReDim RawLegalMoves(1 To 2, 1 To 8) As Integer
StoreOriginalRow = Row
StoreOriginalColumn = Column
BreakSwitch = False
CounterX1 = 0
CounterX2 = 1

If Scan = False Then


CounterX1 = 0

Do Until CounterX1 = 8

Let CounterX1 = CounterX1 + 1

' Filter Out squares that do not exist on the board
If Overflow(Row + UnitGeneralMoves(1, CounterX1), Column + UnitGeneralMoves(2, CounterX1)) = True Then GoTo Skip

' FIlter out Squares occupied by Ally units
If ScoutPlace(UnitID, Row + UnitGeneralMoves(1, CounterX1), Column + UnitGeneralMoves(2, CounterX1)) = "Ally Occupied" Then GoTo Skip
    
    RawLegalMoves(1, CounterX1) = Row + UnitGeneralMoves(1, CounterX1)
    RawLegalMoves(2, CounterX1) = Column + UnitGeneralMoves(2, CounterX1)


Skip:


Loop





ElseIf Scan = True Then



Do

    ' Stop Scanning in one direction when reached end of board.
    

    If Overflow(Row + UnitGeneralMoves(1, CounterX2), Column + UnitGeneralMoves(2, CounterX2)) = True Then
        Let BreakSwitch = True
        GoTo ChangeDirection
    End If
    

    ' If and When Enemy unit blocks path, then Stop Scanning at the Square occupied by the enemy.
    
    
    
    If ScoutPlace(UnitID, Row + UnitGeneralMoves(1, CounterX2), Column + UnitGeneralMoves(2, CounterX2)) = "Enemy Occupied" Then
        
        Let BreakSwitch = True
       
    End If
    
    
    ' If and When Ally unit blocks path, then Stop Scanning at "one Square before" the ally uit.
    
    If ScoutPlace(UnitID, Row + UnitGeneralMoves(1, CounterX2), Column + UnitGeneralMoves(2, CounterX2)) = "Ally Occupied" Then
        Let BreakSwitch = True
        GoTo ChangeDirection
    End If
    
    
   
    
   ' Scan Process
        
    
   CounterX1 = CounterX1 + 1
        
    Let Row = Row + UnitGeneralMoves(1, CounterX2)
    Let Column = Column + UnitGeneralMoves(2, CounterX2)
        
    ReDim Preserve RawLegalMoves(1 To 2, 1 To CounterX1) As Integer
        
    RawLegalMoves(1, CounterX1) = Row
    RawLegalMoves(2, CounterX1) = Column

    
    
    
ChangeDirection: ' Shift Scanning Direction
       
  If BreakSwitch = True Then
  
       
    If CounterX2 < UBound(UnitGeneralMoves, 2) Then
        Let CounterX2 = CounterX2 + 1
        Let Row = StoreOriginalRow
        Let Column = StoreOriginalColumn
        Let BreakSwitch = False
    
    Else
        
        Exit Do
    End If
    End If
    
    
    
    Loop


Let Column = StoreOriginalColumn
Let Row = StoreOriginalRow



End If


End Function

Public Function AssignMoves(UnitID As Integer, Row As Integer, Column As Integer) As Integer

'--------------------------
' This function assigns "Possible Moves" to the selected Chess Piece
'--------------------------


Dim StoreOriginalRow As Integer
Dim StoreOriginalColumn As Integer
Dim BreakSwitch As Boolean
Dim CounterX1 As Integer
Dim CounterX2 As Integer
'Dim unitType As String

Erase RawLegalMoves ' Erase Old Array ... otherwise its possible that old Data will cause glitches.

unitType = Info(UnitID, 2)

Select Case unitType

'------------------------------------------------------------------------
Case "Pawn"
'------------------------------------------------------------------------
ReDim RawLegalMoves(1 To 2, 1 To 4) As Integer

ReDim UnitGeneralMoves(1 To 2, 1 To 8) As Integer



If Turn = True Then
    
    
        If UnitLocations(1, UnitID) = 2 And (ScoutPlace(UnitID, Row + 2, Column) = "Empty") And (ScoutPlace(UnitID, Row + 1, Column) = "Empty") Then
        
            UnitGeneralMoves(1, 1) = 2
            UnitGeneralMoves(2, 1) = 0
                
        End If
        
        If ScoutPlace(UnitID, Row + 1, Column) = "Empty" Then
                        
            UnitGeneralMoves(1, 2) = 1
            UnitGeneralMoves(2, 2) = 0
            
        End If
                        
        If ScoutPlace(UnitID, Row + 1, Column + 1) = "Enemy Occupied" Then
                        
            UnitGeneralMoves(1, 3) = 1
            UnitGeneralMoves(2, 3) = 1
            
        End If
                    
        If ScoutPlace(UnitID, Row + 1, Column - 1) = "Enemy Occupied" Then
                        
            UnitGeneralMoves(1, 4) = 1
            UnitGeneralMoves(2, 4) = -1
            
        End If
            
    ElseIf Turn = False Then
    
    
    If UnitLocations(1, UnitID) = 7 And (ScoutPlace(UnitID, Row - 2, Column) = "Empty") And (ScoutPlace(UnitID, Row - 1, Column) = "Empty") Then
        
            UnitGeneralMoves(1, 1) = -2
            UnitGeneralMoves(2, 1) = 0
                
        End If
        
        If ScoutPlace(UnitID, Row - 1, Column) = "Empty" Then
                        
            UnitGeneralMoves(1, 2) = -1
            UnitGeneralMoves(2, 2) = 0
            
        End If
                        
        If ScoutPlace(UnitID, Row - 1, Column + 1) = "Enemy Occupied" Then
                        
            UnitGeneralMoves(1, 3) = -1
            UnitGeneralMoves(2, 3) = 1
            
        End If
                    
        If ScoutPlace(UnitID, Row - 1, Column - 1) = "Enemy Occupied" Then
                        
            UnitGeneralMoves(1, 4) = -1
            UnitGeneralMoves(2, 4) = -1
            
        End If
    

End If

Call FilterMoves(UnitID, False, Row, Column)



'--------------------------------------------------------------------------
Case "King"                      ' King
'--------------------------------------------------------------------------

ReDim RawLegalMoves(1 To 2, 1 To 8) As Integer

ReDim UnitGeneralMoves(1 To 2, 1 To 8) As Integer


                                ' Row + 1 and Column
UnitGeneralMoves(1, 1) = 1
UnitGeneralMoves(2, 1) = 0

                                ' Row - 1 and Column
UnitGeneralMoves(1, 2) = -1
UnitGeneralMoves(2, 2) = 0

                                ' Row and Column + 1
UnitGeneralMoves(1, 3) = 0
UnitGeneralMoves(2, 3) = 1

                                ' Row and Column - 1
UnitGeneralMoves(1, 4) = 0
UnitGeneralMoves(2, 4) = -1
                                
                                ' Row + 1 and Column + 1
UnitGeneralMoves(1, 5) = 1
UnitGeneralMoves(2, 5) = 1

                                ' Row - 1 and Column - 1
UnitGeneralMoves(1, 6) = -1
UnitGeneralMoves(2, 6) = -1

                                ' Row + 1 and Column - 1
UnitGeneralMoves(1, 7) = 1
UnitGeneralMoves(2, 7) = -1

                                ' Row - 1 and Column + 1
UnitGeneralMoves(1, 8) = -1
UnitGeneralMoves(2, 8) = 1


Call FilterMoves(UnitID, False, Row, Column)


'--------------------------------------------------------------------------

Case "Knight"      ' Knight

'--------------------------------------------------------------------------


ReDim RawLegalMoves(1 To 2, 1 To 8) As Integer

ReDim UnitGeneralMoves(1 To 2, 1 To 8) As Integer


                                ' Row + 2 and Column +1
UnitGeneralMoves(1, 1) = 2
UnitGeneralMoves(2, 1) = 1

                                ' Row + 1 and Column +2
UnitGeneralMoves(1, 2) = 1
UnitGeneralMoves(2, 2) = 2

                                ' Row -1 and Column + 2
UnitGeneralMoves(1, 3) = -1
UnitGeneralMoves(2, 3) = 2

                                ' Row-2 and Column +1
UnitGeneralMoves(1, 4) = -2
UnitGeneralMoves(2, 4) = 1
                                
                                ' Row -2 and Column - 1
UnitGeneralMoves(1, 5) = -2
UnitGeneralMoves(2, 5) = -1

                                ' Row - 1 and Column - 2
UnitGeneralMoves(1, 6) = -1
UnitGeneralMoves(2, 6) = -2

                                ' Row + 1 and Column - 2
UnitGeneralMoves(1, 7) = 1
UnitGeneralMoves(2, 7) = -2

                                ' Row + 2 and Column - 1
UnitGeneralMoves(1, 8) = 2
UnitGeneralMoves(2, 8) = -1

Call FilterMoves(UnitID, False, Row, Column)

'--------------------------------------------------------------------------

Case "Rook"      ' Rook \ Castle

'--------------------------------------------------------------------------


ReDim UnitGeneralMoves(1 To 2, 1 To 4) As Integer


                                ' Row + 1 and Column
UnitGeneralMoves(1, 1) = 1
UnitGeneralMoves(2, 1) = 0

                                ' Row - 1 and Column
UnitGeneralMoves(1, 2) = -1
UnitGeneralMoves(2, 2) = 0

                                ' Row and Column + 1
UnitGeneralMoves(1, 3) = 0
UnitGeneralMoves(2, 3) = 1

                                ' Row and Column  - 1
UnitGeneralMoves(1, 4) = 0
UnitGeneralMoves(2, 4) = -1
                             

Call FilterMoves(UnitID, True, Row, Column)


'------------------------------------

'--------------------------------------------------------------------------
Case "Bishop"  ' Bishop
'--------------------------------------------------------------------------


ReDim UnitGeneralMoves(1 To 2, 1 To 4) As Integer


                                ' Row + 1 and Column + 1
UnitGeneralMoves(1, 1) = 1
UnitGeneralMoves(2, 1) = 1

                                ' Row - 1 and Column - 1
UnitGeneralMoves(1, 2) = -1
UnitGeneralMoves(2, 2) = -1

                                ' Row -1 and Column + 1
UnitGeneralMoves(1, 3) = -1
UnitGeneralMoves(2, 3) = 1

                                ' Row +1 and Column  - 1
UnitGeneralMoves(1, 4) = 1
UnitGeneralMoves(2, 4) = -1
                             


Call FilterMoves(UnitID, True, Row, Column)


'--------------------------------------------------------------------------
Case "Queen" ' Queen
'--------------------------------------------------------------------------
                   

ReDim UnitGeneralMoves(1 To 2, 1 To 8) As Integer
                       
                               ' Row + 1 and Column +0
UnitGeneralMoves(1, 1) = 1
UnitGeneralMoves(2, 1) = 0

                                ' Row - 1 and Column +0
UnitGeneralMoves(1, 2) = -1
UnitGeneralMoves(2, 2) = 0

                                ' Row + 0 and Column + 1
UnitGeneralMoves(1, 3) = 0
UnitGeneralMoves(2, 3) = 1

                                ' Row + 0 and Column  - 1
UnitGeneralMoves(1, 4) = 0
UnitGeneralMoves(2, 4) = -1
                             

                                ' Row + 1 and Column +1
UnitGeneralMoves(1, 5) = 1
UnitGeneralMoves(2, 5) = 1

                                ' Row - 1 and Column -1
UnitGeneralMoves(1, 6) = -1
UnitGeneralMoves(2, 6) = -1

                                ' Row -1 and Column + 1
UnitGeneralMoves(1, 7) = -1
UnitGeneralMoves(2, 7) = 1

                                ' Row +1 and Column  - 1
UnitGeneralMoves(1, 8) = 1
UnitGeneralMoves(2, 8) = -1
                             
Call FilterMoves(UnitID, True, Row, Column)

End Select


'---------------------------------------------------------------



' Filter Raw-Legal-Moves Data




CounterX1 = 0
CounterX2 = 0

Erase RefinedLegalMoves



Do Until CounterX1 = UBound(RawLegalMoves, 2)

CounterX1 = CounterX1 + 1

If Not Overflow(RawLegalMoves(1, CounterX1), RawLegalMoves(2, CounterX1)) = True Then

Let CounterX2 = CounterX2 + 1

ReDim Preserve RefinedLegalMoves(1 To 2, 1 To CounterX2) As Integer
RefinedLegalMoves(1, CounterX2) = RawLegalMoves(1, CounterX1)
RefinedLegalMoves(2, CounterX2) = RawLegalMoves(2, CounterX1)

End If




Loop

If CounterX2 = 0 Then Let AssignMoves = 300
If CounterX2 > 0 Then Let AssignMoves = 100




End Function

Public Function XMoveUnit(UnitID As Integer, Row As Integer, Column As Integer, Force As Boolean) As Integer

'-----------------------------------------
' This Function Moves the Graphical Units on the Board when Instructed.
'-----------------------------------------


If Not Force = True Then
    If Not (frmChessMain.Grid(GridIndex).BackColor = RGB(0, 128, 192) Or (frmChessMain.Grid(GridIndex).BackColor = RGB(190, 57, 1))) Then GoTo InvalidMove2
End If

       

' Remove Enemy Unit if Space Occupied
        
        If Not Board(Row, Column) = 0 Then
        
        frmChessMain.BlackxPieces(Board(Row, Column)).Enabled = False
        
        
        frmChessMain.BlackxPieces(Board(Row, Column)).Left = PrisonerLocations(1, Board(Row, Column))
        frmChessMain.BlackxPieces(Board(Row, Column)).Top = PrisonerLocations(2, Board(Row, Column))
        If Turn = True Then
            frmChessMain.BlackxPieces(Board(Row, Column)).BackColor = frmChessMain.picWhtPrison.BackColor
        ElseIf Turn = False Then
                frmChessMain.BlackxPieces(Board(Row, Column)).BackColor = frmChessMain.picBlckPrison.BackColor
        End If
        
        
        UnitLocations(1, Board(Row, Column)) = 0
        UnitLocations(2, Board(Row, Column)) = 0
        
        End If
      
    

' Clear the Old Position in all Positioning Arrays.

        Board(UnitLocations(1, UnitID), UnitLocations(2, UnitID)) = 0
        UnitLocations(1, UnitID) = 0
        UnitLocations(2, UnitID) = 0


' Record its new position in All Arrays

        Board(Row, Column) = UnitID
        UnitLocations(1, UnitID) = Row
        UnitLocations(2, UnitID) = Column
        

   ' Actually move the unit art to its new Position.

        frmChessMain.BlackxPieces(UnitID).Left = frmChessMain.Grid(gridTolinear(Row, Column)).Left + 240
        frmChessMain.BlackxPieces(UnitID).Top = frmChessMain.Grid(gridTolinear(Row, Column)).Top + 120
        frmChessMain.BlackxPieces(UnitID).BackColor = frmChessMain.Grid(gridTolinear(Row, Column)).BackColor
        frmChessMain.BlackxPieces(UnitID).BorderStyle = 0
        
    
'Call PawnPromotionCheck(UnitID)
Call PawnPromotionCheck(UnitID)



        
' Clear the Unit
        Call ClearHilighting
        PieceIndex = 0
        






        
    
' Change Turns
    
    If Not Force = True Then
            
        If Turn = True Then
            Let Turn = False
            frmChessMain.TopTimer.Enabled = True
            frmChessMain.BottomTimer.Enabled = False
            frmChessMain.imgBackground.Picture = LoadPicture(App.Path + "\GameArt\Table\BoardTurn02.jpg")
            
            If WhiteAI = True Then
                frmChessMain.tmrAI.Enabled = True
            End If
            
        ElseIf Turn = False Then
            
            Let Turn = True
            frmChessMain.TopTimer.Enabled = False
            frmChessMain.BottomTimer.Enabled = True
            frmChessMain.imgBackground.Picture = LoadPicture(App.Path + "\GameArt\Table\BoardTurn01.jpg")
            
            If BlackAI = True Then
                frmChessMain.tmrAI.Enabled = True
            End If
            
            End If
        
    End If
    

If Not Force = True Then
    Call MateCheck(1) ' Check for Checkmate
End If

    
GoTo Finish

InvalidMove2:
        MsgBox "This unit can not move to this position.", vbOKOnly, "Occupied"
        GoTo Finish
        
InvalidMove:
        MsgBox "This Square is already Occupied by Your own Unit.", vbOKOnly, "Occupied"
        GoTo Finish
        
Finish:


End Function


Public Function HilightMoves(UnitID As Integer, Row As Integer, Column As Integer) As Integer

'---------------------------------------------------------------------
' This function Hilights Squares on the Board, depending on where a selected unit can Move to.
'---------------------------------------------------------------------

Dim CounterX1 As Integer


Let Berty = AssignMoves(UnitID, Row, Column)


If Berty = 300 Then
    Let NoMoves = True
    GoTo skipper
End If

If Berty = 100 Then
    Let NoMoves = False
End If


CounterX1 = 0


Do Until CounterX1 = UBound(RefinedLegalMoves, 2)
CounterX1 = CounterX1 + 1
        If Not Overflow(RefinedLegalMoves(1, CounterX1), RefinedLegalMoves(2, CounterX1)) = True Then
            If ScoutPlace(UnitID, RefinedLegalMoves(1, CounterX1), RefinedLegalMoves(2, CounterX1)) = "Empty" Then
                frmChessMain.Grid(gridTolinear(RefinedLegalMoves(1, CounterX1), RefinedLegalMoves(2, CounterX1))).BackColor = RGB(0, 128, 192)
            ElseIf ScoutPlace(UnitID, RefinedLegalMoves(1, CounterX1), RefinedLegalMoves(2, CounterX1)) = "Enemy Occupied" Then
                frmChessMain.Grid(gridTolinear(RefinedLegalMoves(1, CounterX1), RefinedLegalMoves(2, CounterX1))).BackColor = RGB(190, 57, 1)
                frmChessMain.BlackxPieces(Board(RefinedLegalMoves(1, CounterX1), RefinedLegalMoves(2, CounterX1))).BackColor = RGB(190, 57, 1)
            End If
        End If
Loop


skipper:
End Function


Public Function ClearHilighting()

'---------------------------------
' This function clears any old "Hilighting" on squares when a new piece is selected
'---------------------------------

Dim CounterX10 As Integer

CounterX10 = 0

Do Until CounterX10 = 64
CounterX10 = CounterX10 + 1

If GetRow(CounterX10) Mod 2 = 0 Then

    If CounterX10 Mod 2 = 0 Then
        frmChessMain.Grid(CounterX10).BackColor = &H404040
    Else
        frmChessMain.Grid(CounterX10).BackColor = &HE0E0E0
    
    End If
    

Else

    If CounterX10 Mod 2 = 0 Then
        frmChessMain.Grid(CounterX10).BackColor = &HE0E0E0
    Else
        frmChessMain.Grid(CounterX10).BackColor = &H404040


    End If


End If


Loop

Dim Counter01 As Integer


Counter01 = 0

Do Until Counter01 = 32
Let Counter01 = Counter01 + 1
If Not UnitLocations(1, Counter01) = 0 Then
If Not frmChessMain.BlackxPieces(Counter01).BackColor = frmChessMain.Grid(gridTolinear(UnitLocations(1, Counter01), UnitLocations(2, Counter01))).BackColor Then
frmChessMain.BlackxPieces(Counter01).BackColor = frmChessMain.Grid(gridTolinear(UnitLocations(1, Counter01), UnitLocations(2, Counter01))).BackColor
End If
End If

Loop



Let ClearHilighting = 1

End Function


