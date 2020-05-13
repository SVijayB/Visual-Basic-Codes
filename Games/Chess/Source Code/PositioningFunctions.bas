Attribute VB_Name = "PositioningFunctions"

Public Function GetRow(Linear As Integer) As Integer

Dim Counter01 As Integer
Dim BreakSwitch As Boolean

BreakSwitch = False
Counter01 = 1

Do Until BreakSwitch = True

If Linear <= 8 * Counter01 Then
    Let GetRow = Counter01
    Let BreakSwitch = True
Else
Counter01 = Counter01 + 1
End If

Loop

End Function

Public Function GetColumn(Linear As Integer) As Integer

GetColumn = 8 - ((8 * GetRow(Linear)) - Linear)

End Function


Public Function gridTolinear(ROWnumber As Integer, COLnumber As Integer) As Integer


gridTolinear = 8 * (ROWnumber - 1) + COLnumber

End Function

Public Function Overflow(RowModified As Integer, ColumnModified As Integer) As Boolean


If (RowModified >= 1 And RowModified <= 8 And ColumnModified <= 8 And ColumnModified >= 1) Then
    Let Overflow = False
    
ElseIf (RowModified < 1 Or RowModified > 8 Or ColumnModified > 8 Or ColumnModified < 1) Then
    Let Overflow = True
    
Else
   
    
End If
        
End Function

Public Function ScoutPlace(UnitID As Integer, Row As Integer, Column As Integer) As String

If Overflow(Row, Column) = False Then
    
    If UnitID >= 17 Then
        If Board(Row, Column) >= 17 Then Let ScoutPlace = "Ally Occupied"
        If Board(Row, Column) < 17 And Not Board(Row, Column) = 0 Then Let ScoutPlace = "Enemy Occupied"
        If Board(Row, Column) = 0 Then Let ScoutPlace = "Empty"
    
    ElseIf UnitID < 17 Then
        
        If Board(Row, Column) >= 17 Then Let ScoutPlace = "Enemy Occupied"
        If Board(Row, Column) < 17 And Not Board(Row, Column) = 0 Then Let ScoutPlace = "Ally Occupied"
        If Board(Row, Column) = 0 Then Let ScoutPlace = "Empty"
    
    End If



End If
  

End Function

Public Function MyRow(UnitID As Integer) As Integer

MyRow = UnitLocations(1, UnitID)

End Function

Public Function MyColumn(UnitID As Integer) As Integer
MyColumn = UnitLocations(2, UnitID)

End Function

Public Function MyLinearPosition(UnitID As Integer) As Integer
MyLinearPosition = gridTolinear(MyRow(UnitID), MyColumn(UnitID))


End Function


Public Function Info(UnitID As Integer, InfoType As Byte) As String

Dim TempVar As String
Dim SubS As Byte

Select Case InfoType


Case 0

If UnitID <= 16 Then

            Info = 1
    
    ElseIf UnitID > 16 Then

            Info = 2

End If



Case 1


    If UnitID <= 16 Then

            Info = "Black"
    
    ElseIf UnitID > 16 Then

            Info = "White"

    End If



Case 2

    If (UnitID >= 9 And UnitID <= 16) Or (UnitID >= 17 And UnitID <= 24) Then
    
    
        If UnitID <= 16 Then SubS = 1
        If UnitID > 16 Then SubS = 2
    
        If PromotedPawns(PawnNumber(UnitID), SubS) = "Pawn" Then
    
                 Let Info = "Pawn"
             
        Else
    
                Info = PromotedPawns(PawnNumber(UnitID), SubS)
    
        End If
    
    
    ElseIf UnitID = 29 Or UnitID = 5 Then

                Let Info = "King"
    
    ElseIf UnitID = 26 Or UnitID = 31 Or UnitID = 2 Or UnitID = 7 Then

             Let Info = "Knight"

    ElseIf UnitID = 1 Or UnitID = 8 Or UnitID = 25 Or UnitID = 32 Then

             Let Info = "Rook"

    ElseIf UnitID = 27 Or UnitID = 30 Or UnitID = 3 Or UnitID = 6 Then

            Let Info = "Bishop"

    ElseIf UnitID = 28 Or UnitID = 4 Then

            Let Info = "Queen"

    End If



Case 3


Let TempVar = Info(UnitID, 2)

If TempVar = "Pawn" Then Let Info = 1
If TempVar = "Knight" Then Let Info = 3
If TempVar = "Bishop" Then Let Info = 3
If TempVar = "Rook" Then Let Info = 5
If TempVar = "Queen" Then Let Info = 9


End Select




End Function

