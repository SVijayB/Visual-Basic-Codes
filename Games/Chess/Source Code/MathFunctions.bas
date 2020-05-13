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
