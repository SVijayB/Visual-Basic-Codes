Attribute VB_Name = "WinCheck"

Public Function GameOver(Winner As String, WinType As String)


' Disable AI
frmChessMain.tmrAI.Enabled = False


' Disable all Units

For i = 1 To 32
        frmChessMain.BlackxPieces(i).Enabled = False
Next


MsgBox WinType + " " + Winner + " " + "Wins", vbOKOnly, WinType


End Function



Public Function MateCheck(Mode As Byte)

Dim KingID As Byte
Dim c As Integer
c = 0

If Mode = 1 Then

    If Turn = True Then KingID = 29
    If Turn = False Then KingID = 5

ElseIf Mode = 2 Then

    If Turn = False Then KingID = 29
    If Turn = True Then KingID = 5

End If



Call CalculateAIMoves(Turn)

Do Until c = UBound(AIMoves, 2)
c = c + 1

If UnitLocations(1, KingID) = AIMoves(2, c) And UnitLocations(2, KingID) = AIMoves(3, c) Then
    If Mode = 1 Then
    
        Call ApplyCheckMate(c)
        Exit Do
        
    ElseIf Mode = 2 Then
    
        Let MateCheck = True
        Exit Do
        
    End If
    
End If


Loop


 




End Function

Public Function ApplyCheckMate(UnitID As Integer)


' Hilight Checkmate Situation

Call HilightMoves(AIMoves(1, UnitID), MyRow(AIMoves(1, UnitID)), MyColumn(AIMoves(1, UnitID)))
frmChessMain.BlackxPieces(AIMoves(1, UnitID)).BorderStyle = 1

If Turn = True Then

    Call GameOver("Black", "Checkmate")
    
    
ElseIf Turn = False Then
    
    Call GameOver("White", "Checkmate")
    
End If



End Function
