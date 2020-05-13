Attribute VB_Name = "PawnPromotion"
Public UnitStorage As Byte



Public Function PawnPromotionCheck(UnitID As Integer)


If Info(UnitID, 2) = "Pawn" Then

    If Turn = True And MyRow(UnitID) = 8 And WhiteAI = False Then
        Call SetupPawnPromoGUI
        frmPawnPromotion.Show
        Let UnitStorage = UnitID
    
    ElseIf Turn = True And MyRow(UnitID) = 8 And WhiteAI = True Then
        Call PromotePawn("Queen", UnitID)
        

    ElseIf Turn = False And MyRow(UnitID) = 1 And BlackAI = False Then
        Call SetupPawnPromoGUI
        frmPawnPromotion.Show
        Let UnitStorage = UnitID
        
    ElseIf Turn = False And MyRow(UnitID) = 1 And BlackAI = True Then
        Call PromotePawn("Queen", UnitID)
        
    End If
    

End If

End Function



Public Function PromotePawn(ChosenOne As String, Optional PassedUnitID As Integer)

Dim UnitID As Integer


If PassedUnitID = 0 Then
    Let UnitID = UnitStorage
Else
    Let UnitID = PassedUnitID
End If

Dim playerX As String

playerX = Info(UnitID, 1)



PromotedPawns(PawnNumber(UnitID), Val(Info(UnitID, 0))) = ChosenOne
frmChessMain.BlackxPieces(UnitID).Picture = LoadPicture(App.Path + "\GameArt\Units\" + playerX + "x" + PromotedPawns(PawnNumber(UnitID), Info(UnitID, 0)) + ".gif")



End Function

Public Function SetupPawnPromoGUI()

Dim playerX As String

If Turn = False Then Let playerX = "white"
If Turn = True Then Let playerX = "black"

        frmPawnPromotion.ImgPromo(1).Picture = LoadPicture(App.Path + "\GameArt\Units\" + playerX + "xQueen.gif")
        frmPawnPromotion.ImgPromo(2).Picture = LoadPicture(App.Path + "\GameArt\Units\" + playerX + "xRook.gif")
        frmPawnPromotion.ImgPromo(3).Picture = LoadPicture(App.Path + "\GameArt\Units\" + playerX + "xBishop.gif")
        frmPawnPromotion.ImgPromo(4).Picture = LoadPicture(App.Path + "\GameArt\Units\" + playerX + "xKnight.gif")
        
End Function

Public Function PawnNumber(UnitID As Integer) As Integer


If Info(UnitID, 1) = "Black" Then

    Let PawnNumber = UnitID - 8
 

ElseIf Info(UnitID, 1) = "White" Then

    PawnNumber = UnitID - 16
    
End If


End Function

Public Function GetPawnProperID(PawnREF As Integer) As Integer

GetPawnProperID = PawnREF + 8

End Function

