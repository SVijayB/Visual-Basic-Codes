Attribute VB_Name = "SavingFunctions"

Public Function PrintArray(DisplayObject As PictureBox, ArrayName As Variant) As Integer

Dim CurrentRow As Integer
Dim CurrentColumn As Integer

Let CurrentRow = 1
Let CurrentColumn = 0

DisplayObject.Cls
DisplayObject.Print


Looper:


Do

Let CurrentColumn = CurrentColumn + 1
If CurrentColumn > UBound(ArrayName, 1) Then Exit Do
DisplayObject.Print "  "; ArrayName(CurrentColumn, CurrentRow);

Loop


If CurrentRow < UBound(ArrayName, 2) Then
    Let CurrentRow = CurrentRow + 1
    Let CurrentColumn = 0
    DisplayObject.Print
    GoTo Looper
End If



End Function


Public Function SaveGame(FileName As String) As Integer

Dim CurrentRow As Integer
Dim CurrentColumn As Integer

Let CurrentRow = 1
Let CurrentColumn = 0

Dim Str01 As String

  
If Not FileName = "" Then
If Right(FileName, 5) = ".chss" Then
    Str01 = App.Path + "\SaveGames\" + FileName
Else
    Str01 = App.Path + "\SaveGames\" + FileName + ".chss"
End If
Open Str01 For Output As #1

Print #1, "Chess & Checkers: Save File !!! 13020348 - 13020341 !!! "
Print #1, "TimeStamp:"; DateValue(Now); TimeValue(Now)

Looper:

Do

Let CurrentColumn = CurrentColumn + 1
If CurrentColumn > UBound(UnitLocations, 1) Then Exit Do
Print #1, UnitLocations(CurrentColumn, CurrentRow);

Loop

If CurrentRow < UBound(UnitLocations, 2) Then
    Let CurrentRow = CurrentRow + 1
    Let CurrentColumn = 0
    Print #1,
    GoTo Looper
End If

Print #1,

'----------------------
Let CurrentRow = 1
Let CurrentColumn = 0


Looper22:

Do

Let CurrentColumn = CurrentColumn + 1
If CurrentColumn > UBound(PromotedPawns, 1) Then Exit Do

If CurrentColumn = UBound(PromotedPawns, 1) Then
    Print #1, PromotedPawns(CurrentColumn, CurrentRow);
Else
        Print #1, PromotedPawns(CurrentColumn, CurrentRow) + ",";
End If


Loop

If CurrentRow < UBound(PromotedPawns, 2) Then
    Let CurrentRow = CurrentRow + 1
    Let CurrentColumn = 0
    Print #1,
    GoTo Looper22
End If

Print #1,

'-----------------------

Print #1, ((Val(frmChessMain.TopMIN.Text)) * 60 + Val(frmChessMain.TOPsec.Text))
Print #1, ((Val(frmChessMain.BottomMin.Text)) * 60 + Val(frmChessMain.BottomSec.Text))



Print #1, Turn


Close #1

End If


End Function


Public Function LoadGame2(Xfilename As String) As Integer

Dim Counter As Integer


Dim Time01 As Integer
Dim Time02 As Integer


Dim SomeVar As Integer
Dim Row As Integer
Dim Column As Integer

Dim Str01 As String

Let Str01 = "SaveGames\" + Xfilename

Erase UnitLocations
Erase Board

Let Column = 1

Open Str01 For Input As #1

Input #1, x
Input #1, x

Looper02:

Do

Let Row = Row + 1
If Row > UBound(UnitLocations, 1) Then Exit Do

Input #1, SomeVar

UnitLocations(Row, Column) = SomeVar

Loop

If Column < UBound(UnitLocations, 2) Then
    Let Column = Column + 1
    Let Row = 0
    GoTo Looper02
    
End If


For i = 1 To 16
    Input #1, Str01
        Call PromotePawn(Str01, GetPawnProperID(Val(i)))
Next



Input #1, Time01
Input #1, Time02

Timings(1, 1) = Int(Time01 / 60)
Timings(2, 1) = Time01 Mod 60

Timings(1, 2) = Int(Time02 / 60)
Timings(2, 2) = Time02 Mod 60



Input #1, Str01

If Str01 = "True" Then Let Turn = True
If Str01 = "False" Then Let Turn = False

    
Close #1

 If Turn = True Then
        Let StoreTurn = True
 Else
        Let StoreTurn = False
 End If
 
   
   Do
   
   Counter = Counter + 1
   If Counter = 33 Then Exit Do
    
    frmChessMain.BlackxPieces(Counter).Enabled = True
   
   If MyColumn(Counter) = 0 Or MyRow(Counter) = 0 Then
   
        frmChessMain.BlackxPieces(Counter).Enabled = False
        
        
        frmChessMain.BlackxPieces(Counter).Left = PrisonerLocations(1, Counter)
        frmChessMain.BlackxPieces(Counter).Top = PrisonerLocations(2, Counter)
        
        If Counter <= 16 Then
                
            frmChessMain.BlackxPieces(Counter).BackColor = frmChessMain.picBlckPrison.BackColor
        
        ElseIf Counter > 16 Then
        
            frmChessMain.BlackxPieces(Counter).BackColor = frmChessMain.picWhtPrison.BackColor
                
        End If
        
         
   
   Else
   
    Call XMoveUnit(Counter, MyRow(Counter), MyColumn(Counter), True)
   
   End If
   
   frmChessMain.TopMIN.Text = Timings(1, 1)
   frmChessMain.TOPsec.Text = Timings(2, 1)
   
   frmChessMain.BottomMin.Text = Timings(1, 2)
   frmChessMain.BottomSec.Text = Timings(2, 2)
   
   
   
   Loop
   
      
 If StoreTurn = False Then
            
            Let Turn = False
            frmChessMain.TopTimer.Enabled = True
            frmChessMain.BottomTimer.Enabled = False
            frmChessMain.imgBackground.Picture = LoadPicture(App.Path + "\GameArt\Table\BoardTurn02.jpg")
            
            If WhiteAI = True Then
                frmChessMain.tmrAI.Enabled = True
            End If
            
   ElseIf StoreTurn = True Then
         
            Let Turn = True
            frmChessMain.TopTimer.Enabled = False
            frmChessMain.BottomTimer.Enabled = True
            frmChessMain.imgBackground.Picture = LoadPicture(App.Path + "\GameArt\Table\BoardTurn01.jpg")
            
            If BlackAI = True Then
                frmChessMain.tmrAI.Enabled = True
            End If
End If
  
   
   


        


End Function






