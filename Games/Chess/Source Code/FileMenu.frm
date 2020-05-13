VERSION 5.00
Begin VB.Form frmFileMenu 
   AutoRedraw      =   -1  'True
   Caption         =   "File Menu"
   ClientHeight    =   4665
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7860
   Begin VB.CommandButton cmdDelete02 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveLoad02 
      Caption         =   "Save"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5175
   End
   Begin VB.FileListBox flstFileList 
      Height          =   3210
      Left            =   120
      Pattern         =   "*.chss"
      TabIndex        =   0
      Top             =   1200
      Width           =   7575
   End
End
Attribute VB_Name = "frmFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete02_Click()

If flstFileList.FileName = "" Then

MsgBox "Please select a file to Delete"

Else

Kill (flstFileList.Path + "\" + flstFileList.FileName)
flstFileList.Refresh

End If

End Sub

Private Sub cmdSaveLoad02_Click()

If cmdSaveLoad02.Caption = "Load" Then
If flstFileList.FileName = "" Then

MsgBox "Please select a file to load"

Else

    Call LoadGame2(flstFileList.FileName)
    frmFileMenu.Hide
    frmChessMain.SetFocus
End If

ElseIf cmdSaveLoad02.Caption = "Save" Then
If txtFileName.Text = "" Then

MsgBox "Please enter a File Name for the save."

Else

    Call SaveGame(txtFileName.Text)
    frmFileMenu.Hide
    frmChessMain.SetFocus
    
End If

End If


End Sub

Private Sub flstFileList_Click()

txtFileName.Text = flstFileList.FileName

End Sub


Private Sub Form_Load()


flstFileList.Path = App.Path + "\SaveGames\"
flstFileList.Refresh

    
End Sub
