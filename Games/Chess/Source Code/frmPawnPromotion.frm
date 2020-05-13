VERSION 5.00
Begin VB.Form frmPawnPromotion 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   285
   ClientTop       =   4560
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6885
   Begin VB.CommandButton cmdPromotion 
      Caption         =   "Confirm"
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.OptionButton optPawnPromo 
      Caption         =   "Knight"
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.OptionButton optPawnPromo 
      Caption         =   "Bishop"
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.OptionButton optPawnPromo 
      Caption         =   "Rook"
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.OptionButton optPawnPromo 
      Caption         =   "Queen"
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image ImgPromo 
      Height          =   675
      Index           =   4
      Left            =   3600
      Top             =   4440
      Width           =   675
   End
   Begin VB.Image ImgPromo 
      Height          =   675
      Index           =   3
      Left            =   3600
      Top             =   3720
      Width           =   675
   End
   Begin VB.Image ImgPromo 
      Height          =   675
      Index           =   2
      Left            =   3600
      Top             =   3000
      Width           =   675
   End
   Begin VB.Image ImgPromo 
      Height          =   675
      Index           =   1
      Left            =   3600
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label lblPawnPromotion 
      Alignment       =   2  'Center
      Caption         =   "A pawn has been promoted: What do you want him to be converted into?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "frmPawnPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OptionChosen As Byte
Dim PromotedPawns(32) As String


Private Sub cmdPromotion_Click()

Dim Sry As String

If Turn = True Then Let Sry = "white"
If Turn = False Then Let Sry = "black"


Select Case OptionChosen

Case 0

MsgBox "Please choose which unit you want this Pawn to be replaced with"


Case 1

Call PromotePawn("Queen")
Me.Hide
Case 2

Call PromotePawn("Rook")
Me.Hide
Case 3

Call PromotePawn("Bishop")
Me.Hide
Case 4

Call PromotePawn("Knight")
Me.Hide
End Select

OptionChosen = 0


End Sub

Private Sub optPawnPromo_Click(Index As Integer)

OptionChosen = Index

End Sub






