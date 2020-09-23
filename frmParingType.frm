VERSION 5.00
Begin VB.Form frmParingType 
   Caption         =   "Paring type"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   2670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Type 2"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Type 1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Give the type of Paring U want"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmParingType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ParingType = 0
    Me.Hide
End Sub

Private Sub Command2_Click()
    ParingType = 1
    Me.Hide
End Sub
