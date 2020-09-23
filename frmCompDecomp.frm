VERSION 5.00
Begin VB.Form frmCompDecomp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "De/Comp"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   1995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton btnDeCompress 
      Caption         =   "Decompress"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnCompress 
      Caption         =   "Compress"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCompDecomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    CompDecomp = 0
    Me.Hide
End Sub

Private Sub btnCompress_Click()
    CompDecomp = 1
    Me.Hide
End Sub

Private Sub btnDeCompress_Click()
    CompDecomp = 2
    Me.Hide
End Sub

Private Sub Form_Load()
    CompDecomp = 0
End Sub
