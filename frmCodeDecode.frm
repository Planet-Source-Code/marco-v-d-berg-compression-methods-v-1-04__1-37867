VERSION 5.00
Begin VB.Form frmCodeDecode 
   Caption         =   "Code/Decode"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   1965
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton btnDecode 
      Caption         =   "Decode"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnCode 
      Caption         =   "Code"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCodeDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    CompDecomp = 0
    Me.Hide
End Sub

Private Sub btnCode_Click()
    CompDecomp = 1
    Me.Hide
End Sub

Private Sub btnDecode_Click()
    CompDecomp = 2
    Me.Hide
End Sub

Private Sub Form_Load()
    CompDecomp = 0
End Sub
