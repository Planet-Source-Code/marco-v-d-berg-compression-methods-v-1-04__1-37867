VERSION 5.00
Begin VB.Form ChooseDictSize 
   Caption         =   "Choose dictionary size"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox CMBDictionarySize 
      Height          =   315
      ItemData        =   "ChooseDictSize.frx":0000
      Left            =   2280
      List            =   "ChooseDictSize.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Kb"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum size is"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the maximum dictionary size for this compressor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "ChooseDictSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    DictionarySize = Val(CMBDictionarySize.Text)
    Me.Hide
    DoEvents
End Sub

Private Sub CMBDictionarySize_Change()
    If Val(CMBDictionarySize.Text) < 1 Or Val(CMBDictionarySize.Text) > 64 Then
        MsgBox "This is not an item from the list"
        CMBDictionarySize.Text = Str(DictionarySize)
    End If
End Sub

Private Sub Form_Load()
    Dim X As Integer
    CMBDictionarySize.Clear
    For X = 1 To 64
        CMBDictionarySize.AddItem Str(X)
    Next
    If DictionarySize = 0 Then DictionarySize = 16
    CMBDictionarySize.Text = Str(DictionarySize)
End Sub
