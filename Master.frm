VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Master 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Programm For Compressors V1.04"
   ClientHeight    =   7650
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnViewContentsTarget 
      Caption         =   "View contents"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton btnViewContentsOrig 
      Caption         =   "View contents"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox AscTab 
      Height          =   2985
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox FreqTab 
      Height          =   2985
      Index           =   1
      Left            =   8040
      TabIndex        =   2
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ListBox AscTab 
      Height          =   2985
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ListBox FreqTab 
      Height          =   2985
      Index           =   0
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Sort on Frequentie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   19
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Sort on ASCII"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Sort on Frequentie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   17
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sort on ASCII"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   " 0                                               128                                            255"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   7080
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   " 0                                               128                                            255"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Label FileSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   3840
      Width           =   4575
   End
   Begin VB.Label FileSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label LowValue 
      Caption         =   "Lowest"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   6840
      Width           =   800
   End
   Begin VB.Label MidValue 
      Caption         =   "Midium"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   5400
      Width           =   800
   End
   Begin VB.Label MaxValue 
      Caption         =   "Maximum"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   800
   End
   Begin VB.Label LowValue 
      Caption         =   "Lowest"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   800
   End
   Begin VB.Label MidValue 
      Caption         =   "Midium"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   800
   End
   Begin VB.Label MaxValue 
      Caption         =   "Maximum"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   800
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   511
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   510
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   509
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   508
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   507
      X1              =   360
      X2              =   360
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   506
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   505
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   504
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   503
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   502
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   501
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   500
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   499
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   498
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   497
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   496
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   495
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   494
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   493
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   492
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   491
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   490
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   489
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   488
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   487
      X1              =   480
      X2              =   480
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   486
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   485
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   484
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   483
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   482
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   481
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   480
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   479
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   478
      X1              =   480
      X2              =   480
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   477
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   476
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   475
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   474
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   473
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   472
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   471
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   470
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   469
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   468
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   467
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   466
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   465
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   464
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   463
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   462
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   461
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   6240
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   460
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   459
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   458
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   457
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   456
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   455
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   454
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   453
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   452
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   451
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   450
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   449
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   448
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   447
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   446
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   445
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   444
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   443
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   442
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   441
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   440
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   439
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   438
      X1              =   240
      X2              =   240
      Y1              =   5040
      Y2              =   6120
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   437
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   436
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   435
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   434
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   433
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   432
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   431
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   430
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   429
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   428
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   427
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   426
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   425
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   424
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   423
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   422
      X1              =   240
      X2              =   240
      Y1              =   5040
      Y2              =   6120
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   421
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   420
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   419
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   418
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   417
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   416
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   415
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   414
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   413
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   412
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   411
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   410
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   409
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   408
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   407
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   406
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   405
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   404
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   403
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   402
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   401
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   400
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   399
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   398
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   397
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   396
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   395
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   394
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   393
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   392
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   391
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   390
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   389
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   388
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   387
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   386
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   385
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   384
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   383
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   382
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   381
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   380
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   379
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   378
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   377
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   376
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   375
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   374
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   373
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   372
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   371
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   370
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   369
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   368
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   367
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   366
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   365
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   364
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   363
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   362
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   361
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   360
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   359
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   358
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   357
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   356
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   355
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   354
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   353
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   352
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   351
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   350
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   349
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   348
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   347
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   346
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   345
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   344
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   343
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   342
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   341
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   340
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   339
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   338
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   337
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   336
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   335
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   334
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   333
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   332
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   331
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   330
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   329
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   328
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   327
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   326
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   325
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   324
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   323
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   322
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   321
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   320
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   319
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   318
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   317
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   316
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   315
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   314
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   313
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   312
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   311
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   310
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   309
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   308
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   307
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   306
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   305
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   304
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   303
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   302
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   301
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   300
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   299
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   298
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   297
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   296
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   295
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   294
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   293
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   292
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   291
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   290
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   289
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   288
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   287
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   286
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   285
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   284
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   283
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   282
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   281
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   280
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   279
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   278
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   277
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   276
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   275
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   274
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   273
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   272
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   271
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   270
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   269
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   268
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   267
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   266
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   265
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   264
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   263
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   262
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   261
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   260
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   259
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   258
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   257
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   256
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   255
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   254
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   253
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   252
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   251
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   250
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   249
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   248
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   247
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   246
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   245
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   244
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   243
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   242
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   241
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   240
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   239
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   238
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   237
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   236
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   235
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   234
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   233
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   232
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   231
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   230
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   229
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   228
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   227
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   226
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   225
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   224
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   223
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   222
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   221
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   220
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   219
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   218
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   217
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   216
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   215
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   214
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   213
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   212
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   211
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   210
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   209
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   208
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   207
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   206
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   205
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   204
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   203
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   202
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   201
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   200
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   199
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   198
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   197
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   196
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   195
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   194
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   193
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   192
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   191
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   190
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   189
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   188
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   187
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   186
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   185
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   184
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   183
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   182
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   181
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   180
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   179
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   178
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   177
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   176
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   175
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   174
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   173
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   172
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   171
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   170
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   169
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   168
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   167
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   166
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   165
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   164
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   163
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   162
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   161
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   160
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   159
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   158
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   157
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   156
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   155
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   154
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   153
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   152
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   151
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   150
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   149
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   148
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   147
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   146
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   145
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   144
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   143
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   142
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   141
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   140
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   139
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   138
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   137
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   136
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   135
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   134
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   133
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   132
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   131
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   130
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   129
      X1              =   360
      X2              =   360
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   128
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   127
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   126
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   125
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   124
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   123
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   122
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   121
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   120
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   119
      X1              =   360
      X2              =   360
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   118
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   117
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   116
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   115
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   114
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   113
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   112
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   111
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   110
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   109
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   108
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   107
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   106
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   105
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   104
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   103
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   102
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   101
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   100
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   99
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   98
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   97
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   96
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   95
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   94
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   93
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   92
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   91
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   90
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   89
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   88
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   87
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   86
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   85
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   84
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   83
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   82
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   81
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   80
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   79
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   78
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   77
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   76
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   75
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   74
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   73
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   72
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   71
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   70
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   69
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   68
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   67
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   66
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   65
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   64
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   63
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   62
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   61
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   60
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   59
      X1              =   360
      X2              =   360
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   58
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   57
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   56
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   55
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   54
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   53
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   52
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   51
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   50
      X1              =   480
      X2              =   480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   49
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   48
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   47
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   46
      X1              =   480
      X2              =   480
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   45
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   44
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   43
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   42
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   41
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   40
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   39
      X1              =   360
      X2              =   360
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   38
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   37
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   36
      X1              =   480
      X2              =   480
      Y1              =   5880
      Y2              =   6960
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   35
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   34
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   33
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   32
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   31
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   30
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   29
      X1              =   360
      X2              =   360
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   28
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   27
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   26
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   25
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   24
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   23
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   22
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   21
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   20
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   19
      X1              =   360
      X2              =   360
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   18
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   17
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   16
      X1              =   480
      X2              =   480
      Y1              =   5520
      Y2              =   6600
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   15
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   14
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   13
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   12
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   11
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   10
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   9
      X1              =   360
      X2              =   360
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   8
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   7
      X1              =   240
      X2              =   240
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   6
      X1              =   480
      X2              =   480
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   5
      X1              =   240
      X2              =   240
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   4
      X1              =   480
      X2              =   480
      Y1              =   6000
      Y2              =   7080
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Line Bars 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   480
      X2              =   480
      Y1              =   5280
      Y2              =   6360
   End
   Begin VB.Shape Graphic 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   3015
      Index           =   1
      Left            =   1320
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Shape Graphic 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   3015
      Index           =   0
      Left            =   1320
      Top             =   480
      Width           =   4575
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Load 
         Caption         =   "&Load"
      End
      Begin VB.Menu SaveSourceAs 
         Caption         =   "&Save source as"
      End
      Begin VB.Menu SaveTargetas 
         Caption         =   "S&ave Target as"
      End
      Begin VB.Menu Blank1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu extra 
      Caption         =   "&Extra"
      Begin VB.Menu CopyWorkToOrg 
         Caption         =   "Replace Source with Target"
      End
      Begin VB.Menu RestoreOrig 
         Caption         =   "Restore original Data"
      End
      Begin VB.Menu CompareSWT 
         Caption         =   "Compare Source with Target"
      End
      Begin VB.Menu AutoDecode 
         Caption         =   "Auto Decode/Decompress"
      End
      Begin VB.Menu CalculateEntropy 
         Caption         =   "Calculate Entropy"
      End
   End
   Begin VB.Menu Compressors 
      Caption         =   "&Compressors"
      Begin VB.Menu Huffman 
         Caption         =   "Huffman"
         Begin VB.Menu HuffLong 
            Caption         =   "LongDict"
         End
         Begin VB.Menu HuffShort 
            Caption         =   "ShortDict"
         End
         Begin VB.Menu Huff16 
            Caption         =   "16 Chars"
         End
         Begin VB.Menu HuffmanDynamic 
            Caption         =   "Dynamic"
         End
         Begin VB.Menu HuffNoGreed 
            Caption         =   "Non Greedy"
            Begin VB.Menu HuffNoGreed1 
               Caption         =   "Type 1"
            End
            Begin VB.Menu HuffNoGreed2 
               Caption         =   "Type 2"
            End
         End
      End
      Begin VB.Menu LZW 
         Caption         =   "LZW"
         Begin VB.Menu LZWStat 
            Caption         =   "Static"
         End
         Begin VB.Menu LZWHash 
            Caption         =   "Static with Hashing"
         End
         Begin VB.Menu LZWDyn 
            Caption         =   "Dynamic  (9 to 16 bits)"
         End
         Begin VB.Menu LZWDynamicHash 
            Caption         =   "Dynamic with Hashing"
         End
         Begin VB.Menu LZWPre 
            Caption         =   "Predefined"
         End
         Begin VB.Menu LZWMultidict1 
            Caption         =   "Dynamic Multidictionary"
            Begin VB.Menu Multy1 
               Caption         =   "1 Stream"
            End
            Begin VB.Menu Multy4S 
               Caption         =   "4 streams"
            End
         End
         Begin VB.Menu LZW1Dict 
            Caption         =   "1 Dictionary (Like LZSS)"
         End
      End
      Begin VB.Menu LZSS 
         Caption         =   "LZSS"
         Begin VB.Menu LZSSNorm 
            Caption         =   "Normal"
         End
         Begin VB.Menu LZZ2 
            Caption         =   "Normal but no use of strings"
         End
         Begin VB.Menu LZSS3 
            Caption         =   "bytes and strings"
         End
         Begin VB.Menu LZSSLazy 
            Caption         =   "With Lazy Matching"
         End
      End
      Begin VB.Menu RLE 
         Caption         =   "RLE"
         Begin VB.Menu RLE4isRun 
            Caption         =   "4 = run"
         End
         Begin VB.Menu RLEVar 
            Caption         =   "RLE-Var 1 run"
         End
         Begin VB.Menu RLEVarLoop 
            Caption         =   "RLE-Var Loop"
         End
         Begin VB.Menu RLEVarWidth 
            Caption         =   "RLE variable width"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Ari 
         Caption         =   "Arithmetic"
         Begin VB.Menu AriStat 
            Caption         =   "Static"
         End
         Begin VB.Menu AriDyn 
            Caption         =   "Dynamic"
         End
         Begin VB.Menu AriShortDict 
            Caption         =   "Dynamic with dictionary"
         End
         Begin VB.Menu AriShortDictRescale 
            Caption         =   "Dynamic with dictionary with rescale"
         End
         Begin VB.Menu AriDMC 
            Caption         =   "Dynamic Bitwise Coding"
         End
         Begin VB.Menu AriDMCRescale 
            Caption         =   "Dynamic Bitwise Coding with rescale"
         End
      End
      Begin VB.Menu LBE 
         Caption         =   "Location Based"
         Begin VB.Menu LBEFlat 
            Caption         =   "LBE-Flat"
         End
         Begin VB.Menu LBE_3D 
            Caption         =   "LBE-3D"
         End
         Begin VB.Menu LBE_3D_2 
            Caption         =   "LBE-3D / 2"
         End
      End
      Begin VB.Menu Grouping 
         Caption         =   "Grouping"
         Begin VB.Menu group64 
            Caption         =   "64"
         End
         Begin VB.Menu GroupingSmart 
            Caption         =   "Smart"
            Begin VB.Menu SmartGr1Stream 
               Caption         =   "1 stream"
            End
            Begin VB.Menu SmartGr4Streams 
               Caption         =   "4 Streams"
            End
         End
      End
      Begin VB.Menu VBC 
         Caption         =   "VBC"
         Begin VB.Menu VBC1_Run1 
            Caption         =   "1 Run - 1"
         End
         Begin VB.Menu VBC1_Run2 
            Caption         =   "1 Run - 2"
         End
         Begin VB.Menu VBCReorderble 
            Caption         =   "Reorderble"
         End
         Begin VB.Menu VBCdynamic 
            Caption         =   "Dynamic"
            Begin VB.Menu VBCdynamic1 
               Caption         =   "Type 1"
            End
            Begin VB.Menu VBCdynamic2 
               Caption         =   "Type 2"
            End
         End
      End
      Begin VB.Menu Eliminator 
         Caption         =   "Eliminator"
         Begin VB.Menu EL1run 
            Caption         =   "1 run"
         End
         Begin VB.Menu ElimLoop 
            Caption         =   "Loop till no compression"
         End
      End
      Begin VB.Menu Combiner 
         Caption         =   "Combiner"
         Begin VB.Menu Comb2 
            Caption         =   "2 bytes"
         End
         Begin VB.Menu comb3 
            Caption         =   "3 bytes"
         End
         Begin VB.Menu CombVar 
            Caption         =   "Variable"
         End
      End
      Begin VB.Menu Reducer 
         Caption         =   "Reducer"
         Begin VB.Menu ReducerStat 
            Caption         =   "Static"
         End
         Begin VB.Menu ReducerDynamic 
            Caption         =   "Dynamic"
         End
         Begin VB.Menu RedDynPre1 
            Caption         =   "Dynamic Predefined 1"
         End
         Begin VB.Menu RedDynPre2 
            Caption         =   "Dynamic Predefined 2"
         End
         Begin VB.Menu RedDynPre3 
            Caption         =   "Dynamic Predefined 3"
         End
         Begin VB.Menu ReducerGol 
            Caption         =   "Dynamic with golomb codes"
         End
         Begin VB.Menu ReducerDynEG 
            Caption         =   "Dynamic with Elias gamma codes"
         End
         Begin VB.Menu RedDynHuff 
            Caption         =   "Dynamic with Huffcodes"
         End
         Begin VB.Menu RedHalfDict 
            Caption         =   "Dynamic Half Dictionary with Huffcodes"
         End
         Begin VB.Menu Red16 
            Caption         =   "Dynamic 16 dict with Huffcodes"
         End
      End
      Begin VB.Menu Paring 
         Caption         =   "Paring"
         Begin VB.Menu Paring64 
            Caption         =   "64 chars"
         End
         Begin VB.Menu Paring128 
            Caption         =   "128 chars"
         End
      End
      Begin VB.Menu Shortener 
         Caption         =   "Shortener"
      End
      Begin VB.Menu Stripper1 
         Caption         =   "Stripper"
      End
      Begin VB.Menu word 
         Caption         =   "Word"
         Begin VB.Menu Word1 
            Caption         =   "Needs equal number"
         End
         Begin VB.Menu word2 
            Caption         =   "No special needs"
         End
      End
      Begin VB.Menu Elias 
         Caption         =   "Elias"
         Begin VB.Menu EliasG 
            Caption         =   "Gamma"
         End
         Begin VB.Menu EliasD 
            Caption         =   "Delta"
         End
      End
      Begin VB.Menu Fibonacci 
         Caption         =   "Fibonacci"
      End
      Begin VB.Menu Orderer 
         Caption         =   "Orderer"
         Begin VB.Menu OrdererCompress 
            Caption         =   "Compresses only the low bytes <64"
         End
      End
   End
   Begin VB.Menu Coders 
      Caption         =   "C&oders"
      Begin VB.Menu Differentiator 
         Caption         =   "Differentiator"
      End
      Begin VB.Menu FreqShift 
         Caption         =   "Frequentie Shifter"
      End
      Begin VB.Menu BWT 
         Caption         =   "BWT"
         Begin VB.Menu BWT1 
            Caption         =   "Method 1"
         End
         Begin VB.Menu BWT2 
            Caption         =   "Method 2"
         End
         Begin VB.Menu BWTnoRecursion 
            Caption         =   "Method 3"
         End
      End
      Begin VB.Menu Fix128 
         Caption         =   "Fix 128"
         Begin VB.Menu Fix128SB7 
            Caption         =   "Split bit 7"
         End
         Begin VB.Menu Fix128SB1 
            Caption         =   "Split bit 1"
         End
      End
      Begin VB.Menu Seperator 
         Caption         =   "Seperator"
      End
      Begin VB.Menu Base64 
         Caption         =   "Base 64"
      End
      Begin VB.Menu Flat64 
         Caption         =   "Flatter 64"
      End
      Begin VB.Menu Flatter16 
         Caption         =   "Flatter 16"
      End
      Begin VB.Menu Numerator 
         Caption         =   "Numerator"
         Begin VB.Menu Numerator1 
            Caption         =   "Version 1"
         End
         Begin VB.Menu Numerator2 
            Caption         =   "Version 2"
         End
      End
      Begin VB.Menu MTF 
         Caption         =   "Move To Front"
         Begin VB.Menu MTFNoHead 
            Caption         =   "No Header"
         End
         Begin VB.Menu MTFWithHead 
            Caption         =   "With Header"
         End
      End
      Begin VB.Menu UTF 
         Caption         =   "Used To Front"
      End
      Begin VB.Menu SSCoder 
         Caption         =   "Sort & Swap"
      End
      Begin VB.Menu ValDShift 
         Caption         =   "Value down shifter"
      End
      Begin VB.Menu VUS 
         Caption         =   "Value Up Shifter"
      End
      Begin VB.Menu ValTwist 
         Caption         =   "Value Twister"
      End
      Begin VB.Menu AddDifferentiator 
         Caption         =   "Add Differantiator"
      End
      Begin VB.Menu Scrambler 
         Caption         =   "Scrambler"
      End
   End
   Begin VB.Menu CRC32calc 
      Caption         =   "CRC 32"
      Begin VB.Menu CalcCRC32Source 
         Caption         =   "Calculate Source"
      End
      Begin VB.Menu CalcCRC32Target 
         Caption         =   "Calculate Target"
      End
   End
End
Attribute VB_Name = "Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddDifferentiator_Click()
    Call Start_Coder(Coder_AddDifferantiator)
End Sub

Private Sub AriDMC_Click()
    Call Start_Compressor(Compressor_Arithmetic_DMC)
End Sub

Private Sub AriDMCRescale_Click()
    Call Start_Compressor(Compressor_Arithmetic_DMC_Rescale)
End Sub

Private Sub AriDyn_Click()
    Call Start_Compressor(Compressor_Arithmetic_Dynamic)
End Sub

Private Sub AriShortDict_Click()
    Call Start_Compressor(Compressor_Arithmetic_Dynamic_With_Dict)
End Sub

Private Sub AriShortDictRescale_Click()
    Call Start_Compressor(Compressor_Arithmetic_Dynamic_With_Dict_Rescale)
End Sub

Private Sub AriStat_Click()
    Call Start_Compressor(Compressor_Arithmetic)
End Sub

Private Sub AutoDecode_Click()
    Call Auto_Decode_Depack
End Sub

Private Sub Base64_Click()
    Call Start_Coder(Coder_Base64)
End Sub

Private Sub btnViewContentsOrig_Click()
    Call Show_Contents(OriginalArray)
End Sub

Private Sub btnViewContentsTarget_Click()
    Call Show_Contents(WorkArray)
End Sub

Private Sub BWT1_Click()
    Call Start_Coder(Coder_BWT)
End Sub

Private Sub BWT2_Click()
    Call Start_Coder(Coder_BWT2)
End Sub

Private Sub BWTnoRecursion_Click()
    Call Start_Coder(Coder_BWT_noRecursion)
End Sub

Private Sub CalcCRC32Source_Click()
    Dim CRCSource As Long
    On Error GoTo No_Data
    CRCSource = UBound(OriginalArray)
    On Error GoTo 0
    CRCSource = calcCRC32(OriginalArray)
    MsgBox Hex(CRCSource)
No_Data:
End Sub

Private Sub CalcCRC32Target_Click()
    Dim CRCTarget As Long
    On Error GoTo No_Data
    CRCTarget = UBound(WorkArray)
    On Error GoTo 0
    CRCTarget = calcCRC32(WorkArray)
    MsgBox Hex(CRCTarget)
No_Data:
End Sub

Private Sub CalculateEntropy_Click()
    Call Calculate_Entropy(OriginalArray)
End Sub

Private Sub Comb2_Click()
    Call Start_Compressor(Compressor_Combiner2Bytes)
End Sub

Private Sub comb3_Click()
    Call Start_Compressor(Compressor_Combiner3Bytes)
End Sub

Private Sub CombVar_Click()
    Call Start_Compressor(Compressor_CombinerVariable)
End Sub

Private Sub CompareSWT_Click()
    Call Compare_Source_With_Target
End Sub

Private Sub CopyWorkToOrg_Click()
    If UBound(WorkArray) = 0 Then
        MsgBox "There is nothing to copy"
        Exit Sub
    End If
    Call Copy_Work2Orig
    Call AddCoder2List(LastCoder)
    Call Show_Statistics(True, OriginalArray)
End Sub

Private Sub Differentiator_Click()
    Call Start_Coder(Coder_Differantiator)
End Sub

Private Sub EL1run_Click()
    Call Start_Compressor(Compressor_Eliminator)
End Sub

Private Sub EliasD_Click()
    Call Start_Compressor(Compressor_EliasDelta)
End Sub

Private Sub EliasG_Click()
    Call Start_Compressor(Compressor_EliasGamma)
End Sub

Private Sub ElimLoop_Click()
    Call Start_Compressor(Compressor_Eliminator_Loop)
End Sub

Private Sub ExitProg_Click()
    End
End Sub

Private Sub Fibonacci_Click()
    Call Start_Compressor(Compressor_Fibonacci)
End Sub

Private Sub Fix128SB1_Click()
    Call Start_Coder(Coder_Fix128B)
End Sub

Private Sub Fix128SB7_Click()
    Call Start_Coder(Coder_Fix128)
End Sub

Private Sub Flat64_Click()
    Call Start_Coder(Coder_Flatter64)
End Sub

Private Sub Flatter16_Click()
    Call Start_Coder(Coder_Flatter16)
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    Dim MaxWidth As Double
    For X = 0 To 1
        MaxWidth = Graphic(X).Width / 256
        For Y = 0 To 255
            Bars(X * 256 + Y).BorderWidth = 1
            Bars(X * 256 + Y).BorderStyle = 1
            Bars(X * 256 + Y).X1 = Graphic(X).Left + Y * MaxWidth
            Bars(X * 256 + Y).X2 = Bars(X * 256 + Y).X1
            Bars(X * 256 + Y).Y2 = Graphic(X).Top + Graphic(X).Height
            Bars(X * 256 + Y).Y1 = Bars(X * 256 + Y).Y2
            Bars(X * 256 + Y).Visible = True
        Next
    Next
    RGBColor(0) = vbBlue
    RGBColor(1) = vbCyan
    RGBColor(2) = vbGreen
    RGBColor(3) = vbMagenta
    RGBColor(4) = vbRed
    RGBColor(5) = vbYellow
    RGBColor(6) = vbWhite
    Call Init_CoderNameDataBase
    ReDim OriginalArray(0)
    ReDim WorkArray(0)
    ReDim UsedCodecs(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub FreqShift_Click()
    Call Start_Coder(Coder_FrequentieShifter)
End Sub

Private Sub group64_Click()
    Call Start_Compressor(Compressor_Grouping64)
End Sub

Private Sub Huff16_Click()
    Call Start_Compressor(Compressor_HuffmanShort16Chars)
End Sub

Private Sub HuffLong_Click()
    Call Start_Compressor(Compressor_HuffmanCodes)
End Sub

Private Sub HuffmanDynamic_Click()
    Call Start_Compressor(Compressor_Huffman_Dynamic)
End Sub

Private Sub HuffNoGreed1_Click()
    Call Start_Compressor(Compressor_Huffman_Non_Greedy)
End Sub

Private Sub HuffNoGreed2_Click()
    Call Start_Compressor(Compressor_Huffman_Non_Greedy2)
End Sub

Private Sub HuffShort_Click()
    Call Start_Compressor(Compressor_HuffmanShortDict)
End Sub

Private Sub LBE_3D_2_Click()
    Call Start_Compressor(Compressor_LBE_3D_2)
End Sub

Private Sub LBE_3D_Click()
    Call Start_Compressor(Compressor_LBE_3D)
End Sub

Private Sub LBEFlat_Click()
    Call Start_Compressor(Compressor_LBE_Flat)
End Sub

Private Sub Load_Click()
    Dim OldFileName As String
    OldFileName = LoadFileName
    Cdlg.DialogTitle = "Select the file you want to explore"
    Cdlg.FileName = ""
    Cdlg.ShowOpen
    LoadFileName = Cdlg.FileName
    Call load_File(LoadFileName)
    If LoadFileName = "" Then LoadFileName = OldFileName
End Sub

Private Sub LZSS3_Click()
    Call Start_Compressor(Compressor_LZSS_3)
End Sub

Private Sub LZSSLazy_Click()
    Call Start_Compressor(Compressor_LZSS_Lazy_Matching)
End Sub

Private Sub LZSSNorm_Click()
    Call Start_Compressor(Compressor_LZSS)
End Sub

Private Sub LZW1Dict_Click()
    Call Start_Compressor(Compressor_LZW_LZSS)
End Sub

Private Sub LZWDyn_Click()
    Call Start_Compressor(Compressor_LZW_Dynamic)
End Sub

Private Sub LZWDynamicHash_Click()
    Call Start_Compressor(Compressor_LZW_Dynamic_Hash)
End Sub

Private Sub LZWHash_Click()
    Call Start_Compressor(Compressor_LZW_Static_Hash)
End Sub

Private Sub LZWPre_Click()
    Call Start_Compressor(Compressor_LZW_Predefined)
End Sub

Private Sub LZWStat_Click()
    Call Start_Compressor(Compressor_LZW_Static)
End Sub

Private Sub LZZ2_Click()
    Call Start_Compressor(Compressor_LZSS_2)
End Sub

Private Sub MTFNoHead_Click()
    Call Start_Coder(Coder_MTF_No_Header)
End Sub

Private Sub MTFWithHead_Click()
    Call Start_Coder(Coder_MTF_With_Header)
End Sub

Private Sub Multy1_Click()
    Call Start_Compressor(Compressor_LZW_Multidict1Stream)
End Sub

Private Sub Multy4S_Click()
    Call Start_Compressor(Compressor_LZW_Multidict4Streams)
End Sub

Private Sub Numerator1_Click()
    Call Start_Coder(Coder_Numerator)
End Sub

Private Sub Numerator2_Click()
    Call Start_Coder(Coder_Numerator2)
End Sub

Private Sub OrdererCompress_Click()
    Call Start_Compressor(Compressor_Orderer)
End Sub

Private Sub Paring128_Click()
    Call Start_Compressor(Compressor_Pairing128)
End Sub

Private Sub Paring64_Click()
    Call Start_Compressor(Compressor_Pairing)
End Sub

Private Sub Red16_Click()
    Call Start_Compressor(Compressor_Reducer_Dict16withHuffcodes)
End Sub

Private Sub RedDynHuff_Click()
    Call Start_Compressor(Compressor_Reducer_withHuffcodes)
End Sub

Private Sub RedDynPre1_Click()
    Call Start_Compressor(Compressor_Reducer_Preselect1)
End Sub

Private Sub RedDynPre2_Click()
    Call Start_Compressor(Compressor_Reducer_Preselect2)
End Sub

Private Sub RedDynPre3_Click()
    Call Start_Compressor(Compressor_Reducer_Preselect3)
End Sub

Private Sub RedHalfDict_Click()
    Call Start_Compressor(Compressor_Reducer_HalfDictwithHuffcodes)
End Sub

Private Sub ReducerDynamic_Click()
    Call Start_Compressor(Compressor_Reducer_Dynamic)
End Sub

Private Sub ReducerDynEG_Click()
    Call Start_Compressor(Compressor_Reducer_Dynamic_Elias_Gamma)
End Sub

Private Sub ReducerGol_Click()
    Call Start_Compressor(Compressor_Reducer_Dynamic_Golomb)
End Sub

Private Sub ReducerStat_Click()
    Call Start_Compressor(Compressor_Reducer_Static)
End Sub

Private Sub RestoreOrig_Click()
    If LoadFileName = "" Then
        MsgBox "There was'nt yet original data"
        Exit Sub
    End If
    Call load_File(LoadFileName)
End Sub

Private Sub RLE4isRun_Click()
    Call Start_Compressor(Compressor_RLE4isRun)
End Sub

Private Sub RLEVar_Click()
    Call Start_Compressor(Compressor_RLEVar)
End Sub

Private Sub RLEVarLoop_Click()
    Call Start_Compressor(Compressor_RLEVarLoop)
End Sub

Private Sub RLEVarWidth_Click()
    Call Start_Compressor(Compressor_RLE3)
End Sub

Private Sub SaveSourceAs_Click()
    Call Save_File_As(OriginalArray, True)
End Sub

Private Sub SaveTargetas_Click()
    Call Save_File_As(WorkArray, False)
End Sub

Private Sub Scrambler_Click()
    Call Start_Coder(Coder_Scrambler)
End Sub

Private Sub Seperator_Click()
    Call Start_Coder(Coder_Seperator)
End Sub

Private Sub Shortener_Click()
    Call Start_Compressor(Compressor_Shortener)
End Sub

Private Sub SmartGr1Stream_Click()
    Call Start_Compressor(Compressor_SmartGrouping)
End Sub

Private Sub SmartGr4Streams_Click()
    Call Start_Compressor(Compressor_SmartGrouping4Streams)
End Sub

Private Sub SSCoder_Click()
    Call Start_Coder(Coder_SortSwap)
End Sub

Private Sub Stripper1_Click()
    Call Start_Compressor(Compressor_Stripper)
End Sub

Private Sub UTF_Click()
    Call Start_Coder(Coder_Used_To_Front)
End Sub

Private Sub ValDShift_Click()
    Call Start_Coder(Coder_ValueDownShift)
End Sub

Private Sub ValTwist_Click()
    Call Start_Coder(Coder_ValueTwister)
End Sub

Private Sub VBC1_Run1_Click()
    Call Start_Compressor(Compressor_VBC1)
End Sub

Private Sub VBC1_Run2_Click()
    Call Start_Compressor(Compressor_VBC2)
End Sub

Private Sub VBCdynamic1_Click()
    Call Start_Compressor(Compressor_VBC_Dynamic)
End Sub

Private Sub VBCdynamic2_Click()
    Call Start_Compressor(Compressor_VBC_Dynamic2)
End Sub

Private Sub VBCReorderble_Click()
    Call Start_Compressor(Compressor_VBCR)
End Sub

Private Sub VUS_Click()
    Call Start_Coder(Coder_ValueUpShift)
End Sub

Private Sub Word1_Click()
    Call Start_Compressor(Compressor_Word1)
End Sub

Private Sub word2_Click()
    Call Start_Compressor(Compressor_Word2)
End Sub
