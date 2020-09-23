VERSION 5.00
Begin VB.Form frmTetris 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blocks!      V 1.5"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   1620
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTetris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTetris.frx":08CA
   ScaleHeight     =   8775
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSpecialColors 
      Interval        =   10
      Left            =   720
      Top             =   8040
   End
   Begin VB.Timer tmrDescend 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   7560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "MindPortal Entertainment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "http://mindportal.cjb.net"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   8160
      Width           =   3975
   End
   Begin VB.Label lblPreBlockType 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   1100
      Width           =   1455
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   8160
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   7920
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   7680
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   7440
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   8160
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   7920
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7680
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   7440
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   8160
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   7920
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   7680
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   7440
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   8160
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   7920
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   7680
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpPreview 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   7440
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Next Block - "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "E&xit"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Options"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label lblNewGame 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "&New Game"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Shape shpExit 
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Shape shpOptions 
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Shape shpNewGame 
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   5775
      Left            =   1200
      Top             =   1080
      Width           =   3930
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   383
      Left            =   4800
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   382
      Left            =   4560
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   381
      Left            =   4320
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   380
      Left            =   4080
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   379
      Left            =   3840
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   378
      Left            =   3600
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   377
      Left            =   3360
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   376
      Left            =   3120
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   375
      Left            =   2880
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   374
      Left            =   2640
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   373
      Left            =   2400
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   372
      Left            =   2160
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   371
      Left            =   1920
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   370
      Left            =   1680
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   369
      Left            =   1440
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   368
      Left            =   1200
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   367
      Left            =   4800
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   366
      Left            =   4560
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   365
      Left            =   4320
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   364
      Left            =   4080
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   363
      Left            =   3840
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   362
      Left            =   3600
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   361
      Left            =   3360
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   360
      Left            =   3120
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   359
      Left            =   2880
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   358
      Left            =   2640
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   357
      Left            =   2400
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   356
      Left            =   2160
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   355
      Left            =   1920
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   354
      Left            =   1680
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   353
      Left            =   1440
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   352
      Left            =   1200
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   351
      Left            =   4800
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   350
      Left            =   4560
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   349
      Left            =   4320
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   348
      Left            =   4080
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   347
      Left            =   3840
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   346
      Left            =   3600
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   345
      Left            =   3360
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   344
      Left            =   3120
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   343
      Left            =   2880
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   342
      Left            =   2640
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   341
      Left            =   2400
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   340
      Left            =   2160
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   339
      Left            =   1920
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   338
      Left            =   1680
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   337
      Left            =   1440
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   336
      Left            =   1200
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   335
      Left            =   4800
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   334
      Left            =   4560
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   333
      Left            =   4320
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   332
      Left            =   4080
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   331
      Left            =   3840
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   330
      Left            =   3600
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   329
      Left            =   3360
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   328
      Left            =   3120
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   327
      Left            =   2880
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   326
      Left            =   2640
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   325
      Left            =   2400
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   324
      Left            =   2160
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   323
      Left            =   1920
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   322
      Left            =   1680
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   321
      Left            =   1440
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   320
      Left            =   1200
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   319
      Left            =   4800
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   318
      Left            =   4560
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   317
      Left            =   4320
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   316
      Left            =   4080
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   315
      Left            =   3840
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   314
      Left            =   3600
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   313
      Left            =   3360
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   312
      Left            =   3120
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   311
      Left            =   2880
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   310
      Left            =   2640
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   309
      Left            =   2400
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   308
      Left            =   2160
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   307
      Left            =   1920
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   306
      Left            =   1680
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   305
      Left            =   1440
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   304
      Left            =   1200
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   303
      Left            =   4800
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   302
      Left            =   4560
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   301
      Left            =   4320
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   300
      Left            =   4080
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   299
      Left            =   3840
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   298
      Left            =   3600
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   297
      Left            =   3360
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   296
      Left            =   3120
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   295
      Left            =   2880
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   294
      Left            =   2640
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   293
      Left            =   2400
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   292
      Left            =   2160
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   291
      Left            =   1920
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   290
      Left            =   1680
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   289
      Left            =   1440
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   288
      Left            =   1200
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   287
      Left            =   4800
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   286
      Left            =   4560
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   285
      Left            =   4320
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   284
      Left            =   4080
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   283
      Left            =   3840
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   282
      Left            =   3600
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   281
      Left            =   3360
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   280
      Left            =   3120
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   279
      Left            =   2880
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   278
      Left            =   2640
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   277
      Left            =   2400
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   276
      Left            =   2160
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   275
      Left            =   1920
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   274
      Left            =   1680
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   273
      Left            =   1440
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   272
      Left            =   1200
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   271
      Left            =   4800
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   270
      Left            =   4560
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   269
      Left            =   4320
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   268
      Left            =   4080
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   267
      Left            =   3840
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   266
      Left            =   3600
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   265
      Left            =   3360
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   264
      Left            =   3120
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   263
      Left            =   2880
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   262
      Left            =   2640
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   261
      Left            =   2400
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   260
      Left            =   2160
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   259
      Left            =   1920
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   258
      Left            =   1680
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   257
      Left            =   1440
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   256
      Left            =   1200
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   255
      Left            =   4800
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   254
      Left            =   4560
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   253
      Left            =   4320
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   252
      Left            =   4080
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   251
      Left            =   3840
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   250
      Left            =   3600
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   249
      Left            =   3360
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   248
      Left            =   3120
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   247
      Left            =   2880
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   246
      Left            =   2640
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   245
      Left            =   2400
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   244
      Left            =   2160
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   243
      Left            =   1920
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   242
      Left            =   1680
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   241
      Left            =   1440
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   240
      Left            =   1200
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   239
      Left            =   4800
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   238
      Left            =   4560
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   237
      Left            =   4320
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   236
      Left            =   4080
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   235
      Left            =   3840
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   234
      Left            =   3600
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   233
      Left            =   3360
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   232
      Left            =   3120
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   231
      Left            =   2880
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   230
      Left            =   2640
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   229
      Left            =   2400
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   228
      Left            =   2160
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   227
      Left            =   1920
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   226
      Left            =   1680
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   225
      Left            =   1440
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   224
      Left            =   1200
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   223
      Left            =   4800
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   222
      Left            =   4560
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   221
      Left            =   4320
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   220
      Left            =   4080
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   219
      Left            =   3840
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   218
      Left            =   3600
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   217
      Left            =   3360
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   216
      Left            =   3120
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   215
      Left            =   2880
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   214
      Left            =   2640
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   213
      Left            =   2400
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   212
      Left            =   2160
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   211
      Left            =   1920
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   210
      Left            =   1680
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   209
      Left            =   1440
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   208
      Left            =   1200
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   207
      Left            =   4800
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   206
      Left            =   4560
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   205
      Left            =   4320
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   204
      Left            =   4080
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   203
      Left            =   3840
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   202
      Left            =   3600
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   201
      Left            =   3360
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   200
      Left            =   3120
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   199
      Left            =   2880
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   198
      Left            =   2640
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   197
      Left            =   2400
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   196
      Left            =   2160
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   195
      Left            =   1920
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   194
      Left            =   1680
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   193
      Left            =   1440
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   192
      Left            =   1200
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   191
      Left            =   4800
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   190
      Left            =   4560
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   189
      Left            =   4320
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   188
      Left            =   4080
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   187
      Left            =   3840
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   186
      Left            =   3600
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   185
      Left            =   3360
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   184
      Left            =   3120
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   183
      Left            =   2880
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   182
      Left            =   2640
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   181
      Left            =   2400
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   180
      Left            =   2160
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   179
      Left            =   1920
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   178
      Left            =   1680
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   177
      Left            =   1440
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   176
      Left            =   1200
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   175
      Left            =   4800
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   174
      Left            =   4560
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   173
      Left            =   4320
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   172
      Left            =   4080
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   171
      Left            =   3840
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   170
      Left            =   3600
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   169
      Left            =   3360
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   168
      Left            =   3120
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   167
      Left            =   2880
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   166
      Left            =   2640
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   165
      Left            =   2400
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   164
      Left            =   2160
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   163
      Left            =   1920
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   162
      Left            =   1680
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   161
      Left            =   1440
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   160
      Left            =   1200
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   159
      Left            =   4800
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   158
      Left            =   4560
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   157
      Left            =   4320
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   156
      Left            =   4080
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   155
      Left            =   3840
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   154
      Left            =   3600
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   153
      Left            =   3360
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   152
      Left            =   3120
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   151
      Left            =   2880
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   150
      Left            =   2640
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   149
      Left            =   2400
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   148
      Left            =   2160
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   147
      Left            =   1920
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   146
      Left            =   1680
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   145
      Left            =   1440
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   144
      Left            =   1200
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   143
      Left            =   4800
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   142
      Left            =   4560
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   141
      Left            =   4320
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   140
      Left            =   4080
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   139
      Left            =   3840
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   138
      Left            =   3600
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   137
      Left            =   3360
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   136
      Left            =   3120
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   135
      Left            =   2880
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   134
      Left            =   2640
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   133
      Left            =   2400
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   132
      Left            =   2160
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   131
      Left            =   1920
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   130
      Left            =   1680
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   129
      Left            =   1440
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   128
      Left            =   1200
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   127
      Left            =   4800
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   126
      Left            =   4560
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   125
      Left            =   4320
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   124
      Left            =   4080
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   123
      Left            =   3840
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   122
      Left            =   3600
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   121
      Left            =   3360
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   120
      Left            =   3120
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   119
      Left            =   2880
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   118
      Left            =   2640
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   117
      Left            =   2400
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   116
      Left            =   2160
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   115
      Left            =   1920
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   114
      Left            =   1680
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   113
      Left            =   1440
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   112
      Left            =   1200
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   111
      Left            =   4800
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   110
      Left            =   4560
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   109
      Left            =   4320
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   108
      Left            =   4080
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   107
      Left            =   3840
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   106
      Left            =   3600
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   105
      Left            =   3360
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   104
      Left            =   3120
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   103
      Left            =   2880
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   102
      Left            =   2640
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   101
      Left            =   2400
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   100
      Left            =   2160
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   99
      Left            =   1920
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   1680
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   1440
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   1200
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   4800
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   4560
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   4320
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   4080
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   3840
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   3600
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   3360
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   3120
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   2880
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   2640
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   2400
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   2160
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   1920
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   1680
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   1440
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   1200
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   4800
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   4560
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   4320
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   4080
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   3840
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   3600
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   3360
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   3120
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   2880
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   2640
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   2400
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   2160
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   1920
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   1680
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   1440
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   1200
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   4800
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   4560
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   4320
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   4080
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   3840
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   3600
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   3360
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   3120
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   2880
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   2640
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   2400
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   2160
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   1920
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   1680
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   1440
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   1200
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   4800
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   4560
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   4320
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   4080
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   3840
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   3600
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   3360
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   3120
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   2880
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   2640
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   2400
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   2160
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   1920
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   1680
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   1440
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   1200
      Top             =   6120
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   4800
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   4560
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   4320
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   4080
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   3840
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   3600
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   3360
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   3120
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   2880
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   2640
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   2400
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   2160
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   1920
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   1680
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   1440
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   1200
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   4800
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   4560
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   4320
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   4080
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   3840
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   3600
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   3360
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   3120
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   2880
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   2640
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   2400
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   2160
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   1920
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   1680
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1440
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape shpBlock 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1200
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "frmTetris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bytBlockType As Byte, PreBlockType As Byte
Private Const SQUAREBLOCK As Byte = 1, TBLOCK As Byte = 2, LINEBLOCK As Byte = 3, RIGHTL As Byte = 4, LEFTL As Byte = 5, SBLOCK As Byte = 6, ZBLOCK As Byte = 7, SCATTERBLOCK As Byte = 8, BOMBBLOCK As Byte = 9
Private BlockColor As String, PreBlockColor As String
Private Block1Pos As Integer, Block2Pos As Integer, Block3Pos As Integer, Block4Pos As Integer
Private PreBlock1Pos As Integer, PreBlock2Pos As Integer, PreBlock3Pos As Integer, PreBlock4Pos As Integer
Private strPreBlockType As String
Private LeftSide(0 To 383) As Integer, RightSide(0 To 399) As Integer, BottomSide(0 To 399) As Integer, AllSide(-32 To 399) As Integer
Private BlockMode As Byte
Private intScore As Long, intTime As Long
Public GameEnabled As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If GameEnabled = True Then
        If KeyCode = vbKeyRight Then
            Call MoveRight
        ElseIf KeyCode = vbKeyLeft Then
            Call MoveLeft
        ElseIf KeyCode = vbKeyDown Then
            Call MoveDown
        ElseIf KeyCode = vbKeyUp Then
            Call MoveUp
        ElseIf KeyCode = vbKeySpace Then
            Call MoveUp
        ElseIf KeyCode = vbKeyN Then
            Call lblNewGame_Click
        ElseIf KeyCode = vbKeyO Then
            Call lblOptions_Click
        ElseIf KeyCode = vbKeyX Then
            Call lblExit_Click
        End If
    End If
End Sub
Sub NewGame()
Dim X As Integer

    GameEnabled = True
     
    For X = 0 To 383
        LeftSide(X) = 0
        RightSide(X) = 0
        BottomSide(X) = 0
        AllSide(X) = 0
        intScore = 0
        lblScore.Caption = intScore
        shpBlock.Item(X).FillColor = vbBlack
    Next X
    
    For X = 0 To 15
        shpPreview.Item(X).FillColor = vbBlack
    Next X
    
    BlockMode = 1
    
    Call NonMoveable
    Call GenerateFirstBlock
    Call GenerateNewBlock
    tmrDescend.Enabled = True
    tmrDescend.Interval = frmOptions.intCurrentLevel
End Sub
Sub ClearBlocks()
    shpBlock.Item(Block1Pos).FillColor = vbBlack
    shpBlock.Item(Block2Pos).FillColor = vbBlack
    shpBlock.Item(Block3Pos).FillColor = vbBlack
    shpBlock.Item(Block4Pos).FillColor = vbBlack
End Sub

Sub MoveUp()
    
    If bytBlockType = TBLOCK Then
        If BlockMode = 1 Then
            If BottomSide(Block2Pos) <> 1 Then
                If AllSide(Block1Pos - 15) = 1 Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 15
        ElseIf BlockMode = 2 Then
            If LeftSide(Block2Pos) = 1 Or AllSide(Block2Pos - 1) = 1 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block2Pos = Block2Pos - 1
            Block3Pos = Block3Pos - 1
            Block4Pos = Block4Pos - 15
        ElseIf BlockMode = 3 Then
            If Block4Pos + 15 <= 383 Then
                If AllSide(Block4Pos + 15) = 1 Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            Call ClearBlocks
            Block4Pos = Block4Pos + 15
        ElseIf BlockMode = 4 Then
            If RightSide(Block3Pos) = 1 Or AllSide(Block3Pos + 1) = 1 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 15
            Block2Pos = Block2Pos + 1
            Block3Pos = Block3Pos + 1
        End If
    ElseIf bytBlockType = LINEBLOCK Then
        If BlockMode = 1 Or BlockMode = 3 Then
            If BottomSide(Block1Pos) <> 1 And Block4Pos + 46 <= 383 Then
                If AllSide(Block2Pos + 16) = 1 Or AllSide(Block3Pos + 31) = 1 Or AllSide(Block4Pos + 46) = 1 Then
                    Exit Sub
                    BlockMode = BlockMode - 1
                End If
            Else
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 1
            Block2Pos = Block2Pos + 16
            Block3Pos = Block3Pos + 31
            Block4Pos = Block4Pos + 46
        ElseIf BlockMode = 2 Or BlockMode = 4 Then
                If AllSide(Block1Pos - 1) = 1 Or AllSide(Block1Pos + 1) = 1 Or AllSide(Block1Pos + 2) = 1 Or _
                    RightSide(Block1Pos) = 1 Or RightSide(Block1Pos + 1) = 1 Or _
                    LeftSide(Block2Pos) = 1 Or LeftSide(Block2Pos - 1) = 1 Or LeftSide(Block2Pos - 2) = 1 Then
                    Exit Sub
                End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 1
            Block2Pos = Block2Pos - 16
            Block3Pos = Block3Pos - 31
            Block4Pos = Block4Pos - 46
        End If
    ElseIf bytBlockType = RIGHTL Then
        If BlockMode = 1 Then
            If AllSide(Block1Pos + 1) = 1 Or AllSide(Block2Pos + 1) = 1 Or AllSide(Block2Pos - 1) = 1 Or _
                LeftSide(Block2Pos) = 1 Or RightSide(Block4Pos) = 1 Then
                    Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 2
            Block3Pos = Block3Pos - 15
            Block4Pos = Block4Pos - 15
        ElseIf BlockMode = 2 Then
            If AllSide(Block1Pos - 1) = 1 Or AllSide(Block4Pos + 16) = 1 Or Block4Pos + 16 > 383 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 2
            Block2Pos = Block2Pos - 15
            Block4Pos = Block4Pos + 15
        ElseIf BlockMode = 3 Then
            If AllSide(Block3Pos - 1) = 1 Or AllSide(Block2Pos + 1) = 1 Or _
                RightSide(Block3Pos) = 1 Then
                    Exit Sub
            End If
            Call ClearBlocks
            Block3Pos = Block3Pos - 15
            Block4Pos = Block4Pos - 17
        ElseIf BlockMode = 4 Then
            If AllSide(Block4Pos + 1) = 1 Or AllSide(Block4Pos + 17) = 1 Or AllSide(Block4Pos + 18) = 1 Or _
                 Block4Pos + 17 > 383 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block2Pos = Block2Pos + 15
            Block3Pos = Block3Pos + 30
            Block4Pos = Block4Pos + 17
        End If
    ElseIf bytBlockType = LEFTL Then
        If BlockMode = 1 Then
            If AllSide(Block2Pos + 1) = 1 Or AllSide(Block2Pos - 1) = 1 Or AllSide(Block4Pos + 1) = 1 Or RightSide(Block2Pos) = 1 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 15
            Block3Pos = Block3Pos - 14
            Block4Pos = Block4Pos + 1
        ElseIf BlockMode = 2 Then
            If AllSide(Block2Pos - 16) = 1 Or AllSide(Block4Pos - 1) = 1 Or AllSide(Block3Pos - 16) = 1 Or BottomSide(Block1Pos) = 1 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 16
            Block2Pos = Block2Pos - 16
            Block3Pos = Block3Pos - 2
            Block4Pos = Block4Pos - 2
        ElseIf BlockMode = 3 Then
            If AllSide(Block2Pos + 1) = 1 Or AllSide(Block3Pos - 1) = 1 Or LeftSide(Block4Pos) = 1 Or RightSide(Block3Pos + 1) = 1 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 16
            Block2Pos = Block2Pos + 31
            Block3Pos = Block3Pos + 17
            Block4Pos = Block4Pos + 2
        ElseIf BlockMode = 4 Then
            If AllSide(Block3Pos + 16) = 1 Or AllSide(Block2Pos + 16) = 1 Or AllSide(Block1Pos + 1) = 1 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 15
            Block2Pos = Block2Pos - 15
            Block3Pos = Block3Pos - 1
            Block4Pos = Block4Pos - 1
        End If
    ElseIf bytBlockType = ZBLOCK Then
        If BlockMode = 1 Or BlockMode = 3 Then
            If AllSide(Block4Pos + 16) = 1 Or AllSide(Block1Pos - 1) = 1 Or _
                Block4Pos + 16 > 383 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 1
            Block2Pos = Block2Pos + 14
            Block3Pos = Block3Pos + 1
            Block4Pos = Block4Pos + 16
        ElseIf BlockMode = 2 Or BlockMode = 4 Then
            If AllSide(Block1Pos + 1) = 1 Or AllSide(Block1Pos + 2) = 1 Or _
                RightSide(Block3Pos) = 1 Then
                    Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 1
            Block2Pos = Block2Pos - 14
            Block3Pos = Block3Pos - 1
            Block4Pos = Block4Pos - 16
        End If
    ElseIf bytBlockType = SBLOCK Then
        If BlockMode = 1 Or BlockMode = 3 Then
            If AllSide(Block3Pos - 1) = 1 Or AllSide(Block3Pos + 14) = 1 Or _
                Block3Pos + 15 > 383 Then
                Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos + 1
            Block2Pos = Block2Pos + 15
            Block4Pos = Block4Pos + 14
        ElseIf BlockMode = 2 Or BlockMode = 4 Then
            If AllSide(Block1Pos - 1) = 1 Or AllSide(Block3Pos + 1) = 1 Or _
                RightSide(Block3Pos) = 1 Then
                    Exit Sub
            End If
            Call ClearBlocks
            Block1Pos = Block1Pos - 1
            Block2Pos = Block2Pos - 15
            Block4Pos = Block4Pos - 14
        End If
    End If
    
    BlockMode = BlockMode + 1
    
    If BlockMode = 5 Then
        BlockMode = 1
    End If
    
    shpBlock.Item(Block1Pos).FillColor = BlockColor
    shpBlock.Item(Block2Pos).FillColor = BlockColor
    shpBlock.Item(Block3Pos).FillColor = BlockColor
    shpBlock.Item(Block4Pos).FillColor = BlockColor
    
End Sub
Sub NonMoveable()
Dim X As Integer

    X = 0
    
    Do
        LeftSide(X) = 1
        X = X + 16
    Loop While X <> 368
    
    X = 0
    
    Do
        BottomSide(X) = 1
        X = X + 1
    Loop While X <> 15
    
    X = 15
    
    Do
        RightSide(X) = 1
        X = X + 16
    Loop While X <> 383
    
    X = -16
    
    Do
        AllSide(X) = 1
        X = X + 1
    Loop While X <> -1
    
End Sub

Private Sub Form_Load()
    Randomize
    frmOptions.intCurrentLevel = 1000
    
End Sub
Sub GenerateFirstBlock()
Dim RandomBlock As Byte
Dim RandomColor As Byte

    RandomBlock = Int(16 * Rnd) + 1
    RandomColor = Int(7 * Rnd) + 1
    
    If RandomColor = 1 Then
        PreBlockColor = vbRed
    ElseIf RandomColor = 2 Then
        PreBlockColor = vbYellow
    ElseIf RandomColor = 3 Then
        PreBlockColor = vbGreen
    ElseIf RandomColor = 4 Then
        PreBlockColor = vbWhite
    ElseIf RandomColor = 5 Then
        PreBlockColor = vbBlue
    ElseIf RandomColor = 6 Then
        PreBlockColor = &H80FF&
    ElseIf RandomColor = 7 Then
        PreBlockColor = &HFFFF00
    End If
    
    If RandomBlock = 1 Or RandomBlock = 2 Then
        PreBlockType = SQUAREBLOCK
        strPreBlockType = "Square Block"
    ElseIf RandomBlock = 3 Or RandomBlock = 4 Then
        PreBlockType = TBLOCK
        strPreBlockType = "T-Block"
    ElseIf RandomBlock = 5 Or RandomBlock = 6 Then
        PreBlockType = LINEBLOCK
        strPreBlockType = "Line Block"
    ElseIf RandomBlock = 7 Or RandomBlock = 8 Then
        PreBlockType = RIGHTL
        strPreBlockType = "Right-L Block"
    ElseIf RandomBlock = 9 Or RandomBlock = 10 Then
        PreBlockType = LEFTL
        strPreBlockType = "Left-L Block"
    ElseIf RandomBlock = 11 Or RandomBlock = 12 Then
        PreBlockType = SBLOCK
        strPreBlockType = "S-Block"
    ElseIf RandomBlock = 13 Or RandomBlock = 14 Then
        PreBlockType = ZBLOCK
        strPreBlockType = "Z-Block"
    ElseIf RandomBlock = 15 Then
        PreBlockType = SCATTERBLOCK
        strPreBlockType = "Scatter Block!"
    ElseIf RandomBlock = 16 Then
        strPreBlockType = "Bomb Block!"
    End If
    
End Sub
Sub GenerateNewBlock()
Dim RandomBlock As Byte
Dim RandomColor As Byte
If GameEnabled = True Then
    
    RandomBlock = Int(9 * Rnd) + 1
    RandomColor = Int(7 * Rnd) + 1
    BlockMode = 1
    
    BlockColor = PreBlockColor
    bytBlockType = PreBlockType
    
    If RandomColor = 1 Then
        PreBlockColor = vbRed
    ElseIf RandomColor = 2 Then
        PreBlockColor = vbYellow
    ElseIf RandomColor = 3 Then
        PreBlockColor = vbGreen
    ElseIf RandomColor = 4 Then
        PreBlockColor = vbWhite
    ElseIf RandomColor = 5 Then
        PreBlockColor = vbBlue
    ElseIf RandomColor = 6 Then
        PreBlockColor = &H80FF&
    ElseIf RandomColor = 7 Then
        PreBlockColor = &HFFFF00
    End If
    
    If RandomBlock = 1 Then
        PreBlockType = SQUAREBLOCK
        strPreBlockType = "Square Block"
    ElseIf RandomBlock = 2 Then
        PreBlockType = TBLOCK
        strPreBlockType = "T-Block"
    ElseIf RandomBlock = 3 Then
        PreBlockType = LINEBLOCK
        strPreBlockType = "Line Block"
    ElseIf RandomBlock = 4 Then
       PreBlockType = RIGHTL
       strPreBlockType = "Right-L Block"
    ElseIf RandomBlock = 5 Then
        PreBlockType = LEFTL
        strPreBlockType = "Left-L Block"
    ElseIf RandomBlock = 6 Then
        PreBlockType = SBLOCK
        strPreBlockType = "S-Block"
    ElseIf RandomBlock = 7 Then
       PreBlockType = ZBLOCK
       strPreBlockType = "Z-Block"
    ElseIf RandomBlock = 8 Then
        PreBlockType = SCATTERBLOCK
        strPreBlockType = "Scatter Block!"
    ElseIf RandomBlock = 9 Then
        PreBlockType = BOMBBLOCK
        strPreBlockType = "Bomb Block!"
    End If
    
    Call CreateNewPreview(PreBlockType)
    Call CreateNewBlock(bytBlockType)
End If
End Sub
Sub CreateNewPreview(BlockType As Byte)
Dim X As Integer

    If BlockType = SQUAREBLOCK Or BlockType = SCATTERBLOCK Or BlockType = BOMBBLOCK Then
        PreBlock1Pos = 5
        PreBlock2Pos = 6
        PreBlock3Pos = 9
        PreBlock4Pos = 10
    ElseIf BlockType = TBLOCK Then
        PreBlock1Pos = 9
        PreBlock2Pos = 10
        PreBlock3Pos = 11
        PreBlock4Pos = 6
    ElseIf BlockType = LINEBLOCK Then
        PreBlock1Pos = 8
        PreBlock2Pos = 9
        PreBlock3Pos = 10
        PreBlock4Pos = 11
    ElseIf BlockType = RIGHTL Then
        PreBlock1Pos = 13
        PreBlock2Pos = 9
        PreBlock3Pos = 5
        PreBlock4Pos = 6
    ElseIf BlockType = LEFTL Then
        PreBlock1Pos = 14
        PreBlock2Pos = 10
        PreBlock3Pos = 6
        PreBlock4Pos = 5
    ElseIf BlockType = SBLOCK Then
        PreBlock1Pos = 9
        PreBlock2Pos = 10
        PreBlock3Pos = 6
        PreBlock4Pos = 7
    ElseIf BlockType = ZBLOCK Then
        PreBlock1Pos = 10
        PreBlock2Pos = 11
        PreBlock3Pos = 5
        PreBlock4Pos = 6
    End If
    
    For X = 0 To 15
        shpPreview.Item(X).FillColor = vbBlack
    Next X
    
    shpPreview.Item(PreBlock1Pos).FillColor = PreBlockColor
    shpPreview.Item(PreBlock2Pos).FillColor = PreBlockColor
    shpPreview.Item(PreBlock3Pos).FillColor = PreBlockColor
    shpPreview.Item(PreBlock4Pos).FillColor = PreBlockColor

End Sub
Sub CreateNewBlock(BlockType As Byte)
    If BlockType = SQUAREBLOCK Or BlockType = SCATTERBLOCK Or BlockType = BOMBBLOCK Then
        Block1Pos = 358
        Block2Pos = 359
        Block3Pos = 374
        Block4Pos = 375
    ElseIf BlockType = TBLOCK Then
        Block1Pos = 358
        Block2Pos = 359
        Block3Pos = 360
        Block4Pos = 375
    ElseIf BlockType = LINEBLOCK Then
        Block1Pos = 374
        Block2Pos = 375
        Block3Pos = 376
        Block4Pos = 377
    ElseIf BlockType = RIGHTL Then
        Block1Pos = 342
        Block2Pos = 358
        Block3Pos = 374
        Block4Pos = 375
    ElseIf BlockType = LEFTL Then
        Block1Pos = 344
        Block2Pos = 360
        Block3Pos = 375
        Block4Pos = 376
    ElseIf BlockType = SBLOCK Then
        Block1Pos = 359
        Block2Pos = 360
        Block3Pos = 376
        Block4Pos = 377
    ElseIf BlockType = ZBLOCK Then
        Block1Pos = 360
        Block2Pos = 361
        Block3Pos = 375
        Block4Pos = 376
    End If
    lblPreBlockType.Caption = strPreBlockType
    BlockMode = 1
    Call MoveDown
End Sub
Sub DisableUsedBlocks(Pos1 As Integer, Pos2 As Integer, Pos3 As Integer, Pos4 As Integer)
    
    If bytBlockType = SCATTERBLOCK Then
        shpBlock.Item(Block1Pos).FillColor = vbBlack
        shpBlock.Item(Block2Pos).FillColor = vbBlack
        shpBlock.Item(Block3Pos).FillColor = vbBlack
        shpBlock.Item(Block4Pos).FillColor = vbBlack
        On Error GoTo scatterdebug
        Pos1 = Pos1 - Int(3 * Rnd) + 1
        Pos2 = Pos2 - Int(4 * Rnd) + 2
        Pos3 = Pos3 - Int(5 * Rnd) + 3
        Pos4 = Pos4 - Int(6 * Rnd) + 1
scatterdebug:
    If Err.Number = 341 Then
        Resume Next
    End If
        shpBlock.Item(Pos1).FillColor = BlockColor
        shpBlock.Item(Pos2).FillColor = BlockColor
        shpBlock.Item(Pos3).FillColor = BlockColor
        shpBlock.Item(Pos4).FillColor = BlockColor
        
        AllSide(Pos1) = 1
        AllSide(Pos2) = 1
        AllSide(Pos3) = 1
        AllSide(Pos4) = 1
    ElseIf bytBlockType = BOMBBLOCK Then
        On Error GoTo bombblockdebug
        shpBlock.Item(Block1Pos).FillColor = vbBlack
        shpBlock.Item(Block2Pos).FillColor = vbBlack
        shpBlock.Item(Block3Pos).FillColor = vbBlack
        shpBlock.Item(Block4Pos).FillColor = vbBlack
        shpBlock.Item(Block1Pos - 17).FillColor = vbBlack
        shpBlock.Item(Block1Pos - 16).FillColor = vbBlack
        shpBlock.Item(Block2Pos - 16).FillColor = vbBlack
        shpBlock.Item(Block2Pos - 15).FillColor = vbBlack
        shpBlock.Item(Block2Pos + 1).FillColor = vbBlack
        shpBlock.Item(Block4Pos + 1).FillColor = vbBlack
        shpBlock.Item(Block4Pos + 17).FillColor = vbBlack
        shpBlock.Item(Block4Pos + 16).FillColor = vbBlack
        shpBlock.Item(Block3Pos + 16).FillColor = vbBlack
        shpBlock.Item(Block3Pos + 15).FillColor = vbBlack
        shpBlock.Item(Block3Pos - 1).FillColor = vbBlack
        shpBlock.Item(Block1Pos - 1).FillColor = vbBlack
bombblockdebug:
    If Err.Number = 341 Then
        Resume Next
    End If
        AllSide(Block1Pos) = 0
        AllSide(Block2Pos) = 0
        AllSide(Block3Pos) = 0
        AllSide(Block4Pos) = 0
        AllSide(Block1Pos - 17) = 0
        AllSide(Block1Pos - 16) = 0
        AllSide(Block2Pos - 16) = 0
        AllSide(Block2Pos - 15) = 0
        AllSide(Block2Pos + 1) = 0
        AllSide(Block4Pos + 1) = 0
        AllSide(Block4Pos + 17) = 0
        AllSide(Block4Pos + 16) = 0
        AllSide(Block3Pos + 16) = 0
        AllSide(Block3Pos + 15) = 0
        AllSide(Block3Pos - 1) = 0
        AllSide(Block1Pos - 1) = 0
    Else
        AllSide(Pos1) = 1
        AllSide(Pos2) = 1
        AllSide(Pos3) = 1
        AllSide(Pos4) = 1
    End If
    
    
    intScore = intScore + Int(20 * Rnd) + 5
    lblScore.Caption = intScore
    
End Sub
Sub EliminateLine(intRow As Integer)
Dim Temp(0 To 399) As Variant
Dim X As Integer
    
    For X = intRow To 383
        Temp(X) = AllSide(X + 16)
    Next X
    
    For X = intRow To 383
        AllSide(X) = Temp(X)
    Next X
    
    For X = intRow To 367
        Temp(X) = shpBlock(X + 16).FillColor
    Next X
    
    For X = intRow To 383
        shpBlock(X).FillColor = Temp(X)
    Next X
    
    intScore = intScore + Int(250 * Rnd) + 100
    lblScore.Caption = intScore
    Call LineCheck
End Sub
Sub LineCheck()
Dim Temp As Integer
Dim X As Integer
Dim LineRow As Integer

    X = -16
    
    Do
        'Get each row
        X = X + 16
        If AllSide(X) = 1 And AllSide(X + 1) = 1 And AllSide(X + 2) = 1 And AllSide(X + 3) = 1 And AllSide(X + 4) = 1 And AllSide(X + 5) = 1 And AllSide(X + 6) = 1 And AllSide(X + 7) = 1 And AllSide(X + 8) = 1 And AllSide(X + 9) = 1 And AllSide(X + 10) = 1 And AllSide(X + 11) = 1 And AllSide(X + 12) = 1 And AllSide(X + 13) = 1 And AllSide(X + 14) = 1 And AllSide(X + 15) = 1 Then
            Call EliminateLine(X)
            Exit Sub
        End If
    Loop While X < 368

End Sub
Function MoveDownCheck() As Boolean

    If (Block1Pos - 16 < 0) Or (Block2Pos - 16 < 0) Or (Block3Pos - 16 < 0) Or (Block4Pos - 16 < 0) Or _
        (AllSide(Block1Pos - 16) = 1 Or AllSide(Block2Pos - 16) = 1 Or AllSide(Block3Pos - 16) = 1 Or AllSide(Block4Pos - 16) = 1) Then
            MoveDownCheck = False
    Else
        MoveDownCheck = True
    End If
    
End Function
Sub DoGameOver()
Dim X As Integer
    MsgBox ("You have lost Blocks!!!"), vbExclamation, "GAME OVER"
    tmrDescend.Enabled = False
    strPreBlockType = "None"
    lblPreBlockType.Caption = strPreBlockType
    GameEnabled = False
    For X = 0 To 383
        shpBlock.Item(X).FillColor = vbBlack
    Next X
    For X = 0 To 15
        shpPreview.Item(X).FillColor = vbBlack
    Next X
    intScore = 0
    lblScore.Caption = intScore
    Me.SetFocus
    
End Sub
Sub MoveDown()

    On Error GoTo gameover
    If bytBlockType = SQUAREBLOCK Then
        If AllSide(358) = 1 Or AllSide(359) = 1 Or AllSide(374) = 1 Or AllSide(375) = 1 Then
            Call DoGameOver
            Exit Sub
        End If
    ElseIf bytBlockType = TBLOCK Then
        If AllSide(358) = 1 Or AllSide(359) = 1 Or AllSide(360) = 1 Or AllSide(375) = 1 Then
            Call DoGameOver
            Exit Sub
        End If
    ElseIf bytBlockType = RIGHTL Then
        If AllSide(342) Or AllSide(358) = 1 Or AllSide(374) = 1 Or AllSide(375) = 1 Then
            Call DoGameOver
            Exit Sub
        End If
    ElseIf bytBlockType = LEFTL Then
        If AllSide(344) = 1 Or AllSide(360) = 1 Or AllSide(375) = 1 Or AllSide(376) = 1 Then
            Call DoGameOver
            Exit Sub
        End If
    ElseIf bytBlockType = SBLOCK Then
        If AllSide(359) = 1 Or AllSide(360) = 1 Or AllSide(376) = 1 Or AllSide(377) = 1 Then
            Call DoGameOver
            Exit Sub
        End If
    ElseIf bytBlockType = ZBLOCK Then
        If AllSide(360) = 1 Or AllSide(361) = 1 Or AllSide(375) = 1 Or AllSide(376) = 1 Then
            Call DoGameOver
            Exit Sub
        End If
    End If
gameover:
        If Err.Number = 9 Then
            Resume Next
        End If

    If MoveDownCheck = True Then
        shpBlock.Item(Block1Pos).FillColor = vbBlack
        shpBlock.Item(Block2Pos).FillColor = vbBlack
        shpBlock.Item(Block3Pos).FillColor = vbBlack
        shpBlock.Item(Block4Pos).FillColor = vbBlack
    
        Block1Pos = Block1Pos - 16
        Block2Pos = Block2Pos - 16
        Block3Pos = Block3Pos - 16
        Block4Pos = Block4Pos - 16
        
        shpBlock.Item(Block1Pos).FillColor = BlockColor
        shpBlock.Item(Block2Pos).FillColor = BlockColor
        shpBlock.Item(Block3Pos).FillColor = BlockColor
        shpBlock.Item(Block4Pos).FillColor = BlockColor
    Else
        Call DisableUsedBlocks(Block1Pos, Block2Pos, Block3Pos, Block4Pos)
        Call LineCheck
        Call GenerateNewBlock
    End If
End Sub
Function MoveRightCheck() As Boolean
    If (RightSide(Block1Pos) = 1 Or RightSide(Block2Pos) = 1 Or RightSide(Block3Pos) = 1 Or RightSide(Block4Pos) = 1) Or _
        (AllSide(Block1Pos + 1) = 1 Or AllSide(Block2Pos + 1) = 1 Or AllSide(Block3Pos + 1) = 1 Or AllSide(Block4Pos + 1) = 1) Then
        MoveRightCheck = False
    Else
        MoveRightCheck = True
    End If
End Function
Sub MoveRight()
    If MoveRightCheck = True Then
        shpBlock.Item(Block1Pos).FillColor = vbBlack
        shpBlock.Item(Block2Pos).FillColor = vbBlack
        shpBlock.Item(Block3Pos).FillColor = vbBlack
        shpBlock.Item(Block4Pos).FillColor = vbBlack
        
        Block1Pos = Block1Pos + 1
        Block2Pos = Block2Pos + 1
        Block3Pos = Block3Pos + 1
        Block4Pos = Block4Pos + 1
        
        shpBlock.Item(Block1Pos).FillColor = BlockColor
        shpBlock.Item(Block2Pos).FillColor = BlockColor
        shpBlock.Item(Block3Pos).FillColor = BlockColor
        shpBlock.Item(Block4Pos).FillColor = BlockColor
    End If
End Sub
Function MoveLeftCheck() As Boolean
    If (LeftSide(Block1Pos) = 1 Or LeftSide(Block2Pos) = 1 Or LeftSide(Block3Pos) = 1 Or LeftSide(Block4Pos) = 1) Or _
        (AllSide(Block1Pos - 1) = 1 Or AllSide(Block2Pos - 1) = 1 Or AllSide(Block3Pos - 1) = 1 Or AllSide(Block4Pos - 1) = 1) Then
        MoveLeftCheck = False
    Else
        MoveLeftCheck = True
    End If
End Function

Sub MoveLeft()
    If MoveLeftCheck = True Then
        shpBlock.Item(Block1Pos).FillColor = vbBlack
        shpBlock.Item(Block2Pos).FillColor = vbBlack
        shpBlock.Item(Block3Pos).FillColor = vbBlack
        shpBlock.Item(Block4Pos).FillColor = vbBlack
        
        Block1Pos = Block1Pos - 1
        Block2Pos = Block2Pos - 1
        Block3Pos = Block3Pos - 1
        Block4Pos = Block4Pos - 1
        
        shpBlock.Item(Block1Pos).FillColor = BlockColor
        shpBlock.Item(Block2Pos).FillColor = BlockColor
        shpBlock.Item(Block3Pos).FillColor = BlockColor
        shpBlock.Item(Block4Pos).FillColor = BlockColor
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOptions.ForeColor = vbGreen
    lblNewGame.ForeColor = vbGreen
    lblExit.ForeColor = vbGreen
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpExit.FillColor = vbBlue
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblExit.ForeColor = vbRed
    lblOptions.ForeColor = vbGreen
    lblNewGame.ForeColor = vbGreen
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpExit.FillColor = vbBlack
End Sub

Private Sub lblNewGame_Click()
    Call NewGame
End Sub

Private Sub lblNewGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpNewGame.FillColor = vbBlue
End Sub

Private Sub lblNewGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblNewGame.ForeColor = vbRed
    lblOptions.ForeColor = vbGreen
    lblExit.ForeColor = vbGreen
End Sub

Private Sub lblNewGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpNewGame.FillColor = vbBlack
End Sub

Private Sub lblOptions_Click()
    tmrDescend.Enabled = False
    frmOptions.Show vbModal
End Sub

Private Sub lblOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpOptions.FillColor = vbBlue
End Sub

Private Sub lblOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOptions.ForeColor = vbRed
    lblExit.ForeColor = vbGreen
    lblNewGame.ForeColor = vbGreen
End Sub

Private Sub lblOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpOptions.FillColor = vbBlack
End Sub

Private Sub tmrDescend_Timer()
    Call MoveDown
End Sub

Private Sub tmrSpecialColors_Timer()
Dim bytRandomColor As Byte, bytPreBlockColor As Variant
Dim intBlockColor As Variant

    If PreBlockType = SCATTERBLOCK Or PreBlockType = BOMBBLOCK Then
        bytRandomColor = Int(7 * Rnd) + 1
        
        If bytRandomColor = 1 Then
            bytPreBlockColor = vbRed
        ElseIf bytRandomColor = 2 Then
            bytPreBlockColor = vbYellow
        ElseIf bytRandomColor = 3 Then
            bytPreBlockColor = vbGreen
        ElseIf bytRandomColor = 4 Then
            bytPreBlockColor = vbWhite
        ElseIf bytRandomColor = 5 Then
            bytPreBlockColor = vbBlue
        ElseIf bytRandomColor = 6 Then
            bytPreBlockColor = &H80FF&
        ElseIf bytRandomColor = 7 Then
            bytPreBlockColor = &HFFFF00
        End If
    
     shpPreview.Item(5).FillColor = bytPreBlockColor
     shpPreview.Item(6).FillColor = bytPreBlockColor
     shpPreview.Item(9).FillColor = bytPreBlockColor
     shpPreview.Item(10).FillColor = bytPreBlockColor
    End If
    If bytBlockType = SCATTERBLOCK Or bytBlockType = BOMBBLOCK Then
        bytRandomColor = Int(7 * Rnd) + 1
    
        If bytRandomColor = 1 Then
            intBlockColor = vbRed
        ElseIf bytRandomColor = 2 Then
            intBlockColor = vbYellow
        ElseIf bytRandomColor = 3 Then
            intBlockColor = vbGreen
        ElseIf bytRandomColor = 4 Then
            intBlockColor = vbWhite
        ElseIf bytRandomColor = 5 Then
            intBlockColor = vbBlue
        ElseIf bytRandomColor = 6 Then
            intBlockColor = &H80FF&
        ElseIf bytRandomColor = 7 Then
            intBlockColor = &HFFFF00
        End If
        
        shpBlock.Item(Block1Pos).FillColor = intBlockColor
        shpBlock.Item(Block2Pos).FillColor = intBlockColor
        shpBlock.Item(Block3Pos).FillColor = intBlockColor
        shpBlock.Item(Block4Pos).FillColor = intBlockColor

    End If
    
End Sub
