VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   7185
   ClientLeft      =   180
   ClientTop       =   1125
   ClientWidth     =   9600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Options.frx":0000
   ScaleHeight     =   467.317
   ScaleMode       =   0  'User
   ScaleWidth      =   635.039
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "&Speed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "&Credits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<-- &Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "E&xit"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://getpaidtosurf.groovy.nu"
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   5370
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get Paid to Surf the Internet!:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5085
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://mindportal.cjb.net"
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   1575
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "or"
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   2670
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://mindportal.groovy.nu"
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   1425
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MindPortal Entertainment Website:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   795
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programming - Kamran Aslam   Graphics - Kamran Aslam          Beta Testing - Joseph C."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   7185
      Left            =   0
      Top             =   0
      Width           =   9555
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intCurrentLevel As Integer
Private Sub cmdBack_Click()
    
    frmTetris.tmrDescend.Interval = intCurrentLevel
    If frmTetris.GameEnabled = True Then
        frmTetris.tmrDescend.Enabled = True
    End If
    Me.Hide
    
End Sub

Private Sub cmdCredits_Click()
    
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    
End Sub

Private Sub cmdSpeed_Click()
Dim Temp As Integer

    Temp = Val(InputBox("Enter the game speed. (1-10)", "Blocks!"))
       
    If Temp > 10 Or Temp < 1 Then
        MsgBox ("You have entered an invalid speed."), vbCritical, "Bocks!"
        Call cmdSpeed_Click
    Else
        intCurrentLevel = 1100 - (Temp * 100)
    End If
    
End Sub


Private Sub Form_Load()
    Me.Left = (Screen.Width - Width) \ 2
    Me.Top = (Screen.Height - Height) \ 2
End Sub

