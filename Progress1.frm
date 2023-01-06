VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Progress1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4800
      Top             =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "INDIAN RAILWAYS CATERING AND TOURISM CORPORATION LIMITED"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Progress1.frx":0000
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Percentage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   4215
      Left            =   120
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "Progress1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

ProgressBar1.Value = ProgressBar1.Value + 5

Select Case ProgressBar1.Value

Case 0 To 20
Status.Caption = "Loading... Please Wait!"
Percentage.Caption = ProgressBar1.Value & "%"
Timer1.Interval = 200

Case 21 To 50
Status.Caption = "Checking Connections...Please Wait!"
Percentage.Caption = ProgressBar1.Value & "%"
Timer1.Interval = 250

Case 51 To 70
Status.Caption = "Loading Database...Please Wait!"
Percentage.Caption = ProgressBar1.Value & "%"
Timer1.Interval = 300

Case 71 To 90
Status.Caption = "Checking Server...Please Wait!"
Percentage.Caption = ProgressBar1.Value & "%"
Timer1.Interval = 200

Case 91 To 100
Status.Caption = "Loading Other Connection...Please Wait!"
Percentage.Caption = ProgressBar1.Value & "%"
Timer1.Interval = 250

End Select

If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me
Account.Show
End If

End Sub
