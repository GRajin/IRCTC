VERSION 5.00
Begin VB.Form Account 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "CREATE"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click To Register"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click To Login"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Account.frx":0000
      Top             =   240
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   4095
      Left            =   120
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "CLICK TO CREATE A NEW ACCOUNT"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "CLICK TO LOGIN INTO YOUR ACCOUNT"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Login.Show
Account.Hide

End Sub

Private Sub Command2_Click()

register.Show
Account.Hide

End Sub

Private Sub Command3_Click()

End

End Sub

