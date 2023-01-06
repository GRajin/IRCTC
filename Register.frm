VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form register 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7455
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc registerado 
      Height          =   330
      Left            =   360
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RAJIN\MVB\IRCTC.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RAJIN\MVB\IRCTC.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Close"
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "I Agree To The Terms And Conditions"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   435
      Left            =   1680
      TabIndex        =   14
      Top             =   5400
      Width           =   5655
   End
   Begin VB.CommandButton Cancelbtn 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click To Go Back"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Resetbtn 
      BackColor       =   &H0080FFFF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click To Erase"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Regbtn 
      BackColor       =   &H0000FF00&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Click To Register"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txtuser 
      DataField       =   "Username"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      DataField       =   "Password"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txtadd 
      DataField       =   "Address"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txtphone 
      DataField       =   "Contact"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Register.frx":0000
      Top             =   240
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7215
      Left            =   120
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "REGISTRATION FORM"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelbtn_Click()

Account.Show
register.Hide

End Sub

Private Sub Command1_Click()

End

End Sub

Private Sub Form_Load()

registerado.Recordset.AddNew

End Sub

Private Sub Regbtn_Click()

registerado.Recordset.Fields("Name") = txtname.Text
registerado.Recordset.Fields("Username") = txtuser.Text
registerado.Recordset.Fields("Password") = txtpass.Text
registerado.Recordset.Fields("Address") = txtadd.Text
registerado.Recordset.Fields("Contact") = txtphone.Text

If (txtname.Text = "") Or (txtuser.Text = "") Or (txtpass.Text = "") Or (txtadd.Text = "") Or (txtphone.Text = "") Then
MsgBox ("Please Fill All The Details."), vbCritical
ElseIf Check1.Value = 0 Then
MsgBox ("Please Agree The Terms And Conditions "), vbCritical
Else
registerado.Recordset.Update
MsgBox "Registration Successful", vbInformation
Login.Show
register.Hide
End If

End Sub

Private Sub Resetbtn_Click()

txtname.Text = ""
txtuser.Text = ""
txtpass.Text = ""
txtadd.Text = ""
txtphone.Text = ""

End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case vbKey0 To vbKey9

Case vbKeyBack, vbKeyClear, vbKeyDelete

Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab

Case Else
KeyAscii = 0

End Select

End Sub

