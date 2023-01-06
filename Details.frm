VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Details 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "CabinSketch"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   11295
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
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Details1ado 
      Height          =   375
      Left            =   600
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "Table3"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "CabinSketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Detailsado 
      Height          =   375
      Left            =   480
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "Table2"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "CabinSketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Proceed"
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   465
      Left            =   9480
      TabIndex        =   35
      Top             =   7680
      Width           =   210
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   465
      Left            =   9480
      TabIndex        =   34
      Top             =   7080
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   465
      Left            =   9480
      TabIndex        =   33
      Top             =   6480
      Width           =   210
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   465
      Left            =   9480
      TabIndex        =   32
      Top             =   5880
      Width           =   210
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Details.frx":0000
      Left            =   6240
      List            =   "Details.frx":000A
      TabIndex        =   29
      Top             =   7680
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Details.frx":001C
      Left            =   6240
      List            =   "Details.frx":0026
      TabIndex        =   28
      Top             =   7080
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Details.frx":0038
      Left            =   6240
      List            =   "Details.frx":0042
      TabIndex        =   27
      Top             =   6480
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Details.frx":0054
      Left            =   6240
      List            =   "Details.frx":005E
      TabIndex        =   26
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Passenger Details"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   10695
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         TabIndex        =   24
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         TabIndex        =   23
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         TabIndex        =   22
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5160
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   18
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   17
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         TabIndex        =   16
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1200
         TabIndex        =   15
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Senior Citizen"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   31
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   30
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "S No"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Journey Details"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   10695
      Begin VB.Label F4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   9
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label F3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label F2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   5
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Train Name"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label F1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Train No"
         BeginProperty Font 
            Name            =   "CabinSketch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "INDIAN RAILWAYS CATERING AND TOURISM CORPORATION LIMITED"
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Details.frx":0070
      Top             =   240
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      Height          =   9255
      Left            =   120
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Ticketbook.Show
Details.Hide


End Sub

Private Sub Command2_Click()

If (Text1.Enabled = True) And (Text1.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text2.Enabled = True) And (Text2.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text3.Enabled = True) And (Text3.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text4.Enabled = True) And (Text4.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text5.Enabled = True) And (Text5.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text6.Enabled = True) And (Text6.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text7.Enabled = True) And (Text7.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Text8.Enabled = True) And (Text8.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Combo1.Enabled = True) And (Combo1.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Combo2.Enabled = True) And (Combo2.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Combo3.Enabled = True) And (Combo3.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

ElseIf (Combo4.Enabled = True) And (Combo4.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical

Else
Booking.Show
Details.Hide

End If

Dim IntResult As Integer
Randomize
IntResult = Int((1000 * Rnd) + 10000)
Booking.L1.Caption = IntResult

End Sub
