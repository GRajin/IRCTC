VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Booking 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   360
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Ticketado 
      Height          =   330
      Left            =   600
      Top             =   840
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
      RecordSource    =   "Table4"
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
      BackColor       =   &H00FFC0C0&
      Caption         =   "Proceed For Payment"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6015
      Left            =   120
      Top             =   120
      Width           =   10815
   End
   Begin VB.Shape Shape2 
      Height          =   3135
      Left            =   240
      Top             =   1800
      Width           =   10575
   End
   Begin VB.Label L6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Train Name"
      DataSource      =   "Ticketado"
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
      Left            =   8040
      TabIndex        =   12
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label L5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Train No"
      DataSource      =   "Ticketado"
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
      Left            =   8040
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "TRAIN NAME"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "TRAIN NO"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label L4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Date Of Journey"
      DataSource      =   "Ticketado"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "DATE"
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
      Left            =   480
      TabIndex        =   7
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label L3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "To"
      DataSource      =   "Ticketado"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "TO"
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
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "From"
      DataSource      =   "Ticketado"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "FROM"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      DataField       =   "PNR Number"
      DataSource      =   "Ticketado"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "PNR NUMBER"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
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
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Booking.frx":0000
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "Booking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Booking.Hide
Payment.Show

Project1.Details.Detailsado.Refresh
Project1.Details.Detailsado.Recordset.AddNew
Project1.Details.Detailsado.Recordset.Fields("Train No").Value = Project1.Details.F1.Caption
Project1.Details.Detailsado.Recordset.Fields("Train Name").Value = Project1.Details.F2.Caption
Project1.Details.Detailsado.Recordset.Fields("From").Value = Project1.Details.F3.Caption
Project1.Details.Detailsado.Recordset.Fields("To").Value = Project1.Details.F4.Caption
Project1.Details.Detailsado.Recordset.Update

Project1.Details.Details1ado.Refresh
If (Details.Label10.Enabled = True) And (Details.Text1.Enabled = True) And (Details.Text5.Enabled = True) And (Details.Combo1.Enabled = True) And (Details.Check1.Enabled = True) Then
Project1.Details.Details1ado.Recordset.AddNew
Project1.Details.Details1ado.Recordset.Fields("S No").Value = Project1.Details.Label10.Caption
Project1.Details.Details1ado.Recordset.Fields("Name").Value = Project1.Details.Text1.Text
Project1.Details.Details1ado.Recordset.Fields("Age").Value = Project1.Details.Text5.Text
Project1.Details.Details1ado.Recordset.Fields("Gender").Value = Project1.Details.Combo1.Text
If Details.Check1.Value = 1 Then
Project1.Details.Details1ado.Recordset.Fields("Senior Citizen").Value = "Senior Citizen"
End If
Project1.Details.Details1ado.Recordset.Update
End If

If (Details.Label11.Enabled = True) And (Details.Text2.Enabled = True) And (Details.Text6.Enabled = True) And (Details.Combo2.Enabled = True) And (Details.Check2.Enabled = True) Then
Project1.Details.Details1ado.Refresh
Project1.Details.Details1ado.Recordset.AddNew
Project1.Details.Details1ado.Recordset.Fields("S No").Value = Project1.Details.Label11.Caption
Project1.Details.Details1ado.Recordset.Fields("Name").Value = Project1.Details.Text2.Text
Project1.Details.Details1ado.Recordset.Fields("Age").Value = Project1.Details.Text6.Text
Project1.Details.Details1ado.Recordset.Fields("Gender").Value = Project1.Details.Combo2.Text
If Details.Check2.Value = 1 Then
Project1.Details.Details1ado.Recordset.Fields("Senior Citizen").Value = "Senior Citizen"
End If
Project1.Details.Details1ado.Recordset.Update
End If

If (Details.Label12.Enabled = True) And (Details.Text3.Enabled = True) And (Details.Text7.Enabled = True) And (Details.Combo3.Enabled = True) And (Details.Check3.Enabled = True) Then
Project1.Details.Details1ado.Refresh
Project1.Details.Details1ado.Recordset.AddNew
Project1.Details.Details1ado.Recordset.Fields("S No").Value = Project1.Details.Label12.Caption
Project1.Details.Details1ado.Recordset.Fields("Name").Value = Project1.Details.Text3.Text
Project1.Details.Details1ado.Recordset.Fields("Age").Value = Project1.Details.Text7.Text
Project1.Details.Details1ado.Recordset.Fields("Gender").Value = Project1.Details.Combo3.Text
If Details.Check3.Value = 1 Then
Project1.Details.Details1ado.Recordset.Fields("Senior Citizen").Value = "Senior Citizen"
End If
Project1.Details.Details1ado.Recordset.Update
End If

If (Details.Label13.Enabled = True) And (Details.Text4.Enabled = True) And (Details.Text8.Enabled = True) And (Details.Combo4.Enabled = True) And (Details.Check4.Enabled = True) Then
Project1.Details.Details1ado.Refresh
Project1.Details.Details1ado.Recordset.AddNew
Project1.Details.Details1ado.Recordset.Fields("S No").Value = Project1.Details.Label13.Caption
Project1.Details.Details1ado.Recordset.Fields("Name").Value = Project1.Details.Text4.Text
Project1.Details.Details1ado.Recordset.Fields("Age").Value = Project1.Details.Text8.Text
Project1.Details.Details1ado.Recordset.Fields("Gender").Value = Project1.Details.Combo4.Text
If Details.Check4.Value = 1 Then
Project1.Details.Details1ado.Recordset.Fields("Senior Citizen").Value = "Senior Citizen"
End If
Project1.Details.Details1ado.Recordset.Update
End If

End Sub

Private Sub Form_Load()

L2.Caption = Project1.Details.F3.Caption
L3.Caption = Project1.Details.F4.Caption
L4.Caption = Project1.Ticketbook.DTPicker1.Value
L5.Caption = Project1.Details.F1.Caption
L6.Caption = Project1.Details.F2.Caption

End Sub



