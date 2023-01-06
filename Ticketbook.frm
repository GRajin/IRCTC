VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Ticketbook 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9375
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
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   138870785
      CurrentDate     =   42762
      MaxDate         =   42855
      MinDate         =   42762
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLICK"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text4 
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
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Ticketbook.frx":0000
      Left            =   2760
      List            =   "Ticketbook.frx":0010
      TabIndex        =   5
      Top             =   2640
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Ticketbook.frx":0035
      Left            =   2760
      List            =   "Ticketbook.frx":0045
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "NO.OF PASSENGERS"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "INDIAN RAILWAYS CATERING AND TOURISM CORPORATION LIMITED"
      BeginProperty Font 
         Name            =   "CabinSketch"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Ticketbook.frx":006A
      Top             =   240
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      Height          =   5295
      Left            =   120
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Ticketbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Details.Show
Ticketbook.Hide
Select Case Ticketbook.Text4.Text

Case 1
Details.Label10.Enabled = True
Details.Label10.Caption = "1"
Details.Text1.Enabled = True
Details.Text5.Enabled = True
Details.Combo1.Enabled = True
Details.Check1.Enabled = True

Case 2
Details.Label10.Enabled = True
Details.Label10.Caption = "1"
Details.Text1.Enabled = True
Details.Text5.Enabled = True
Details.Combo1.Enabled = True
Details.Check1.Enabled = True
Details.Label11.Enabled = True
Details.Label11.Caption = "2"
Details.Text2.Enabled = True
Details.Text6.Enabled = True
Details.Combo2.Enabled = True
Details.Check2.Enabled = True

Case 3
Details.Label10.Enabled = True
Details.Label10.Caption = "1"
Details.Text1.Enabled = True
Details.Text5.Enabled = True
Details.Combo1.Enabled = True
Details.Check1.Enabled = True
Details.Label11.Enabled = True
Details.Label11.Caption = "2"
Details.Text2.Enabled = True
Details.Text6.Enabled = True
Details.Combo2.Enabled = True
Details.Check2.Enabled = True
Details.Label12.Enabled = True
Details.Label12.Caption = "3"
Details.Text3.Enabled = True
Details.Text7.Enabled = True
Details.Combo3.Enabled = True
Details.Check3.Enabled = True

Case 4
Details.Label10.Enabled = True
Details.Label10.Caption = "1"
Details.Text1.Enabled = True
Details.Text5.Enabled = True
Details.Combo1.Enabled = True
Details.Check1.Enabled = True
Details.Label11.Enabled = True
Details.Label11.Caption = "2"
Details.Text2.Enabled = True
Details.Text6.Enabled = True
Details.Combo2.Enabled = True
Details.Check2.Enabled = True
Details.Label12.Enabled = True
Details.Label12.Caption = "3"
Details.Text3.Enabled = True
Details.Text7.Enabled = True
Details.Combo3.Enabled = True
Details.Check3.Enabled = True
Details.Label13.Enabled = True
Details.Label13.Caption = "4"
Details.Text4.Enabled = True
Details.Text8.Enabled = True
Details.Combo4.Enabled = True
Details.Check4.Enabled = True

Case Else
MsgBox ("You Are Not Allowed To Book More Than 4 Tickets"), vbCritical
Ticketbook.Show

End Select

If (Project1.Ticketbook.Combo1.Text = "Chennai") And (Project1.Ticketbook.Combo2.Text = "Mumbai") Then
Project1.Details.F1.Caption = 11042
Project1.Details.F2.Caption = "Mumbai Express"
Project1.Details.F3.Caption = "Chennai"
Project1.Details.F4.Caption = "Mumbai"

ElseIf (Project1.Ticketbook.Combo1.Text = "Chennai") And (Project1.Ticketbook.Combo2.Text = "Kolkata") Then
Project1.Details.F1.Caption = 12840
Project1.Details.F2.Caption = "Howrah Mail"
Project1.Details.F3.Caption = "Chennai"
Project1.Details.F4.Caption = "Kolkata"

ElseIf (Project1.Ticketbook.Combo1.Text = "Chennai") And (Project1.Ticketbook.Combo2.Text = "Delhi") Then
Project1.Details.F1.Caption = 12431
Project1.Details.F2.Caption = "Rajdhani Express"
Project1.Details.F3.Caption = "Chennai"
Project1.Details.F4.Caption = "Delhi"

ElseIf (Project1.Ticketbook.Combo1.Text = "Mumbai") And (Project1.Ticketbook.Combo2.Text = "Chennai") Then
Project1.Details.F1.Caption = 11027
Project1.Details.F2.Caption = "Chennai Mail"
Project1.Details.F3.Caption = "Mumbai"
Project1.Details.F4.Caption = "Chennai"

ElseIf (Project1.Ticketbook.Combo1.Text = "Mumbai") And (Project1.Ticketbook.Combo2.Text = "Kolkata") Then
Project1.Details.F1.Caption = 12274
Project1.Details.F2.Caption = "Howrah Duronto"
Project1.Details.F3.Caption = "Mumbai"
Project1.Details.F4.Caption = "Kolkata"

ElseIf (Project1.Ticketbook.Combo1.Text = "Mumbai") And (Project1.Ticketbook.Combo2.Text = "Delhi") Then
Project1.Details.F1.Caption = 12472
Project1.Details.F2.Caption = "Swaraj Express"
Project1.Details.F3.Caption = "Mumbai"
Project1.Details.F4.Caption = "Delhi"

ElseIf (Project1.Ticketbook.Combo1.Text = "Kolkata") And (Project1.Ticketbook.Combo2.Text = "Chennai") Then
Project1.Details.F1.Caption = 12841
Project1.Details.F2.Caption = "Coromandal Express"
Project1.Details.F3.Caption = "Kolkata"
Project1.Details.F4.Caption = "Chennai"

ElseIf (Project1.Ticketbook.Combo1.Text = "Kolkata") And (Project1.Ticketbook.Combo2.Text = "Delhi") Then
Project1.Details.F1.Caption = 12305
Project1.Details.F2.Caption = "Kolkata Rajadhani"
Project1.Details.F3.Caption = "Kolkata"
Project1.Details.F4.Caption = "Delhi"

ElseIf (Project1.Ticketbook.Combo1.Text = "Kolkata") And (Project1.Ticketbook.Combo2.Text = "Mumbai") Then
Project1.Details.F1.Caption = 12860
Project1.Details.F2.Caption = "Gitanjali Express"
Project1.Details.F3.Caption = "Kolkata"
Project1.Details.F4.Caption = "Mumbai"

ElseIf (Project1.Ticketbook.Combo1.Text = "Delhi") And (Project1.Ticketbook.Combo2.Text = "Chennai") Then
Project1.Details.F1.Caption = 12434
Project1.Details.F2.Caption = "Chennai Rajadhani"
Project1.Details.F3.Caption = "Delhi"
Project1.Details.F4.Caption = "Chennai"

ElseIf (Project1.Ticketbook.Combo1.Text = "Delhi") And (Project1.Ticketbook.Combo2.Text = "Kolkata") Then
Project1.Details.F1.Caption = 12302
Project1.Details.F2.Caption = "Kolkata Rajadhani"
Project1.Details.F3.Caption = "Delhi"
Project1.Details.F4.Caption = "Kolkata"

ElseIf (Project1.Ticketbook.Combo1.Text = "Delhi") And (Project1.Ticketbook.Combo2.Text = "Mumbai") Then
Project1.Details.F1.Caption = 12618
Project1.Details.F2.Caption = "Mangalore Lakshadeep Express"
Project1.Details.F3.Caption = "Delhi"
Project1.Details.F4.Caption = "Mumbai"

End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case vbKey0 To vbKey9

Case vbKeyBack, vbKeyClear, vbKeyDelete

Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab

Case Else
KeyAscii = 0

End Select

End Sub
