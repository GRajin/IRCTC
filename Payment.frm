VERSION 5.00
Begin VB.Form Payment 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "CHECK OUT"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "RECEIVE OTP"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   12
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "VALIDATE"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "Payment.frx":0000
      Left            =   4920
      List            =   "Payment.frx":001F
      TabIndex        =   8
      Top             =   2520
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
      ItemData        =   "Payment.frx":0059
      Left            =   3480
      List            =   "Payment.frx":0081
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
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
      Left            =   2280
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
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
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Left            =   3480
      TabIndex        =   3
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox Text1 
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
      Left            =   3480
      MaxLength       =   16
      TabIndex        =   2
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "ENTER OTP"
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
      Left            =   360
      TabIndex        =   13
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "CVN NO."
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
      Left            =   360
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "CARD HOLDER NAME"
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
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "EXPIRY DATE"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "CARD NUMBER"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   120
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
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
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "Payment.frx":00B5
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

MsgBox ("Your OTP Number Is 252246")
If IsNull(Text1) Then
MsgBox ("Please Enter Your Account Number"), vbCritical
End If

End Sub

Private Sub Command1_Click()
Dim vbNumber As String
Dim vbInstr As Integer
Dim vbTemp As String
Dim vbNumber2 As String

vbNumber = Text1.Text
vbInstr = 1
vbTemp = ""

While vbInstr > 0
vbInstr = InStr(vbNumber, "-")

If vbInstr > 0 Then
vbNumber2 = Left$(vbNumber, vbInstr - 1)
Else
vbNumber2 = vbNumber
End If

Wend

If Len(vbTemp) > 1 Then
vbNumber = vbTemp
End If

Select Case Left$(vbNumber, 1)

Case "4"
Picture1.Picture = LoadPicture("D:\RAJIN\MVB\IRCTC\Visa.gif")

Case "5"
Picture1.Picture = LoadPicture("D:\RAJIN\MVB\IRCTC\Mcard.gif")

End Select

End Sub

Private Sub Command3_Click()

Project1.Booking.Ticketado.Recordset.AddNew
Project1.Booking.Ticketado.Recordset.Fields("PNR Number") = Project1.Booking.L1.Caption
Project1.Booking.Ticketado.Recordset.Fields("From") = Project1.Booking.L2.Caption
Project1.Booking.Ticketado.Recordset.Fields("To") = Project1.Booking.L3.Caption
Project1.Booking.Ticketado.Recordset.Fields("Date Of Journey") = Project1.Booking.L4.Caption
Project1.Booking.Ticketado.Recordset.Fields("Train No") = Project1.Booking.L5.Caption
Project1.Booking.Ticketado.Recordset.Fields("Train Name") = Project1.Booking.L5.Caption

If Text4.Text = 252246 Then
Project1.Booking.Ticketado.Recordset.Update
MsgBox ("You Have Booked Your Ticket."), vbInformation
Divert.Show
Payment.Hide
ElseIf Text4.Text = "" Then
MsgBox ("Please Fill The Correct OTP"), vbCritical
ElseIf Text4.Text <> 252246 Then
MsgBox ("Please Fill The Correct OTP")
ElseIf (Text1.Text = "") Or (Combo2.Text = "") Or (Combo3.Text = "") Or (Text2.Text = "") Or (Text3.Text = "") Or (Text4.Text = "") Then
MsgBox ("Please Fill All The Details"), vbCritical
End If

End Sub
