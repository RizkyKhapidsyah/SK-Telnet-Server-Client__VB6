VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Telnet Client Emu"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6870
   Icon            =   "Telnet Client.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2880
      Top             =   2520
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   4935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4320
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu Connection 
      Caption         =   "Connection"
      Begin VB.Menu ExtCon 
         Caption         =   "External Connection"
      End
      Begin VB.Menu CloseCon 
         Caption         =   "Close Connection"
      End
      Begin VB.Menu ClCons 
         Caption         =   "Clear Console"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu LastConnection 
         Caption         =   ""
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu IndexHelp 
         Caption         =   "Index"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
MsgBox "LenSoft Inc. Telnet client emu" & Chr(13) & "Copyright© 2000, LenSoft Inc." & Chr(13) & "All rights reserved.", vbInformation
End Sub

Private Sub ClCons_Click()
Text1.Text = ""
If Winsock1.State = 7 Then Exit Sub
Text2.Text = ""
End Sub

Private Sub CloseCon_Click()
Winsock1.Close
End Sub

Private Sub ExtCon_Click()
Form2.Show 1
End Sub

Private Sub Form_Load()
a = GetSetting("Telnetclient", "Last", "Server", "")
If a = "" Then
line3.Visible = False
LastConnection.Visible = False
Else
LastConnection.Caption = "1. " & a
End If
End Sub

Private Sub Form_Resize()
Text1.Width = Form1.Width - 375
Text2.Width = Form1.Width - 375
Text1.Height = Form1.Height - 1275
Text2.Top = Form1.Height - 1050
End Sub

Private Sub IndexHelp_Click()
MsgBox "Index help is not available at the moment", vbInformation
End Sub

Private Sub LastConnection_Click()
On Error GoTo 1
a = GetSetting("Telnetclient", "Last", "Server", "")
b = GetSetting("Telnetclient", "Last", "Port", 23)
Winsock1.Close
Winsock1.Connect a, b
Timer1.Enabled = True
Exit Sub
1 Winsock1.Close
MsgBox Err.Description, vbCritical
End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
Text2.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo 1
If KeyAscii = 13 Then
Text2.Text = ""
Winsock1.SendData Chr(13)
KeyAscii = 0
Else
Winsock1.SendData Chr(KeyAscii)
End If
Exit Sub
1 KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Winsock1.Close
MsgBox "Unable to connect", vbCritical
End Sub

Private Sub Winsock1_Close()
Timer1.Enabled = False
MsgBox "Connection with host lost"
Me.Caption = "Telnet Client Emu"
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
Me.Caption = "Telnet Client Emu [" & Winsock1.RemoteHostIP & ":" & Winsock1.RemotePort & "]"
Timer1.Enabled = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData buffer$, vbString
Text1.Text = Text1.Text & buffer$
End Sub
