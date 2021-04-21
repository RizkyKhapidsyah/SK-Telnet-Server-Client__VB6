VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "External Connection"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "Ext Conn.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "23"
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Hostname:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo 1
Form1.Winsock1.Close
Form1.Winsock1.Connect Text1.Text, Text2.Text
Form1.Timer1.Enabled = True
SaveSetting "Telnetclient", "Last", "Server", Text1.Text
SaveSetting "Telnetclient", "Last", "Port", Text2.Text
Unload Me
Exit Sub
1 MsgBox Err.Description, vbCritical
Form1.Winsock1.Close
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting("Telnetclient", "Last", "Server", "")
Text2.Text = GetSetting("Telnetclient", "Last", "Port", 23)
End Sub
