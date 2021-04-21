Attribute VB_Name = "Module1"
Option Explicit

Public Enum IP_STATUS
    IP_STATUS_BASE = 11000
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = (11000 + 1)
    IP_DEST_NET_UNREACHABLE = (11000 + 2)
    IP_DEST_HOST_UNREACHABLE = (11000 + 3)
    IP_DEST_PROT_UNREACHABLE = (11000 + 4)
    IP_DEST_PORT_UNREACHABLE = (11000 + 5)
    IP_NO_RESOURCES = (11000 + 6)
    IP_BAD_OPTION = (11000 + 7)
    IP_HW_ERROR = (11000 + 8)
    P_PACKET_TOO_BIG = (11000 + 9)
    IP_REQ_TIMED_OUT = (11000 + 10)
    IP_BAD_REQ = (11000 + 11)
    IP_BAD_ROUTE = (11000 + 12)
    IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
    IP_TTL_EXPIRED_REASSEM = (11000 + 14)
    IP_PARAM_PROBLEM = (11000 + 15)
    IP_SOURCE_QUENCH = (11000 + 16)
    IP_OPTION_TOO_BIG = (11000 + 17)
    IP_BAD_DESTINATION = (11000 + 18)
    IP_ADDR_DELETED = (11000 + 19)
    IP_SPEC_MTU_CHANGE = (11000 + 20)
    IP_MTU_CHANGE = (11000 + 21)
    IP_UNLOAD = (11000 + 22)
    IP_ADDR_ADDED = (11000 + 23)
    IP_GENERAL_FAILURE = (11000 + 50)
    MAX_IP_STATUS = 11000 + 50
    IP_PENDING = (11000 + 255)
    PING_TIMEOUT = 255
End Enum

Public Const DATA_SIZE = 32
Public gbConnectOK As Boolean
Public gbRequestAccepted As Boolean
Public gbDataOK As Boolean
Public gbError As Boolean
Public poForm As Object

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = _
    WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = _
   WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type


Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
  '  Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type



Public Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Public Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal TimeOut As Long) As Long

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
        (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) _
        As Long
 
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
        (ByVal szHost As String, ByVal dwHostLen As Long) _
         As Long

Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
        (ByVal szHost As String) As Long

Public Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal szHost As String, ByVal hLen As Integer, _
         ByVal aType As Integer) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, ByVal _
   hpvSource As Long, ByVal cbCopy As Long)
Public Function TrimWithoutPrejudice _
(ByVal InputString As String) As String

'Trims all non-printing characters from a string
'Snippet taken from http://www.freevbcode.com

Dim sAns As String
Dim sWkg As String
Dim sChar As String
Dim lLen As Long
Dim lCtr As Long

sAns = InputString
lLen = Len(InputString)

If lLen > 0 Then
'Ltrim
    For lCtr = 1 To lLen
        sChar = Mid(sAns, lCtr, 1)
        If Asc(sChar) > 32 Then Exit For
    Next

sAns = Mid(sAns, lCtr)
lLen = Len(sAns)

'Rtrim
    If lLen > 0 Then
        For lCtr = lLen To 1 Step -1
            sChar = Mid(sAns, lCtr, 1)
            If Asc(sChar) > 32 Then Exit For
        Next
    End If
    sAns = Left$(sAns, lCtr)
End If
TrimWithoutPrejudice = sAns

End Function

Public Sub ResetGlobals()
    gbConnectOK = False
    gbRequestAccepted = False
    gbDataOK = False
    gbError = False

End Sub
Public Sub SocketsCleanup()
    WSACleanup
End Sub

Public Function SocketsInitialize() As Boolean


Dim WSAD As WSADATA, sLoByte As String, sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
    
        SocketsInitialize = False
        Exit Function
    End If

    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
      

        SocketsInitialize = False
        Exit Function
    End If
SocketsInitialize = True
End Function


