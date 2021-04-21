Attribute VB_Name = "Module2"
Option Explicit
Option Compare Text
Global RoundTripTime As String

Private plPort As String
Private psHost As String

Private pColErrors As Collection

Private psPort As String

Private psTo As String
Private psToDisplay As String
Private psFrom As String
Private psFromDisplay As String

Private psFromReply As String
Private psReplyTo As String
Private psSubject As String
Private psMessage As String
Private psAttachment As String
Private psAttachContents As String

Private pasTopLevels() As String

Private pbExitImmediately As Boolean
Private psSMTPHost As String


Private Const ERR_INVALID_HOST = "Invalid Host Name"
Private Const ERR_INVALID_PORT = "Invalid Remote Port"
Private Const ERR_INVALID_REC_EMAIL = "Missing or Invalid Recipient E-mail Address"
Private Const ERR_INVALID_SND_EMAIL = "Missing or Invalid Sender E-mail Address"
Private Const ERR_TIMEOUT = "Timeout occurred: The SMTP Host did not respond to the request"
Private Const ERR_FILE_NOT_EXIST = "The file you tried to attach does not exist"
Private Const SETTINGS_KEY = "Settings"
Private Const NUM_TOP_LEVELS = 251



Private Const CONNECT_TIMEOUT = 45
Private Const MSG_TIMEOUT = 60
Private Const MSG_ATTACH_TIMEOUT = 600

Private plConnectTimeout As Long
Private plMessageTimeOut As Long

'Added to Version 1.6
'Allows for different client-side
'validation of SMTP Host addresses
'and e-mail addresses (or no
'validation)

Public Enum VALIDATE_METHOD
    VALIDATE_NONE = 0
    VALIDATE_SYNTAX = 1
    VALIDATE_PING = 2
End Enum
    
Private etEmailValidation As VALIDATE_METHOD
Private etSMTPHostValidation As VALIDATE_METHOD


Private Sub Class_Initialize()
Dim sKey As String
Set pColErrors = New Collection
sKey = App.EXEName
ReDim pasTopLevels(NUM_TOP_LEVELS - 1) As String
InitializeTopLevels

'SMTPHost = GetSetting(sKey, SETTINGS_KEY, "RemoteHost", "")
'SMTPHostValidation = GetSetting(sKey, SETTINGS_KEY, "SMTPHostValidation", VALIDATE_SYNTAX)
'SMTPPort = CLng(GetSetting(sKey, SETTINGS_KEY, "RemotePort", "25"))


ConnectTimeout = CLng(GetSetting(sKey, SETTINGS_KEY, "ConnectTimeout", 0))
MessageTimeout = CLng(GetSetting(sKey, SETTINGS_KEY, "MessageTimeout", 0))

EmailAddressValidation = GetSetting(sKey, SETTINGS_KEY, "EmailValidation", VALIDATE_SYNTAX)
'FromDisplayName = GetSetting(sKey, SETTINGS_KEY, "FromDisplayName", "")
'From = GetSetting(sKey, SETTINGS_KEY, "From", "")
'Set Form1.EmailClient = Me




End Sub

Private Function IsValidIPHost(HostString As String) As Boolean

Dim sHost As String
Dim bDottedQuad As Boolean
Dim sSplit() As String
Dim iCtr As Integer
Dim bAns As Boolean
Dim sTopLevelDomains() As String

sHost = HostString

If InStr(sHost, ".") = 0 Then
    IsValidIPHost = False
    Exit Function
End If

sSplit = Split(sHost, ".")

If UBound(sSplit) = 3 Then
    bDottedQuad = True
    For iCtr = 0 To 3
        If Not IsNumeric(sSplit(iCtr)) Then
            bDottedQuad = False
            Exit For
        End If
    Next



    If bDottedQuad Then
        bAns = True
        For iCtr = 0 To 3
            If iCtr = 0 Then
            bAns = Val(sSplit(iCtr)) <= 239
                If bAns = False Then Exit For
            Else
                bAns = Val(sSplit(iCtr)) <= 255
                If bAns = False Then Exit For
            End If
        Next
        
        IsValidIPHost = bAns

        
        Exit Function
    End If
End If 'ubound(ssplit) = 3

    
    
    IsValidIPHost = isTopLevelDomain(sSplit(UBound(sSplit)))





End Function

Private Function isTopLevelDomain(DomainString As String) As Boolean

Dim iCtr As Integer

Dim bAns As Boolean


For iCtr = 0 To NUM_TOP_LEVELS - 1
    If pasTopLevels(iCtr) = DomainString Then
        bAns = True
        Exit For
    End If
Next

isTopLevelDomain = bAns

End Function
Sub AddError(ErrStr As String)
On Error Resume Next
pColErrors.Add ErrStr, ErrStr

End Sub
Private Sub RemoveError(ErrStr As String)

If pColErrors.Count > 0 Then
    On Error Resume Next
    pColErrors.Remove ErrStr
End If

End Sub

Private Sub InitializeTopLevels()
'Obtained from www.IANA.com.  Can and will change

pasTopLevels(0) = "COM"
pasTopLevels(1) = "ORG"
pasTopLevels(2) = "NET"
pasTopLevels(3) = "EDU"
pasTopLevels(4) = "GOV"
pasTopLevels(5) = "MIL"
pasTopLevels(6) = "INT"
pasTopLevels(7) = "AF"
pasTopLevels(8) = "AL"
pasTopLevels(9) = "DZ"
pasTopLevels(10) = "AS"
pasTopLevels(11) = "AD"
pasTopLevels(12) = "AO"
pasTopLevels(13) = "AI"
pasTopLevels(14) = "AQ"
pasTopLevels(15) = "AG"
pasTopLevels(16) = "AR"
pasTopLevels(17) = "AM"
pasTopLevels(18) = "AW"
pasTopLevels(19) = "AC"
pasTopLevels(20) = "AU"
pasTopLevels(21) = "AT"
pasTopLevels(22) = "AZ"
pasTopLevels(23) = "BS"
pasTopLevels(24) = "BH"
pasTopLevels(25) = "BD"
pasTopLevels(26) = "BB"
pasTopLevels(27) = "BY"
pasTopLevels(28) = "BZ"
pasTopLevels(29) = "BT"
pasTopLevels(30) = "BJ"
pasTopLevels(31) = "BE"
pasTopLevels(32) = "BM"
pasTopLevels(33) = "BO"
pasTopLevels(34) = "BA"
pasTopLevels(35) = "BW"
pasTopLevels(36) = "BV"
pasTopLevels(37) = "BR"
pasTopLevels(38) = "IO"
pasTopLevels(39) = "BN"
pasTopLevels(40) = "BG"
pasTopLevels(41) = "BF"
pasTopLevels(42) = "BI"
pasTopLevels(43) = "KH"
pasTopLevels(44) = "CM"
pasTopLevels(45) = "CA"
pasTopLevels(46) = "CV"
pasTopLevels(47) = "KY"
pasTopLevels(48) = "CF"
pasTopLevels(49) = "TD"
pasTopLevels(50) = "CL"
pasTopLevels(51) = "CN"
pasTopLevels(52) = "CX"
pasTopLevels(53) = "CC"
pasTopLevels(54) = "CO"
pasTopLevels(55) = "KM"
pasTopLevels(56) = "CD"
pasTopLevels(57) = "CG"
pasTopLevels(58) = "CK"
pasTopLevels(59) = "CR"
pasTopLevels(60) = "CI"
pasTopLevels(61) = "HR"
pasTopLevels(62) = "CU"
pasTopLevels(63) = "CY"
pasTopLevels(64) = "CZ"
pasTopLevels(65) = "DK"
pasTopLevels(66) = "DJ"
pasTopLevels(67) = "DM"
pasTopLevels(68) = "DO"
pasTopLevels(69) = "TP"
pasTopLevels(70) = "EC"
pasTopLevels(71) = "EG"
pasTopLevels(72) = "SV"
pasTopLevels(73) = "GQ"
pasTopLevels(74) = "ER"
pasTopLevels(75) = "EE"
pasTopLevels(76) = "ET"
pasTopLevels(77) = "FK"
pasTopLevels(78) = "FO"
pasTopLevels(79) = "FJ"
pasTopLevels(80) = "FI"
pasTopLevels(81) = "FR"
pasTopLevels(82) = "GF"
pasTopLevels(83) = "PF"
pasTopLevels(84) = "TF"
pasTopLevels(85) = "GA"
pasTopLevels(86) = "GM"
pasTopLevels(87) = "GE"
pasTopLevels(88) = "DE"
pasTopLevels(89) = "GH"
pasTopLevels(90) = "GI"
pasTopLevels(91) = "GR"
pasTopLevels(92) = "GL"
pasTopLevels(93) = "GD"
pasTopLevels(94) = "GP"
pasTopLevels(95) = "GU"
pasTopLevels(96) = "GT"
pasTopLevels(97) = "GG"
pasTopLevels(98) = "GN"
pasTopLevels(99) = "GW"
pasTopLevels(100) = "GY"
pasTopLevels(101) = "HT"
pasTopLevels(102) = "HM"
pasTopLevels(103) = "VA"
pasTopLevels(104) = "HN"
pasTopLevels(105) = "HK"
pasTopLevels(106) = "HU"
pasTopLevels(107) = "IS"
pasTopLevels(108) = "IN"
pasTopLevels(109) = "ID"
pasTopLevels(110) = "IR"
pasTopLevels(111) = "IQ"
pasTopLevels(112) = "IE"
pasTopLevels(113) = "IM"
pasTopLevels(114) = "IL"
pasTopLevels(115) = "IT"
pasTopLevels(116) = "JM"
pasTopLevels(117) = "JP"
pasTopLevels(118) = "JE"
pasTopLevels(119) = "JO"
pasTopLevels(120) = "KZ"
pasTopLevels(121) = "KE"
pasTopLevels(122) = "KI"
pasTopLevels(123) = "KP"
pasTopLevels(124) = "KR"
pasTopLevels(125) = "KW"
pasTopLevels(126) = "KG"
pasTopLevels(127) = "LA"
pasTopLevels(128) = "LV"
pasTopLevels(129) = "LB"
pasTopLevels(130) = "LS"
pasTopLevels(131) = "LR"
pasTopLevels(132) = "LY"
pasTopLevels(133) = "LI"
pasTopLevels(134) = "LT"
pasTopLevels(135) = "LU"
pasTopLevels(136) = "MO"
pasTopLevels(137) = "MK"
pasTopLevels(138) = "MG"
pasTopLevels(139) = "MW"
pasTopLevels(140) = "MY"
pasTopLevels(141) = "MV"
pasTopLevels(142) = "ML"
pasTopLevels(143) = "MT"
pasTopLevels(144) = "MH"
pasTopLevels(145) = "MQ"
pasTopLevels(146) = "MR"
pasTopLevels(147) = "MU"
pasTopLevels(148) = "YT"
pasTopLevels(149) = "MX"
pasTopLevels(150) = "FM"
pasTopLevels(151) = "MD"
pasTopLevels(152) = "MC"
pasTopLevels(153) = "MN"
pasTopLevels(154) = "MS"
pasTopLevels(155) = "MA"
pasTopLevels(156) = "MZ"
pasTopLevels(157) = "MM"
pasTopLevels(158) = "NA"
pasTopLevels(159) = "NR"
pasTopLevels(160) = "NP"
pasTopLevels(161) = "NL"
pasTopLevels(162) = "AN"
pasTopLevels(163) = "NC"
pasTopLevels(164) = "NZ"
pasTopLevels(165) = "NI"
pasTopLevels(166) = "NE"
pasTopLevels(167) = "NG"
pasTopLevels(168) = "NU"
pasTopLevels(169) = "NF"
pasTopLevels(170) = "MP"
pasTopLevels(171) = "NO"
pasTopLevels(172) = "OM"
pasTopLevels(173) = "PK"
pasTopLevels(174) = "PW"
pasTopLevels(175) = "PA"
pasTopLevels(176) = "PG"
pasTopLevels(177) = "PY"
pasTopLevels(178) = "PE"
pasTopLevels(179) = "PH"
pasTopLevels(180) = "PN"
pasTopLevels(181) = "PL"
pasTopLevels(182) = "PT"
pasTopLevels(183) = "PR"
pasTopLevels(184) = "QA"
pasTopLevels(185) = "RE"
pasTopLevels(186) = "RO"
pasTopLevels(187) = "RU"
pasTopLevels(188) = "RW"
pasTopLevels(189) = "KN"
pasTopLevels(190) = "LC"
pasTopLevels(191) = "VC"
pasTopLevels(192) = "WS"
pasTopLevels(193) = "SM"
pasTopLevels(194) = "ST"
pasTopLevels(195) = "SA"
pasTopLevels(196) = "SN"
pasTopLevels(197) = "SC"
pasTopLevels(198) = "SL"
pasTopLevels(199) = "SG"
pasTopLevels(200) = "SK"
pasTopLevels(201) = "SI"
pasTopLevels(202) = "SB"
pasTopLevels(203) = "SO"
pasTopLevels(204) = "ZA"
pasTopLevels(205) = "GS"
pasTopLevels(206) = "ES"
pasTopLevels(207) = "LK"
pasTopLevels(208) = "SH"
pasTopLevels(209) = "PM"
pasTopLevels(210) = "SD"
pasTopLevels(211) = "SR"
pasTopLevels(212) = "SJ"
pasTopLevels(213) = "SZ"
pasTopLevels(214) = "SE"
pasTopLevels(215) = "CH"
pasTopLevels(216) = "SY"
pasTopLevels(217) = "TW"
pasTopLevels(218) = "TJ"
pasTopLevels(219) = "TZ"
pasTopLevels(220) = "TH"
pasTopLevels(221) = "TG"
pasTopLevels(222) = "TK"
pasTopLevels(223) = "TO"
pasTopLevels(224) = "TT"
pasTopLevels(225) = "TN"
pasTopLevels(226) = "TR"
pasTopLevels(227) = "TM"
pasTopLevels(228) = "TC"
pasTopLevels(229) = "TV"
pasTopLevels(230) = "UG"
pasTopLevels(231) = "UA"
pasTopLevels(232) = "AE"
pasTopLevels(233) = "GB"
pasTopLevels(234) = "US"
pasTopLevels(235) = "UM"
pasTopLevels(236) = "UY"
pasTopLevels(237) = "UZ"
pasTopLevels(238) = "VU"
pasTopLevels(239) = "VE"
pasTopLevels(240) = "VN"
pasTopLevels(241) = "VG"
pasTopLevels(242) = "VI"
pasTopLevels(243) = "WF"
pasTopLevels(244) = "EH"
pasTopLevels(245) = "YE"
pasTopLevels(246) = "YU"
pasTopLevels(247) = "ZR"
pasTopLevels(248) = "ZM"
pasTopLevels(249) = "ZW"
pasTopLevels(250) = "UK"
End Sub

Public Property Get ReplyToAddress() As String
    ReplyToAddress = psReplyTo
End Property

Public Property Let ReplyToAddress(ByVal NewValue As String)
        psReplyTo = "<" & NewValue & ">"
End Property
Public Sub Send()
Dim sMsg As String
Dim sSenderName As String
Dim sRecipientName As String
Dim sAttachFileName As String
Dim sSplit() As String
Dim lMessageTimeOut As Long
Dim lConnectTimeOut As Long
Dim sAttachArray() As String
Dim iAttach As Integer

lConnectTimeOut = IIf(plConnectTimeout > 0, plConnectTimeout, CONNECT_TIMEOUT)

If plMessageTimeOut > 0 Then
    lMessageTimeOut = plMessageTimeOut
Else
    lMessageTimeOut = IIf(psAttachment = "", MSG_TIMEOUT, MSG_ATTACH_TIMEOUT)
End If


If pColErrors.Count > 0 Then
    'if there's already an error, we won't bother to try sending
    SendFail
Else

    ResetGlobals
    pbExitImmediately = False
    

        'RaiseEvent Status("Connecting to Server...")
    With Form1.sckMail
        If .State <> sckConnected Then
            .Connect
            WaitUntilTrue gbConnectOK, lConnectTimeOut
            If pbExitImmediately Then Exit Sub
            
      End If
            If Not gbConnectOK Then
                TimeOut
                Exit Sub
            End If
            'RaiseEvent Status("Initializing Contact With Server...")
            gbRequestAccepted = False
            .SendData "HELO " & Mid$(psFrom, InStr(psFrom, "@") + 1) & vbCrLf
            WaitUntilTrue gbRequestAccepted, lConnectTimeOut
            If pbExitImmediately Then Exit Sub
            
            If Not gbRequestAccepted Then
                TimeOut
                Exit Sub
            End If
            

            
            
            
            gbRequestAccepted = False
            'RaiseEvent Status("Sending Sender Information...")
            .SendData "MAIL FROM: <" & psFrom & ">" & vbCrLf
           
           WaitUntilTrue gbRequestAccepted, lConnectTimeOut
             If pbExitImmediately Then Exit Sub
            If Not gbRequestAccepted Then
                TimeOut
                Exit Sub
            End If

         gbRequestAccepted = False
        'RaiseEvent Status("Sending Recipient Information...")
         .SendData "RCPT TO: <" & psTo & ">" & vbCrLf
         
            WaitUntilTrue gbRequestAccepted, lConnectTimeOut
             If pbExitImmediately Then Exit Sub
            If Not gbRequestAccepted Then
                TimeOut
                Exit Sub
            End If
            
            gbDataOK = False
            'RaiseEvent Status("Sending Message...")
            .SendData "DATA" & vbCrLf
            WaitUntilTrue gbDataOK, lConnectTimeOut
             If pbExitImmediately Then Exit Sub
            If Not gbDataOK Then
                TimeOut
                Exit Sub
            End If
            

            gbRequestAccepted = False
            
            
            If Len(psFromDisplay) Then sSenderName = Chr$(34) & psFromDisplay & Chr(34)
            If Len(sSenderName) Then sSenderName = sSenderName & " "
            sSenderName = sSenderName & "<" & psFrom & ">"
            
             If Len(psToDisplay) Then sRecipientName = Chr$(34) & psToDisplay & Chr$(34)
            If Len(sRecipientName) Then sRecipientName = sRecipientName & " "
            sRecipientName = sRecipientName & "<" & psTo & ">"

            .SendData "MIME-Version: 1.0" & vbCrLf
            .SendData "FROM: " & sSenderName & vbCrLf
            .SendData "TO: " & sRecipientName & vbCrLf
            If Len(psReplyTo) Then .SendData "Reply-to: " & psReplyTo & vbCrLf
            
            .SendData "SUBJECT: " & psSubject & vbCrLf
            .SendData "Content-Type: multipart/mixed;" & vbCrLf
           .SendData " boundary=Unique-Boundary" & vbCrLf & vbCrLf
            .SendData Space(10) & vbCrLf & vbCrLf
            .SendData "--Unique-Boundary" & vbCrLf
            .SendData "Content-type: text/plain; charset=US-ASCII" & vbCrLf & vbCrLf
         
           .SendData psMessage & vbCrLf & vbCrLf
         
  

             If psAttachment <> "" Then
    
                sAttachArray = Split(psAttachment, ",")
                For iAttach = 0 To UBound(sAttachArray)
                    If sAttachArray(iAttach) <> "" Then
                        sSplit = Split(sAttachArray(iAttach), "\")
                        sAttachFileName = sSplit(UBound(sSplit))
                    End If
                     'RaiseEvent Status("Sending Attachment...")
                    .SendData "--Unique-Boundary" & vbCrLf
                    .SendData "Content-Type: multipart/parallel; boundary=Unique-Boundary-2" & vbCrLf & vbCrLf
                    .SendData "--Unique-Boundary-2" & vbCrLf
                    .SendData "Content-Type: application/octet-stream;" & vbCrLf
                    .SendData " name=" & sAttachFileName & vbCrLf
                    .SendData "Content-Transfer-Encoding: base64" & vbCrLf
                    .SendData "Content-Disposition: inline;" & vbCrLf
                    .SendData " filename=" & sAttachFileName & vbCrLf & vbCrLf
                    
                    EncodeAndSendFile sAttachArray(iAttach)
                    .SendData "==" & vbCrLf
                Next iAttach
            
            End If
            
            .SendData vbCrLf & "." & vbCrLf
             '.SendData sMsg
             gbRequestAccepted = False
             WaitUntilTrue gbRequestAccepted, lMessageTimeOut
             If pbExitImmediately Then Exit Sub
            If Not gbRequestAccepted Then
                TimeOut
                Exit Sub
            End If
            'If gbRequestAccepted Then RaiseEvent SendSuccesful
            .Close
         

End With
End If



End Sub
Sub SendFail()
Dim iCount As Integer, iCtr As Integer
Dim sErrorString As String
Dim v As Variant
pbExitImmediately = True

iCount = pColErrors.Count
    For iCtr = 1 To iCount
        sErrorString = sErrorString & pColErrors(iCtr)
        If iCtr < iCount Then sErrorString = sErrorString & vbCrLf
        
    Next


'RaiseEvent SendFailed(sErrorString)
With Form1.sckMail

If Not .State = sckClosed Then
    Form1.sckMail.Close
    Do Until Form1.sckMail.State = sckClosed
        DoEvents
    Loop
End If
End With
'clear errors
Set pColErrors = New Collection
End Sub


Private Sub Class_Terminate()
On Error Resume Next
If Form1.sckMail.State = sckClosed Then
    Form1.sckMail.Close
End If

Unload Form1
Set Form1 = Nothing
    
End Sub

Public Function IsValidEmailAddress(AddressString As String)

Dim sHost As String
Dim iPos As Integer

If Len(Trim(AddressString)) = 0 Then
    IsValidEmailAddress = False
    Exit Function
End If


iPos = InStr(AddressString, "@")




If iPos = 0 Or Left(AddressString, 1) = "@" Then
    IsValidEmailAddress = False
    Exit Function
End If

sHost = Mid(AddressString, iPos + 1)
'can't have multiple "@" chars in the string
If InStr(sHost, "@") > 0 Then
    IsValidEmailAddress = False
    Exit Function
End If

IsValidEmailAddress = IsValidIPHost(sHost)


End Function

Private Sub TimeOut()
    AddError ERR_TIMEOUT
    SendFail
End Sub
Private Sub WaitUntilTrue(Flag As Boolean, TimeToWait As Long)
'PURPOSE:  Wait until either
'a condition is true or a timeout occurs

'The condition is specified by the value of
'flag passed by reference

'The TimeOut value is set by the
'timetowait parameter


Dim lStart As Long
Dim bPastMidnight As Boolean
Dim lTimetoQuit As Long


    lStart = CLng(Timer)
    'Deal with timeout being reset at Midnight
    If TimeToWait > 0 Then
        If lStart + TimeToWait < 86400 Then
            lTimetoQuit = lStart + TimeToWait
        Else
            lTimetoQuit = (lStart - 86400) + TimeToWait
            bPastMidnight = True
        End If
    End If

    Do Until Flag = True Or Timer >= lTimetoQuit
        DoEvents
        If pbExitImmediately Then Exit Sub
    Loop
   
End Sub
Private Function EncodeString(ByVal StringValue As String) As String

Dim iCount As Integer
Dim sBinary As String
Dim iDecimal As Integer
Dim sTemp As String
Dim sValue As String

sValue = StringValue
If Len(sValue) = 0 Then Exit Function
iDecimal = Asc(Left$(sValue, 1))

For iCount = 7 To 0 Step -1
If (2 ^ iCount) <= iDecimal Then
sBinary = sBinary & "1"
iDecimal = iDecimal - (2 ^ iCount)
Else
sBinary = sBinary & "0"
End If
Next

If Len(sValue) < 3 Then GoTo unfpassone

iDecimal = Asc(Mid$(sValue, 2, 1))

For iCount = 7 To 0 Step -1
If (2 ^ iCount) <= iDecimal Then
sBinary = sBinary & "1"
iDecimal = iDecimal - (2 ^ iCount)
Else
sBinary = sBinary & "0"
End If
Next

If Len(sValue) < 3 Then GoTo unfpassone

iDecimal = Asc(Right$(sValue, 1))

For iCount = 7 To 0 Step -1
If (2 ^ iCount) <= iDecimal Then
sBinary = sBinary & "1"
iDecimal = iDecimal - (2 ^ iCount)
Else
sBinary = sBinary & "0"
End If
Next

unfpassone:
For iCount = 1 To 19 Step 6
Select Case Val(Mid$(sBinary, iCount, 6))
Case 0
sTemp = sTemp & "A"
Case 1
sTemp = sTemp & "B"
Case 10
sTemp = sTemp & "C"
Case 11
sTemp = sTemp & "D"
Case 100
sTemp = sTemp & "E"
Case 101
sTemp = sTemp & "F"
Case 110
sTemp = sTemp & "G"
Case 111
sTemp = sTemp & "H"
Case 1000
sTemp = sTemp & "I"
Case 1001
sTemp = sTemp & "J"
Case 1010
sTemp = sTemp & "K"
Case 1011
sTemp = sTemp & "L"
Case 1100
sTemp = sTemp & "M"
Case 1101
sTemp = sTemp & "N"
Case 1110
sTemp = sTemp & "O"
Case 1111
sTemp = sTemp & "P"
Case 10000
sTemp = sTemp & "Q"
Case 10001
sTemp = sTemp & "R"
Case 10010
sTemp = sTemp & "S"
Case 10011
sTemp = sTemp & "T"
Case 10100
sTemp = sTemp & "U"
Case 10101
sTemp = sTemp & "V"
Case 10110
sTemp = sTemp & "W"
Case 10111
sTemp = sTemp & "X"
Case 11000
sTemp = sTemp & "Y"
Case 11001
sTemp = sTemp & "Z"
Case 11010
sTemp = sTemp & "a"
Case 11011
sTemp = sTemp & "b"
Case 11100
sTemp = sTemp & "c"
Case 11101
sTemp = sTemp & "d"
Case 11110
sTemp = sTemp & "e"
Case 11111
sTemp = sTemp & "f"
Case 100000
sTemp = sTemp & "g"
Case 100001
sTemp = sTemp & "h"
Case 100010
sTemp = sTemp & "i"
Case 100011
sTemp = sTemp & "j"
Case 100100
sTemp = sTemp & "k"
Case 100101
sTemp = sTemp & "l"
Case 100110
sTemp = sTemp & "m"
Case 100111
sTemp = sTemp & "n"
Case 101000
sTemp = sTemp & "o"
Case 101001
sTemp = sTemp & "p"
Case 101010
sTemp = sTemp & "q"
Case 101011
sTemp = sTemp & "r"
Case 101100
sTemp = sTemp & "s"
Case 101101
sTemp = sTemp & "t"
Case 101110
sTemp = sTemp & "u"
Case 101111
sTemp = sTemp & "v"
Case 110000
sTemp = sTemp & "w"
Case 110001
sTemp = sTemp & "x"
Case 110010
sTemp = sTemp & "y"
Case 110011
sTemp = sTemp & "z"
Case 110100
sTemp = sTemp & "0"
Case 110101
sTemp = sTemp & "1"
Case 110110
sTemp = sTemp & "2"
Case 110111
sTemp = sTemp & "3"
Case 111000
sTemp = sTemp & "4"
Case 111001
sTemp = sTemp & "5"
Case 111010
sTemp = sTemp & "6"
Case 111011
sTemp = sTemp & "7"
Case 111100
sTemp = sTemp & "8"
Case 111101
sTemp = sTemp & "9"
Case 111110
sTemp = sTemp & "+"
Case 111111
sTemp = sTemp & "/"
End Select
Next

EncodeString = sTemp

End Function


Private Sub EncodeAndSendFile(strFile As String)

Dim lCtr As Long

Dim sTemp As String
Dim sInput As String, sOutput As String

Dim lngMax As Long
Dim iFileNum As Integer
Dim sAns As String
Dim lLen As Long


lngMax = 0

iFileNum = FreeFile
Open strFile For Binary Access Read As #iFileNum

Do

    Form1.sckMail.SendData EncodeString(Input(3, #iFileNum))
    lngMax = lngMax + 4

    If lngMax = 72 Then
       lngMax = 0

        Form1.sckMail.SendData vbCrLf
    End If

DoEvents
If EOF(iFileNum) Then Exit Do
Loop

Close #iFileNum


End Sub
Public Property Let Attachment(ByVal NewValue As String)
Dim sAttachName() As String
Dim bNamesOK As Boolean
Dim iAttach As Integer

If Trim(NewValue) <> "" Then
    sAttachName = Split(Trim(NewValue), ",")
    bNamesOK = True
    
    For iAttach = 0 To UBound(sAttachName)
        If Dir(sAttachName(iAttach)) = "" Then bNamesOK = False
    Next iAttach
    
    If bNamesOK = True Then
        psAttachment = NewValue
        RemoveError ERR_FILE_NOT_EXIST
    Else
        AddError ERR_FILE_NOT_EXIST
    End If

Else
    psAttachment = NewValue
    RemoveError ERR_FILE_NOT_EXIST
End If

End Property
Public Property Get Attachment() As String

Attachment = psAttachment
End Property

Private Function IsDottedQuad(HostString As String) As Boolean

Dim sHost As String
Dim sSplit() As String
Dim iCtr As Integer
Dim bAns As Boolean

sHost = HostString

If InStr(sHost, ".") = 0 Then
    IsDottedQuad = False
    Exit Function
End If

sSplit = Split(sHost, ".")

If UBound(sSplit) = 3 Then
    For iCtr = 0 To 3
        If Not IsNumeric(sSplit(iCtr)) Then
            IsDottedQuad = False
            Exit Function
        End If
    Next

        bAns = True
        For iCtr = 0 To 3
            If iCtr = 0 Then
            bAns = Val(sSplit(iCtr)) <= 239
                If bAns = False Then Exit For
            Else
                bAns = Val(sSplit(iCtr)) <= 255
                If bAns = False Then Exit For
            End If
        Next
        
End If 'ubound(ssplit) = 3

IsDottedQuad = bAns


End Function
Function Ping(Address As String) As Boolean

Dim ECHO As ICMP_ECHO_REPLY
Dim pos As Integer
Dim Dt As String
Dim sAddress As String

'THIS CODE IS BASED ON FUNCTIONS
'WITHIN RICHARD DEEMING'S IP UTILITIES:
'http://www.freevbcode.com/showcode.asp?ID=199

On Error GoTo DPErr
    If Not IsDottedQuad(Address) Then
        sAddress = GetIPAddress(Address)
    Else
        sAddress = Address
    End If
    
    If sAddress = "" Then Exit Function
    
    If SocketsInitialize Then
       
        For pos = 1 To DATA_SIZE
            Dt = Dt & Chr$(Rnd() * 254 + 1)
        Next pos
        
        'ping an ip address, passing the
        'address and the ECHO structure
        Ping = DoPing(sAddress, Dt, ECHO)
        
        'display the results from the ECHO structure
        RoundTripTime = ECHO.RoundTripTime & " ms"
        
        'DataSize = ECHO.DataSize & " bytes"
      
        'If Left$(ECHO.Data, 1) <> Chr$(0) Then
        '    pos = InStr(ECHO.Data, Chr$(0))
         '   DataMatch = (Left$(ECHO.Data, pos - 1) = Dt)
        'End If
   
        SocketsCleanup
    Else
        Ping = IP_GENERAL_FAILURE
    End If
    Exit Function
DPErr:
    Ping = IP_GENERAL_FAILURE
End Function

Private Function DoPing(szAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Boolean

Dim hPort As Long, dwAddress As Long, iOpt As Long
Dim lAns As Long

    dwAddress = AddressStringToLong(szAddress)
   
    hPort = IcmpCreateFile()
    
    If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, _
        ECHO, Len(ECHO), PING_TIMEOUT) Then
        'the ping succeeded,
        '.Status will be 0
        '.RoundTripTime is the time in ms for
        '               the ping to complete,
        '.Data is the data returned (NULL terminated)
        '.Address is the Ip address that actually replied
        '.DataSize is the size of the string in .Data
        lAns = IP_SUCCESS
    Else
        If ECHO.Status = 0 Then
            lAns = -1
        Else
            lAns = ECHO.Status * -1
        End If
    End If
                       
    Call IcmpCloseHandle(hPort)
    DoPing = lAns = IP_SUCCESS
End Function
Function GetIPAddress(sHost As String) As String

'Resolves the host-name (or current machine if balnk) to an IP address
Dim sHostName   As String * 256
Dim lpHost      As Long
Dim HOST        As HOSTENT
Dim dwIPAddr    As Long
Dim tmpIPAddr() As Byte
Dim i           As Integer
Dim sIPAddr     As String
Dim werr        As Long

    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
 
    sHostName = Trim$(sHost) & Chr$(0)
    
    lpHost = gethostbyname(sHostName)

    If lpHost = 0 Then
        werr = WSAGetLastError()
        GetIPAddress = ""
                
        SocketsCleanup
        Exit Function
    End If

    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4

    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen

    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next

    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)

    SocketsCleanup
End Function
Private Function AddressStringToLong(ByVal tmp As String) As Long
Dim i As Integer, parts(1 To 4) As String
    i = 0
    'we have to extract each part of the
    '123.456.789.123 string, delimited by
    'a period
    While InStr(tmp, ".") > 0
        i = i + 1
        parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
        tmp = Mid(tmp, InStr(tmp, ".") + 1)
    Wend
    
    i = i + 1
    parts(i) = tmp
    
    If i <> 4 Then
        AddressStringToLong = 0
        Exit Function
    End If
   
    'build the long value out of the
    'hex of the extracted strings
    AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                         Right("00" & Hex(parts(3)), 2) & _
                         Right("00" & Hex(parts(2)), 2) & _
                         Right("00" & Hex(parts(1)), 2))
   
End Function

Public Property Get EmailAddressValidation() As VALIDATE_METHOD
EmailAddressValidation = etEmailValidation

End Property

Public Property Let EmailAddressValidation(ByVal NewValue As VALIDATE_METHOD)

If NewValue >= VALIDATE_NONE And NewValue <= VALIDATE_PING Then
    etEmailValidation = NewValue
    SaveSetting App.EXEName, SETTINGS_KEY, "EmailValidation", NewValue
End If

End Property

Public Property Get SMTPHostValidation() As VALIDATE_METHOD
SMTPHostValidation = etSMTPHostValidation

End Property

Public Property Let SMTPHostValidation(ByVal NewValue As VALIDATE_METHOD)

If NewValue >= VALIDATE_NONE And NewValue <= VALIDATE_PING Then
    etSMTPHostValidation = NewValue
    SaveSetting App.EXEName, SETTINGS_KEY, "SMTPHostValidation", NewValue
End If

'in case this is set after the host value is set
'If psSMTPHost <> "" Then SMTPHost = psSMTPHost
End Property

Public Property Get ConnectTimeout() As Long
ConnectTimeout = plConnectTimeout
End Property

Public Property Let ConnectTimeout(ByVal NewValue As Long)
plConnectTimeout = NewValue
SaveSetting App.EXEName, SETTINGS_KEY, "ConnectTimeout", NewValue
End Property

Public Property Get MessageTimeout() As Long
MessageTimeout = plMessageTimeOut
End Property

Public Property Let MessageTimeout(ByVal NewValue As Long)
plMessageTimeOut = NewValue
SaveSetting App.EXEName, SETTINGS_KEY, "MessageTimeout", NewValue
End Property


