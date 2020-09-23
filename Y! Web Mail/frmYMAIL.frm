VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmYMAIL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo! Web Mail Sender"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSubject 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Text            =   "Subject"
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtTo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Login and send email"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox rtbBuffer 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmYMAIL.frx":0000
   End
   Begin VB.TextBox txtBody 
      Height          =   1245
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmYMAIL.frx":0095
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "password"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Yahoo ID"
      Top             =   360
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbSubject 
      Alignment       =   2  'Center
      Caption         =   "Subject:"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbTo 
      Alignment       =   2  'Center
      Caption         =   "Send To:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lbStats 
      Caption         =   "Status:"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lbPass 
      Alignment       =   2  'Center
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lbID 
      Alignment       =   2  'Center
      Caption         =   "Your Yahoo ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmYMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
' THIS IS A EXAMPLE OF SENDING OUT A EMAIL USING
' 'YAHOO! WEB MAIL' USING THE WINSOCK CONTROL
' TEXT BASED EMAILS ARE ONLY SUPPORTED BUT YOU CAN CHANGE THAT
' BY MINIPULATING THE EMAIL PACKET AND SETTING SOME VALUES
'////////////////////////////////////////////////////////////////
' SOURCE BY: LOUIS VASQUEZ AKA:KILL4
' SITE: HTTP://KILL4.KI.FUNPIC.ORG/
'////////////////////////////////////////////////////////////////
Dim Step As Byte 'varible that holds a byte value used to tell what http packet is to be sent

Dim TheCookies() As String 'a array that holds cookies that the server gives you

Dim Maillocation, Mailhost As String 'varibles that holds your mail server and the url to the main mail page

Dim Mailcookie, Composelocation As String 'Mailcookie varible holds a cookie that the you get when you visit your mail account
'Composelocation varible holds the url to the compose page were we will send the email from

Dim Postcompose, Thecrumb As String 'varibles that hold a composeurl and a key needed to complete the email packet its called thecrumb

'this is the main email string witch holds the data being sent such as the subject,body and to whom your sending the email to if you look at the string closely you wil see other values you can set such as 'OriginalSubject=&InReplyTo=' and others to no how to set these values
' with the right text syntax i would recomend sniffing out the packets while sending out emails with different options set
Const MailString = "SEND=1&SD=&SC=&CAN=&docCharset=windows-1252&PhotoMailUser=&PhotoToolInstall=&OpenInsertPhoto=&PhotoGetStart=0&SaveCopy=yes&.crumb=&FwdFile=&FwdMsg=&FwdSubj=&FwdInline=&OriginalFrom=&OriginalSubject=&InReplyTo=&NumAtt=0&AttData=&UplData=&OldAttData=&OldUplData=&FName=&ATT=&VID=&Markers=&NextMarker=0&Thumbnails=&BrowseState=&PhotoIcon=&ToolbarState=&VirusReport=&Attachments=&Background=&BGRef=&BGDesc=&BGDef=&BGFg=&BGFF=&BGFS=&BGSolid=&BGCustom=&PlainMsg=%3CDIV%3E%3C%2FDIV%3E&PhotoFrame=&PhotoPrintAtHomeLink=&PhotoSlideShowLink=&PhotoPrintLink=&PhotoSaveLink=&PhotoPermCap=&PhotoPermPath=&PhotoDownloadUrl=&PhotoSaveUrl=&PhotoFlags=&start=compose&showcc=&showbcc=&AC_Done=&AC_ToList=0%2C&AC_CcList=&AC_BccList=&sendtop=Send&savedrafttop=Save+as+a+Draft&canceltop=Cancel&To=&Cc=&Bcc=&Subj=&Body=%3CDIV%3E%0D%0A%3CDIV%3ENumber1%3C%2FDIV%3E%3C%2FDIV%3E&Format=html&sendbottom=Send&savedraftbottom=Save+as+a+Draft&cancelbottom=Cancel&cancelbottom=Cancel"

Private Sub cmdSend_Click()
 Step = 0
 sckMail.Close
 sckMail.Connect "mail.yahoo.com", 80
 lbStats.Caption = "Status: connecting.."
End Sub

Private Sub sckMail_Connect()
 '//////////////////////////////////////////////////
 '/ THESE ARE THE HTTP PACKETS BEING SENT OUT
 '/ I SNIFFED OUT THESE PACKETS USING A PACKET
 '/ SNIFFER WHILE GOING THROUGH THE PROCCESS OF
 '/ SENDING OUT A EMAIL
 
 '/ 1.FIRST PACKET IS THE LOGIN PACKET
 '/ 2.SECOND PACKET IS THE REQUEST MAIL PACKET
 '/ 3.THIRD PACKET IS THE MAIN MAIL PAGE PACKET
 '/ 4.FOURTH PACKET IS THE COMPOSE PAGE PACKET
 '/ 5.FIFTH PACKET IS THE PACKET THAT SEND THE EMAIL
 '/////////////////////////////////////////////////////
 Dim Packet As String
 rtbBuffer.Text = ""
 
 Select Case Step
  Case 0
   Packet = "GET /config?&login=" & txtID & "&passwd=" & txtPassword & "&.done=http://mail.yahoo.com HTTP/1.1" & vbCrLf
   Packet = Packet & "Accept: */*" & vbCrLf
   Packet = Packet & "Accept-Language: en-us" & vbCrLf
   Packet = Packet & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows 98; Creative)" & vbCrLf
   Packet = Packet & "Host: login.yahoo.com" & vbCrLf
   Packet = Packet & "Connection: Keep-Alive" & vbCrLf
   Packet = Packet & "Accept: text/html" & vbCrLf & vbCrLf
   
   Packet = Packet
   sckMail.SendData Packet
   lbStats.Caption = "Status: sent login packet.."
   Debug.Print Packet
  
  Case 1
   Packet = "GET / HTTP/1.1" & vbCrLf
   Packet = Packet & "Accept: */*" & vbCrLf
   Packet = Packet & "Accept-Language: en-us" & vbCrLf
   Packet = Packet & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows 98; Creative)" & vbCrLf
    If UBound(TheCookies) = 4 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & vbCrLf
    ElseIf UBound(TheCookies) >= 5 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & "; " & TheCookies(5) & vbCrLf
    End If
   Packet = Packet & "Connection: Keep-Alive" & vbCrLf
   Packet = Packet & "Host: mail.yahoo.com" & vbCrLf
   Packet = Packet & "Accept: text/html" & vbCrLf & vbCrLf
   
   Packet = Packet
   sckMail.SendData Packet
   lbStats.Caption = "Status: sent mail request packet.."
   Debug.Print Packet
   
  Case 2
   Packet = "GET " & Maillocation & " HTTP/1.1" & vbCrLf
   Packet = Packet & "Accept: */*" & vbCrLf
   Packet = Packet & "Accept-Language: en-us" & vbCrLf
   Packet = Packet & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows 98; Creative)" & vbCrLf
   Packet = Packet & "Connection: Keep-Alive" & vbCrLf
   Packet = Packet & "Host: " & Mailhost & vbCrLf
    If UBound(TheCookies) = 4 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & vbCrLf
    ElseIf UBound(TheCookies) >= 5 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & "; " & TheCookies(5) & vbCrLf
    End If
   Packet = Packet & "Accept: text/html" & vbCrLf & vbCrLf

   Packet = Packet
   sckMail.SendData Packet
   lbStats.Caption = "Status: sent request mail packet.."
   Debug.Print Packet
   
  Case 3
   Packet = "GET " & Composelocation & " HTTP/1.1" & vbCrLf
   Packet = Packet & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, */*" & vbCrLf
   Packet = Packet & "Accept-Language: en-us" & vbCrLf
   Packet = Packet & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows 98; Creative)" & vbCrLf
   Packet = Packet & "Host: " & Mailhost & vbCrLf
   Packet = Packet & "Connection: Keep-Alive" & vbCrLf
   If UBound(TheCookies) = 4 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & "; CP=v=50507&br=i&sp=; U=mt=ALtHyZ2MhYo5dNpc0X44pimdVsZzeAby5osb&ux=QG2zCB&un=abp3naemp6dpa; " & Mailcookie & vbCrLf
   ElseIf UBound(TheCookies) >= 5 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & "; " & TheCookies(5) & "; CP=v=50507&br=i&sp=; U=mt=AJLejp2MhYpqXr0hHRUBwg2Z7Uvv9E_leoevYw--&ux=Yo0zCB&un=5jom4vcur0di8; " & Mailcookie & vbCrLf
   End If
   Packet = Packet & "Accept: text/html" & vbCrLf & vbCrLf
   
   Packet = Packet
   sckMail.SendData Packet
   lbStats.Caption = "Status: sent compose request packet.."
   Debug.Print Packet
  
  Case 4
   Dim Postlenth
   Dim Body As String, Subject As String
   Body = txtBody
   Subject = txtSubject
   ReplaceSpaces Body
   ReplaceSpaces Subject
   
   Postlenth = Len(MailString) + Len(Thecrumb) + Len(txtTo) + Len(Body) + Len(Subject)
    
   Packet = "POST " & Postcompose & " HTTP/1.1" & vbCrLf
   Packet = Packet & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, */*" & vbCrLf
   Packet = Packet & "Accept: text/html" & vbCrLf
   Packet = Packet & "Referer: " & Composelocation & vbCrLf
   Packet = Packet & "Accept-Language: en-us" & vbCrLf
   Packet = Packet & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
   Packet = Packet & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows 98; Creative)" & vbCrLf
   Packet = Packet & "Host: " & Mailhost & vbCrLf
   Packet = Packet & "Content-Length: " & Postlenth & vbCrLf
   Packet = Packet & "Connection: Keep-Alive" & vbCrLf
   Packet = Packet & "Cache-Control: no-cache" & vbCrLf
   If UBound(TheCookies) = 4 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & "; CP=v=50507&br=i&sp=; U=mt=ALtHyZ2MhYo5dNpc0X44pimdVsZzeAby5osb&ux=QG2zCB&un=abp3naemp6dpa; " & Mailcookie & vbCrLf & vbCrLf
   ElseIf UBound(TheCookies) >= 5 Then
      Packet = Packet & "Cookie: " & TheCookies(1) & "; " & TheCookies(2) & "; " & TheCookies(3) & "; " & TheCookies(4) & "; " & TheCookies(5) & "; CP=v=50507&br=i&sp=; U=mt=AJLejp2MhYpqXr0hHRUBwg2Z7Uvv9E_leoevYw--&ux=Yo0zCB&un=5jom4vcur0di8; " & Mailcookie & vbCrLf & vbCrLf
   End If
   Packet = Packet & "SEND=1&SD=&SC=&CAN=&docCharset=windows-1252&PhotoMailUser=&PhotoToolInstall=&OpenInsertPhoto=&PhotoGetStart=0&SaveCopy=yes&.crumb=" & Thecrumb & "&box=&FwdFile=&FwdMsg=&FwdSubj=&FwdInline=&OriginalFrom=&OriginalSubject=&InReplyTo=&NumAtt=0&AttData=&UplData=&OldAttData=&OldUplData=&FName=&ATT=&VID=&Markers=&NextMarker=0&Thumbnails=&BrowseState=&PhotoIcon=&ToolbarState=&VirusReport=&Attachments=&Background=&BGRef=&BGDesc=&BGDef=&BGFg=&BGFF=&BGFS=&BGSolid=&BGCustom=&PlainMsg=%3CDIV%3ENumber1%3C%2FDIV%3E&PhotoFrame=&PhotoPrintAtHomeLink=&PhotoSlideShowLink=&PhotoPrintLink=&PhotoSaveLink=&PhotoPermCap=&PhotoPermPath=&PhotoDownloadUrl=&PhotoSaveUrl=&PhotoFlags=&start=compose&showcc=&showbcc=&AC_Done=&AC_ToList=0%2C&AC_CcList=&AC_BccList=&sendtop=Send&savedrafttop=Save+as+a+Draft&canceltop=Cancel&" & _
   "To=" & txtTo & "&Cc=&Bcc=&Subj=" & Subject & "&Body=%3CDIV%3E%0D%0A%3CDIV%3E" & Body & "%3C%2FDIV%3E%3C%2FDIV%3E&Format=html&sendbottom=Send&savedraftbottom=Save+as+a+Draft&cancelbottom=Cancel&cancelbottom=Cancel"
   
   Packet = Packet
   sckMail.SendData Packet
   lbStats.Caption = "Status: email packet sent!.."
   Debug.Print Packet
   
 End Select
End Sub

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)
 Dim Buff As String
 
 sckMail.GetData Buff
 rtbBuffer.Text = rtbBuffer.Text & Buff
 
 If Step = 0 And InStr(1, Buff, "y=v", vbTextCompare) Then 'we logged in and so we call the parsecookies() function
   ParseCookies Buff
   Step = Step + 1
   sckMail.Close
   sckMail.Connect "mail.yahoo.com", 80
 ElseIf Step = 0 And InStr(1, Buff, "200 ok", vbTextCompare) Then 'invalid password
   sckMail.Close
   MsgBox "invalid password"
 ElseIf Step = 1 And InStr(1, Buff, "302 found", vbTextCompare) Then 'the server gave us the main mail page location so we call the getmaillocations() function to parse out the data
  GetMailLocations Buff
  Step = Step + 1
  sckMail.Close
  sckMail.Connect Mailhost, 80
 ElseIf Step = 2 And InStr(1, rtbBuffer.Text, "ym.gen", vbTextCompare) And InStr(1, rtbBuffer.Text, "if( parent.showCompose )", vbTextCompare) Then 'were at the main mail page were we recived the mailcookie and compose page location so we call the function that handles the parsing of the data witch is getcomposelocations()
  GetComposeLocations rtbBuffer.Text
  Step = Step + 1
  sckMail.Close
  sckMail.Connect Mailhost, 80
 ElseIf Step = 3 And InStr(1, rtbBuffer.Text, "if( parent.showcompose )", vbTextCompare) And InStr(1, rtbBuffer.Text, "crumb", vbTextCompare) Then 'were at the compose page were we send out or email from and the we now have our two keys needed to finish sending the email witch is thecrumb and the postcompose locations so we call the needed function for parsing our data sendtools()
  SendTools rtbBuffer.Text
  Step = Step + 1
  sckMail.Close
  sckMail.Connect Mailhost, 80
 End If
 
 
End Sub

Function ParseCookies(Data As String)
 On Error Resume Next
 Dim Index As Integer
 'PARSE FUNCTION THAT GRABS OUR COOKIES AND PLACES THEM IN A ARRAY
 Data = Split(Data, "302 Found")(0 + 1)
 
 TheCookies = Split(Data, "Set-Cookie: ")
 For Index = 1 To UBound(TheCookies)
   TheCookies(Index) = Split(TheCookies(Index), ";")(0)
   TheCookies(Index) = Trim(TheCookies(Index))
 Next
 
End Function

Function GetComposeLocations(Data As String)
 On Error Resume Next
 'PARSES OUT OUR MAIL COOKIE AND COMPOSE URL
 Mailcookie = Split(Data, "Set-Cookie:")(1)
 Mailcookie = Split(Mailcookie, ";")(0)
 Mailcookie = Trim(Mailcookie)
 Composelocation = Split(Data, "if( parent.showCompose )")(1)
 Composelocation = Split(Composelocation, "location='")(1)
 Composelocation = Split(Composelocation, "'")(0)
End Function

Function SendTools(Data As String)
 On Error Resume Next
 'FINAL PARSE FUNCTION
 Postcompose = Split(Data, "if( parent.showCompose )")(1)
 Postcompose = Split(Postcompose, "{location='")(1)
 Postcompose = Split(Postcompose, "'")(0)
 Thecrumb = Split(Data, "crumb" & Chr(34) & " value=" & Chr(34))(1)
 Thecrumb = Split(Thecrumb, Chr(34))(0)
 Thecrumb = Trim(Thecrumb)
End Function

Function GetMailLocations(Data As String)
 On Error Resume Next
 'PARSE FUNCTION THAT PARSES OUT OR MAILURL AND WE ALSO GRAB THE MAILHOST
 Maillocation = Split(Data, "Location:")(1)
 Maillocation = Split(Maillocation, Chr(13))(0)
 Maillocation = Trim(Maillocation)
 Mailhost = Split(Maillocation, "http://")(1)
 Mailhost = Split(Mailhost, "/")(0)
End Function

Function ReplaceSpaces(Thestring As String)
 'REPLACES SPACES WITH '+' THIS IS NEEDED TO REPLACE SPACES IN OUR BODY TEXT AND
 'SUBJECT TEXT
 If InStr(1, Thestring, Space(1), vbTextCompare) Then
  Thestring = Replace(Thestring, Space(1), "+", , , vbTextCompare)
 End If
End Function

