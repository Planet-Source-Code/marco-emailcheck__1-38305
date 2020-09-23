VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Email Check"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   Icon            =   "EmailChecker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdcheckmail 
      Caption         =   "Check Mail"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Messages"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      Begin VB.TextBox txtNumOfMsgs 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message"
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   8775
      Begin RichTextLib.RichTextBox rtfMessage 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6588
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"EmailChecker.frx":000C
      End
   End
   Begin MSWinsockLib.Winsock sckCheck 
      Left            =   8280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum POP3State
 POP3_Connect
 POP3_USER
 POP3_PASS
 POP3_STAT
 POP3_RETR
 POP3_QUIT
End Enum

'Enumerates all the possible POP3 states that can occur while connected

Private MyState As POP3State

Dim MyUser As String
Dim MyPass As String
Dim MyHost As String
Dim NumOfMsgs As Integer
Dim MsgNum As Integer
Dim MsgCounter As Integer

Private Sub cmdcheckmail_Click()
 MyState = POP3_Connect
 sckCheck.Close
 sckCheck.LocalPort = 0
 sckCheck.Connect MyHost, 110
End Sub

Private Sub cmdExit_Click()
 sckCheck.Close
 rtfMessage.Text = ""
 End
End Sub

Private Sub Form_Load()
 MyUser = "username"
 MyPass = "pwd"
 MyHost = "hostname"
 MsgCounter = 0
End Sub

Private Sub sckCheck_DataArrival(ByVal bytesTotal As Long)
 
'All the code is here!
'each time new data arrives to the winsock command, a DATAARRIVAL event is raised
'all you have to do is to store the data and check at which stage of the cheching mail process you are
'while sending your user/pwd information the email server always reply with a "+" followed but something
'that can differ slightly from server to server
'the POP3 protocol accepts commands followed by RETURN KEY pressure, therefore don't forget
'to add vbCrLf at the end of every command

 
 Dim strData As String
 
 sckCheck.GetData strData, vbString

 If Left(strData, 1) = "+" Or MyState = POP3_RETR Then
  Select Case MyState
   Case POP3_Connect
    MyState = POP3_USER
    sckCheck.SendData "USER:" & MyUser & vbCrLf
    'enter your username
   Case POP3_USER
    MyState = POP3_PASS
    sckCheck.SendData "PASS:" & MyPass & vbCrLf
    'enter your password
   Case POP3_PASS
    MyState = POP3_STAT
    sckCheck.SendData "STAT" & vbCrLf
    'the STAT command asks the email server how many messages are stored
   Case POP3_STAT
    NumOfMsgs = CInt(Mid(strData, 5, InStr(5, strData, " ") - 5))
    'extract the number of messages from the server reply
    txtNumOfMsgs.Text = NumOfMsgs
    If NumOfMsgs > 0 Then
     MsgCounter = MsgCounter + 1
     sckCheck.SendData "RETR 1" & vbCrLf
     ' the RETR command retrieves the message
     MyState = POP3_RETR
    Else
     MyState = POP3_QUIT
     sckCheck.SendData "QUIT"
     rtfMessage.Text = "You have no message"
    End If
   Case POP3_RETR
    MsgBox MsgCounter
    Dim strBuffer As String
    strBuffer = strBuffer & strData
    If InStr(1, strBuffer, vbLf & "." & vbCrLf, vbTextCompare) Then
    'the "." simbol states the end of an email message
     rtfMessage.Text = rtfMessage.Text & strBuffer
     strBuffer = ""
     MsgCounter = MsgCounter + 1
     sckCheck.SendData "RETR " & MsgCounter & vbCrLf
     MyState = POP3_RETR
     If MsgCounter > NumOfMsgs Then
      sckCheck.SendData "QUIT"
      rtfMessage.Text = rtfMessage.Text & vbCrLf & "No more msg to download"
      sckCheck.Close
     End If
    End If
   Case POP3_QUIT
    sckCheck.SendData "QUIT"
    'to end a connection type QUIT followed by enter key
    sckCheck.Close
    rtfMessage.Text = rtfMessage.Text & strData
  End Select
 End If
End Sub

Private Sub sckCheck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 MsgBox "Error number: " & Number & " - " & Description, vbCritical, App.Title
End Sub
