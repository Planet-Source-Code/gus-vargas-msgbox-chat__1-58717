VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MsgBox Chat v1.0 - Server - By: Gus Vargas"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtMsg 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Text            =   "666"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Disconnected"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MsgBox Chat v1.0 Source by Gus Vargas
'This source is to help newbie programmers understand how Winsock works
'I also had lots of trouble with Winsock, and still struggle with it
'Thanks for downloading
'
'gus.sytes.net

Private Sub cmdConnect_Click()
On Error Resume Next
Winsock1.Close 'Make sure the socket is closed before connecting and prevent an error
Winsock1.LocalPort = txtPort 'Set the port to listen for connection on
Winsock1.Listen 'Listen for connections

cmdConnect.Enabled = False
cmdDisconnect.Enabled = True
txtPort.Enabled = False
txtName.Enabled = False
lblStatus = "Status: Listening on port " & txtPort & "..."
End Sub

Private Sub cmdDisconnect_Click()
Winsock1.Close 'Close the connection

cmdConnect.Enabled = True
cmdDisconnect.Enabled = False
cmdSend.Enabled = False
txtPort.Enabled = True
txtName.Enabled = True
lblStatus = "Status: Disconnected"
End Sub

Private Sub cmdSend_Click()
Winsock1.SendData txtName & ": " & txtMsg 'Send txtName and ": " and txtMsg, this line together should look like this once sent, ex. "Gus: Hi"
txtMsg = "" 'Clear txtMsg once the message has been sent
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Winsock1_Close()
txtPort.Enabled = True
txtName.Enabled = True

cmdConnect.Enabled = True
cmdDisconnect.Enabled = False
cmdSend.Enabled = False

lblStatus = "Status: Disconnected" 'Just to be sure the connection has been closed
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close 'When the server recieves a connection request stop listening

Winsock1.Accept requestID 'Accept the request

lblStatus = "Status: Connected" 'Make sure the connection has been accepted
cmdSend.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim msg As String 'Set 'msg' as a string
Winsock1.GetData msg 'Recieve the data that has been sent to you and set it as 'msg'
MsgBox msg, vbOKOnly, "MsgBox Chat v1.0" ''Put the data in a MsgBox to view
End Sub
