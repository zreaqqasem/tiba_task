VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   4815
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   10800
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   855
      Left            =   7080
      TabIndex        =   5
      Top             =   7320
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      Text            =   "User Name"
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox txtPort 
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Text            =   "2525"
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox txtIP 
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   735
      Left            =   6960
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtSend 
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Text            =   "message.."
      Top             =   7440
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'connect to the server
Private Sub cmdConnect_Click()
If cmdConnect.Caption = "Connect" Then
sckMain.RemoteHost = txtIP.Text
sckMain.RemotePort = txtPort.Text
sckMain.Connect
List1.AddItem "You Are Joined The Group!"
cmdConnect.Caption = "DisConnect"
Else
sckMain.Close
List1.AddItem " Connection closed! " & vbCrLf & vbCrLf
cmdConnect.Caption = "Connect"
End If
End Sub

'Send new message
Private Sub cmdSend_Click()
If cmdConnect.Caption = "DisConnect" Then
sckMain.SendData " [ " & txtName.Text & ": ] " & txtSend.Text
txtSend.Text = ""
Else
MsgBox "Connect to the server before send message!"
End If
End Sub

'receive message from server and add it to chat list

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
sckMain.GetData Data
List1.AddItem Data
End Sub
