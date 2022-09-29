VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox dataList 
      Height          =   5520
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock server 
      Index           =   0
      Left            =   7680
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Server Logs"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Online Users: 0"
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   2520
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connectionnumber As Integer

'heres the local port which is 2525 and can be any

Private Sub Form_Load()
server(0).LocalPort = "2525"
server(0).Listen
End Sub

'close the server

Private Sub server_Close(Index As Integer)
dataList.AddItem "User : " & server(Index).Index & "  Just Leave Socket! "
server(Index).Close
Unload server(Index)
connectionnumber = connectionnumber - 1
End Sub

'accept clients conenction requests

Private Sub server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
connectionnumber = connectionnumber + 1
Label1.Caption = "Online Users = : " & connectionnumber
Load server(connectionnumber)
server(connectionnumber).Accept requestID
dataList.AddItem "User  " & requestID & "  Just Join Socket!"
End Sub

'receive data from clients and send it to every one in the server

Private Sub server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String
server(Index).GetData data
dataList.AddItem data
dataList.AddItem "****** Data Sending from Server *******"
For i = 1 To connectionnumber
server(i).SendData data
dataList.AddItem "data sent to client " & i
Next
dataList.AddItem "****** Data Sending from Server *******"
End Sub

