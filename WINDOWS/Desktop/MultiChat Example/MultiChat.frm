VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMultiChat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Basics of a multi user chat room"
   ClientHeight    =   3615
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4950
   Icon            =   "MultiChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   2520
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   2160
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMessage 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4695
   End
   Begin VB.TextBox txtChatWindow 
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblUsersConnected 
      Caption         =   "Total users connected: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "Co&nnection"
      Begin VB.Menu mnuStartServer 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu mnuConnectAsClient 
         Caption         =   "Connect As &Client"
      End
      Begin VB.Menu mnuLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEndConnection 
         Caption         =   "&End Connection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMultiChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is increased every time a NEW user connects to the
'server.
Dim SocketCount As Integer
'This counts how many users are connected to the server
Dim TotalUsersConnected As Integer

Private Sub mnuConnectAsClient_Click()

'Set the remote port, this has to be the same as the server
'port
wsClient.RemotePort = 789

'For this example we are connecting to our selves but you can
'change this to what ever you want E.g.

'wsClient.Connect "127.174.45.90"

wsClient.Connect wsClient.LocalIP

'**This is all that is needed to connect the client to the server
'**the rest below is just handeling menu buttons and other stuff**

'**********************************************************************
'**********************************************************************

'Set the forms caption so that we know what we are
'connected as
frmMultiChat.Caption = "Connected As Client"

'As seen as we are the client we cant connect as the server
'so disable the Start Server button
mnuConnectAsClient.Enabled = False

'We have connected as the client so disable the Connect As'
'Client button
mnuStartServer.Enabled = False

'We have started connected to the server so enable the
'End Connection Button
mnuEndConnection.Enabled = True

End Sub

Private Sub mnuEndConnection_Click()

'Close bothe connections
wsServer(0).Close
wsClient.Close

'Enable the Start Server button in the menu
mnuStartServer.Enabled = True

'Enable the Connect As Client button in the menu
mnuConnectAsClient.Enabled = True

'We have no connections to close so disable this button
mnuEndConnection.Enabled = False

'Hide the How Many users connectd label
lblUsersConnected.Visible = False

'Set the forms caption to what it was when the program first
'started
frmMultiChat.Caption = "Basics of a multi user chat room"

'Set the total users connected back to 0
TotalUsersConnected = 0

'Set the Users Connected label back
lblUsersConnected.Caption = "Total users connected: 0"

'Clear the chat window
txtChatWindow.Text = ""

End Sub

Private Sub mnuExit_Click()

End

End Sub

Private Sub mnuStartServer_Click()

'Set the port that we are going to be chatting through
wsServer(0).LocalPort = 789

'Listen for users to connect
wsServer(0).Listen

'Why is this here?: This is here because the server connects
'to its self..Why? because if you look in the DateArrival function
'for wsServer controle you will see that all the messages from
'the clients are sent to the server and then the server sends
'out the message to the clients connected. So for the server
'to recive the message it must connect to its self so that it
'will be in the list of connected clients.:)..If you want the server
'to be just a server, not a chat client as well take this out.
mnuConnectAsClient_Click

'**That is all that is needed to start the server the rest below is
'**is just handeling the menu buttons and other stuff**

'********************************************************************
'********************************************************************

'Set the forms caption so that we know what we are
'connected as
frmMultiChat.Caption = "Connected As Server"

'As you are the server you cant connect as a client so
'disable the menu button
mnuConnectAsClient.Enabled = False

'We have started the server so disable to start server
'button
mnuStartServer.Enabled = False

'The server has started so enable the End Connection
'button on the menu
mnuEndConnection.Enabled = True

'Display the How many users are connected label
lblUsersConnected.Visible = True

End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)

'If the Enter button is pressed send the message
If KeyAscii = 13 Then
    wsClient.SendData wsClient.LocalHostName & ": " & txtMessage.Text
    DoEvents 'Let it finish
    'Clear the text box ready for next message
    txtMessage.Text = ""
End If

End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)

Dim strDataRecived As String

'Recive the message that has just been sent
wsClient.GetData strDataRecived
DoEvents

'Display the message in the chat window
txtChatWindow.Text = txtChatWindow.Text & strDataRecived & vbCrLf

End Sub

Private Sub wsServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

'Every time a new user connects increase the socket count
'by 1
SocketCount = SocketCount + 1

'Use the load function to load another winsock controle for
'this user.
Load wsServer(SocketCount)
'Once we have loaded the Winsock controle accept the
'request from the user that is trying to connect
wsServer(SocketCount).Accept requestID

'**Thats all that is needed to connect multiple clients to a
'**server.
'**********************************************************************
'**********************************************************************

'Increase the user count
TotalUsersConnected = TotalUsersConnected + 1

'Display the total users connected the minus 1 is because the
'server is connected to its self.
lblUsersConnected.Caption = "Total users connected: " & TotalUsersConnected - 1

'The only thing with this is that it will load winsock controle
'after winsock controle it wont look at the past winsock contoles
'to see if no one is connected to it

End Sub


Private Sub wsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim strRecivedData As String
Dim SocketCheck As Integer

'Recive the message from the client
wsServer(Index).GetData strRecivedData

'What this for statement does is go through all the winsocks
'that we have open and make sure that they are connected to
'a client. If they are then send the message to the client
For SocketCheck = 0 To SocketCount Step 1
        'If the winsocks state is Connected then send the message
        'to that client.
        If wsServer(SocketCheck).State = sckConnected Then
                wsServer(SocketCheck).SendData strRecivedData
                DoEvents
        End If
Next SocketCheck

End Sub
