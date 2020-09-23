VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmNetTest 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Width           =   4215
   End
   Begin VB.TextBox txtRcv 
      Height          =   2775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1800
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock wskOut 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskIn 
      Index           =   0
      Left            =   3480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btnStartHost 
      Caption         =   "Start Host"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtServPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtClientPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtClientAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "User Name:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Receive:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Host on Port"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Connect to Port"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Connect to Address"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmNetTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSend_Click()
    If (txtClientPort.Text = "") Or (txtClientAddress.Text = "") Then
        MsgBox "You must specify both the address and port to connect to"
        Exit Sub
    End If

    'Attempt the connection
    wskOut.Connect txtClientAddress.Text, Int(txtClientPort.Text)
End Sub

Private Sub btnStartHost_Click()
    If txtServPort.Text = "" Then
        MsgBox "You must set the port to host on"
        Exit Sub
    End If

    'Setup the port to listen on
    wskIn(0).LocalPort = Int(txtServPort.Text)
    'Begin listening
    wskIn(0).Listen
    
    'Check to make sure we are listening
    If (wskIn(0).State <> sckListening) Then
        MsgBox "Unable to begin hosting"
        End
    End If

    btnStartHost.Enabled = False
End Sub

Private Sub Form_Load()
    txtClientAddress.Text = "127.0.0.1"
    txtClientPort.Text = "2000"
    txtServPort.Text = "2000"
    
    txtName.Text = "Jason"
    txtSend.Text = "Test"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Verify that both connections are closed before exiting
    If wskIn(0).State <> sckClosed Then
        wskIn(0).Close
    End If
    
    If wskOut.State <> sckClosed Then
        wskOut.Close
    End If
End Sub

Private Sub txtRcv_Change()
    txtRcv.SelStart = Len(txtRcv.Text)
End Sub

Private Sub wskIn_Close(Index As Integer)
    'Remove the winsock object since its no longer necessary
    Unload wskIn(Index)
End Sub

Private Sub wskIn_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'Create a new Winsock object to handle the incoming request
    Load wskIn(1)

    'Verify that the new winsock object is closed
    If wskIn(1).State <> sckClosed Then
        wskIn(1).Close
    End If

    'Accept the incoming connection
    wskIn(1).Accept requestID
End Sub

Private Sub wskIn_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String

    'If the winsock object is not connected it will error out when you attempt
    'to receive the data
    If (wskIn(Index).State <> sckConnected) Then
        Exit Sub
    End If
    
    'Accept the incoming data
    wskIn(Index).GetData strData
    
    txtRcv.Text = txtRcv.Text & Chr(13) & Chr(10) & strData
    
    Beep
End Sub

Private Sub wskOut_Connect()
    txtRcv.Text = txtRcv.Text & Chr(13) & Chr(10) & txtName.Text & ": " & txtSend.Text
    
    'Send your data now that you are conneced
    wskOut.SendData txtName.Text & ": " & txtSend.Text
End Sub
    
Private Sub wskOut_SendComplete()
    'The data is sent close the connection
    wskOut.Close

    txtSend.Text = ""
End Sub
