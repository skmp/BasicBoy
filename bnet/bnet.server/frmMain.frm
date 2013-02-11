VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "bnet.server"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reinit Link connection"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   3000
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2040
      Top             =   2280
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "send"
      Default         =   -1  'True
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   2640
      Tag             =   "CLOSED"
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8179
   End
   Begin VB.Label lblSend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Data"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Data"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status: status here..."
      Height          =   255
      Left            =   15
      TabIndex        =   5
      Top             =   255
      Width           =   4695
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblIP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4650
      TabIndex        =   4
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   570
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)   -Patrick             *
'**********************************
'
'PS: Please look for more submissions to PSC by me
'    shortly.  I've recently been working on a lot
'    :))  All my submissions are under author name
'    "Patrick Moore"
'    Vote for me..I'll be happy :DD
'
'Code edited by drk||Raziel

Sub SendData(data As String)
'Check to see if we're connected to a client
If Winsock.Tag = "CONNECTED" Then
    'Send the data
    Winsock.SendData data & vbCrLf
    
    'Send the outgoing data to the textbox as well
    txtData = txtData & "SERVER> " & data & vbCrLf
End If
End Sub
Public Sub SendData2(data As String)
'Check to see if we're connected to the server
If Winsock.Tag = "CONNECTED" Then
    'Send the data
    Winsock.SendData data & vbCrLf
End If
End Sub

Sub Status(data As String)
'Update the status label
lblStatus.Caption = "Status: " & data
End Sub

Private Sub cmdSend_Click()
'Send the data to the client if it
'is not blank
If txtSend.Text <> "" Then
    SendData txtSend.Text
    txtSend.Text = ""
End If
End Sub

Private Sub Command1_Click()
Me.Caption = "bnet.server"
bbAC.link_kill
bbAC.Con
End Sub

Private Sub Form_Load()
bbAC.link_kill
bbAC.Con
'Set the caption with your IP
lblIP.Caption = "your ip: " & Winsock.LocalIP

'Listen for incoming connection requests
Winsock.Listen
Status "Awaiting connection.."
End Sub

Private Sub Timer1_Timer()
bbAC.check_link_connection
End Sub

Private Sub Timer2_Timer()
Check
End Sub

Private Sub txtData_Change()
'Set the cursor to the last character of the textbox
txtData.SelStart = Len(txtData.Text)
End Sub

Private Sub Winsock_Close()
'Client closed connection, close the Winsock on this side
Winsock.Close
Winsock.Tag = "CLOSED"

'Update status
Status "Connection closed, awaiting new connection.."

'Re-listen for incoming connection requests
Winsock.Listen
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
'Update status
Status "Accepting connection request"

'Close winsock
Winsock.Close

'Accept the connection request
Winsock.Accept requestID
Winsock.Tag = "CONNECTED"

'Update status
Status "Connected"
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String, bbb(1) As Byte
'Update status
Status "Data has arrived"


'Get the data being sent by the client
Winsock.GetData Buffer
If Asc(Buffer) = 9 Then
bbb(0) = 1
bbb(1) = Asc(Mid$(Buffer, 2))
Send bbb
Else
'Put incoming data into the Data textbox
txtData = txtData & "CLIENT> " & Buffer
End If
'Instead of using this as a chat program, you could use it
'as a remote network tool, etc.

Buffer = UCase(Buffer)
Buffer = Left(Buffer, Len(Buffer) - 2)


'Update the status back to "Connected"
Status "Connected"
End Sub
