VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "bnet.client"
   ClientHeight    =   3465
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
   ScaleHeight     =   3465
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
      Left            =   2640
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1680
      Top             =   2400
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   3585
      TabIndex        =   7
      Top             =   0
      Width           =   1095
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
      Left            =   2040
      Tag             =   "CLOSED"
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   8174
   End
   Begin VB.Label lblSend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Data"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Data"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status: status here..."
      Height          =   255
      Left            =   15
      TabIndex        =   4
      Top             =   255
      Width           =   4695
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblClient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
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
      Width           =   480
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
'
'Code edited by drk||Raziel


Public Sub SendData(data As String)
'Check to see if we're connected to the server
If Winsock.Tag = "CONNECTED" Then
    'Send the data
    Winsock.SendData data & vbCrLf
    
    'Send the data go the textbox as well
    txtData = txtData & "CLIENT> " & data & vbCrLf
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

Private Sub cmdConnect_Click()
Dim ip As String

If cmdConnect.Caption = "Connect" Then
    'If we want to connect, first ask the user for the
    'server's IP
    ip = InputBox("Enter the server's IP:", "Enter IP")
    'If they didn't cancel, connect to the server
    If ip <> "" Then
        'Close winsock
        Winsock.Close
        
        'Tell winsock what it's connecting to
        Winsock.RemoteHost = ip
        Winsock.RemotePort = 8179 'and what port to use
        
        'Connect
        Winsock.Connect
        cmdConnect.Caption = "Disconnect"
        Exit Sub
    End If
Else
    'Close the winsock
    Winsock.Close
    'Do the code that is in Winsock's Close sub
    Winsock_Close
    cmdConnect.Caption = "Connect"
End If
End Sub

Private Sub cmdSend_Click()
'If text isn't blank, send data to the server
If txtSend.Text <> "" Then
    SendData txtSend.Text
    txtSend.Text = ""
End If
End Sub

Private Sub Command1_Click()
Me.Caption = "bnet.client"
bbAC.link_kill
bbAC.Con
End Sub

Private Sub Form_Load()
bbAC.link_kill
bbAC.Con
Status "Idle.."
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
'Server closed connection, close here as well
Winsock.Close
Winsock.Tag = "CLOSED"

'Update status
Status "Disconnected, Idle.."
End Sub

Private Sub Winsock_Connect()
'We've connected to the server!
Winsock.Tag = "CONNECTED"

'Update status
Status "Connected"
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String, bbb(1) As Byte
'Update status
Status "Data has arrived"

'Get the incoming data, which was
'sent from the server
Winsock.GetData Buffer
If Asc(Buffer) = 9 Then
bbb(0) = 1
bbb(1) = Asc(Mid$(Buffer, 2))
Send bbb
Else
'Send it to the textbox
txtData = txtData & "SERVER> " & Buffer
End If
'Update status
Status "Connected"
End Sub
