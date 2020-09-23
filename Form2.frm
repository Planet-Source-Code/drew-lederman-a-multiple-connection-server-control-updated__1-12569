VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Sample"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtResponse 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "GET_TIME"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtIP 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   3240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Server's response:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Send data:"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   210
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Server's IP:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "To test the server control, you can connect to it and send commands (GET_TIME or GET_DATE)."
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3825
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConnect_Click()
    If btnConnect.Caption = "Connect" Then
        Winsock.Close
        Call Winsock.Connect(txtIP, Val(txtPort))
        
        Do While Winsock.State <> sckConnected: DoEvents
            'wait for socket to connect
            If Winsock.State = sckError Then Exit Sub
        Loop
        
        'connected
        btnConnect.Caption = "Disconnect"
        txtData.Enabled = True
        btnSend.Enabled = True
    ElseIf btnConnect.Caption = "Disconnect" Then
        Winsock.Close
        btnConnect.Caption = "Connect"
        txtData.Enabled = False
        btnSend.Enabled = False
    End If
End Sub

Private Sub btnSend_Click()
    If Winsock.State = sckConnected Then
        Winsock.SendData (txtData)
    End If
End Sub

Private Sub Form_Load()
    txtIP = Form1.Server.ServerIP
    txtPort = Form1.Server.ServerPort
End Sub




Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        btnSend_Click
    End If
End Sub


Private Sub Winsock_Close()
    'connection reset by the server
    Unload Me
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Call Winsock.GetData(Data, , bytesTotal)
    
    txtResponse.Text = Data
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbCritical, "Client Winsock Error: "
End Sub


