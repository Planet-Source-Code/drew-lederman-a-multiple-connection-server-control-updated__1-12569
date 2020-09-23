VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sample Server Project"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   4725
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnTestServer 
      Caption         =   "&Test Server"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame frameLocalInfo 
      Caption         =   "Server IP/Port:"
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtIP 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtPort 
         Height          =   330
         Left            =   2400
         TabIndex        =   7
         Text            =   "420"
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView ServerOutput 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Output"
         Object.Width           =   8378
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Remote IP"
         Object.Width           =   3528
      EndProperty
   End
   Begin Project1.Server Server 
      Left            =   3480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   120
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Sto&p"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      Caption         =   "Server State:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   960
   End
   Begin VB.Label lblServerIP 
      AutoSize        =   -1  'True
      Caption         =   "Server IP:"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblConnections 
      AutoSize        =   -1  'True
      Caption         =   "0 Current Connections"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const white = &H80000005
Const grey = &H8000000F



Private Sub sOutput(strText As String, strIP As String)
    Dim itm As ListItem
    
    Set itm = ServerOutput.ListItems.Add(1, , strText)
    itm.SubItems(1) = strIP
    Set itm = Nothing
End Sub

Private Sub btnStart_Click()
    Call Server.StartServer(Val(txtPort), txtIP)
End Sub

Private Sub btnStop_Click()
    Call Server.StopServer
End Sub

Private Sub btnTestServer_Click()
    Dim frm As Form
    Set frm = New Form2
    
    frm.Show
End Sub


Private Sub Form_Load()
    txtIP = Server.ServerIP
End Sub

Private Sub Server_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    Call sOutput(FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved.", RemoteIP)
    
    If Data = "GET_TIME" Then
        Server.SendData Format(Time, "HH:MM:SS"), SckIndex
        Exit Sub
    End If
    
    If Data = "GET_DATE" Then
        Server.SendData Format(Date, "MM/DD/YYYY"), SckIndex
        Exit Sub
    End If
    
    Server.SendData "Invalid Command.", SckIndex
End Sub


Private Sub Server_Error(ByVal SckIndex As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
    Call sOutput("Server Error! (" & Description & ")", "N/A")
End Sub


Private Sub Server_ServerStarted()
    Call sOutput("Server Started! (" & Format(Time, "H:MM AM/PM") & ")", "N/A")
    btnStart.Enabled = False
    btnStop.Enabled = True
    btnTestServer.Enabled = True
    txtIP.Locked = True
    txtIP.BackColor = grey
    txtPort.Locked = True
    txtPort.BackColor = grey
End Sub


Private Sub Server_ServerStopped()
    Call sOutput("Server Stopped! (" & Format(Time, "H:MM AM/PM") & ")", "N/A")
    btnStart.Enabled = True
    btnStop.Enabled = False
    btnTestServer.Enabled = False
    txtIP.Locked = False
    txtIP.BackColor = white
    txtPort.Locked = False
    txtPort.BackColor = white
End Sub


Private Sub Server_SocketClosed(ByVal SckIndex As Integer, ByVal LocalPort As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    Call sOutput("Connection closed. ", RemoteIP)
End Sub


Private Sub Server_SocketOpened(ByVal SckIndex As Integer, ByVal LocalPort As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    Call sOutput("Connection opened. ", RemoteIP)
End Sub


Private Sub Server_StartFailed()
    Call sOutput("Failed to start server! ", "N/A")
End Sub


Private Sub Timer1_Timer()
    lblConnections = Server.ConnectionCount & " Current Connections"
    lblServerIP = "Server IP: " & Server.ServerIP
    lblState = "Server State: " & Server.State
End Sub


