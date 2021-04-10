VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chatty - Main"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmessage 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Message"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtchat 
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdabout 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection Status"
      Height          =   1215
      Left            =   4560
      TabIndex        =   15
      Top             =   1320
      Width           =   2655
      Begin VB.Label lblconnection 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCC"
      Height          =   1215
      Left            =   4560
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
      Begin VB.Label lblscc 
         BackStyle       =   0  'Transparent
         Caption         =   " Disabled"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblsecure 
         Caption         =   "Secure Communication Channel"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   840
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtsay 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   6975
   End
   Begin VB.OptionButton opttype 
      Caption         =   "Server"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   7
      Top             =   4680
      Width           =   855
   End
   Begin VB.OptionButton opttype 
      Caption         =   "Client"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   4680
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox chksecure 
      Caption         =   "Secure Communication"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   5400
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtip 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmddisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdlisten 
      Caption         =   "Listen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lbltitle 
      Height          =   735
      Left            =   4680
      TabIndex        =   11
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lbltext 
      Caption         =   "Text to Send:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label lblip 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dot As Byte
Dim nm As String
Dim chatter As String
Dim cht As String

Private Sub chksecure_Click()
If chksecure.Value = Checked Then
  Shape1.FillColor = &H8000&
  lblscc.Caption = "   Active"
ElseIf chksecure.Value = Unchecked Then
  Shape1.FillColor = &H808080
  lblscc.Caption = " Disabled"
End If
End Sub

Private Sub cmdabout_Click()
frmabout.Show
End Sub

Private Sub cmdclear_Click()
txtchat.Text = ""
End Sub

Private Sub cmdconnect_Click()
dot = 0
If txtip.Text = "" Then
  MsgBox "Please enter an IP address", vbCritical, "Error !"
  Exit Sub
End If
For k = 1 To Len(txtip.Text)
  a = Mid(txtip.Text, k, 1)
  If a = "." Then
    dot = dot + 1
  End If
Next
If dot <> 3 Then
  MsgBox "Invalid IP address format", vbCritical, "Error !"
  Exit Sub
End If
ws.RemotePort = 45660
ws.RemoteHost = txtip.Text
ws.Connect
cmdconnect.Enabled = False
cmddisconnect.Enabled = True
txtip.Enabled = False
txtip.BackColor = &HC0C0C0
For k = 0 To 1
  opttype(k).Enabled = False
Next
lblconnection.Caption = "Connecting to " & ws.RemoteHost & "...."
End Sub

Private Sub cmddisconnect_Click()
If ws.State <> sckClosed Then
  If ws.State = sckConnected Then
    ws.SendData "BYE"
  End If
  DoEvents
  ws.Close
  txtchat.Text = txtchat.Text & "Chat ended - " & Date & ", " & Time & vbNewLine
  lblconnection.Caption = " Ready"
  cmddisconnect.Enabled = False
  If opttype(1).Value = True Then
    cmdlisten.Enabled = True
    cmdconnect.Enabled = False
    txtip.Enabled = False
    txtip.BackColor = &HC0C0C0
  ElseIf opttype(0).Value = True Then
    cmdlisten.Enabled = False
    cmdconnect.Enabled = True
    txtip.Enabled = True
    txtip.BackColor = vbWhite
  End If
  For k = 0 To 1
    opttype(k).Enabled = True
  Next
  txtsay.Enabled = False
  txtsay.BackColor = &HC0C0C0
End If
End Sub

Private Sub cmdexit_Click()
If ws.State = sckConnected Then
  s = MsgBox("You are currently connected. Are you sure you want to quit ?", vbInformation + vbYesNo, "Confirm Exit")
  If s = vbYes Then
    Unload Me
  End If
End If
Unload Me
End Sub

Private Sub cmdlisten_Click()
ws.LocalPort = 45660
ws.Listen
If ws.State = sckListening Then
  cmdlisten.Enabled = False
  cmddisconnect.Enabled = True
  For k = 0 To 1
    opttype(k).Enabled = False
  Next
  lblconnection.Caption = "Listening on port " & ws.LocalPort & "...."
End If
End Sub

Private Sub cmdmessage_Click()
frmmessage.Show
End Sub

Private Sub Form_Load()
Open "name.cfg" For Input As #1
Line Input #1, nm
Close #1
Me.Caption = "Chatty - " & nm
lbltitle.Caption = "Chatty v1.00 Build 0100" & vbNewLine
lbltitle.Caption = lbltitle.Caption & vbNewLine & "Local IP Address: " & ws.LocalIP
lblconnection.Caption = " Ready"
txtsay.Enabled = False
txtsay.BackColor = &HC0C0C0
key(1) = 35429567
key(2) = 21444671
key(3) = 31393357
p = 3613
q = 8689
PHI = 31381056
End Sub

Private Sub opttype_Click(Index As Integer)
Select Case Index
Case 0:
  txtip.Enabled = True
  txtip.BackColor = vbWhite
  cmdlisten.Enabled = False
  cmdconnect.Enabled = True
Case 1:
  txtip.Enabled = False
  txtip.BackColor = &HC0C0C0
  cmdconnect.Enabled = False
  cmdlisten.Enabled = True
End Select
End Sub

Private Sub txtsay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If ws.State = sckConnected Then
    If chksecure.Value = Checked Then
      ws.SendData enc(txtsay.Text, key(1), key(3)) & "TXT-S"
    Else
      ws.SendData txtsay.Text & "TXT"
    End If
    txtchat.Text = txtchat.Text & nm & ": " & txtsay.Text & vbNewLine
  Else
    MsgBox "No connection. Unable to send message", vbCritical, "Error !"
  End If
  txtsay.Text = ""
End If
End Sub

Private Sub ws_Connect()
lblconnection.Caption = "Connected to " & ws.RemoteHost & " on port 45660"
If chksecure.Value = Checked Then
  ws.SendData enc(nm, key(1), key(3)) & "NM-S"
Else
  ws.SendData nm & "NM"
End If
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
If ws.State <> sckClosed Then
  ws.Close
End If
ws.Accept requestID
lblconnection.Caption = "Connected to " & ws.RemoteHostIP
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
DoEvents
ws.GetData dat, vbString

If dat = "BYE" Then
  txtchat.Text = txtchat.Text & "Chat ended - " & Date & ", " & Time & vbNewLine
  If ws.State <> sckClosed Then
    ws.Close
  End If
  lblconnection.Caption = " Ready"
  cmddisconnect.Enabled = False
  If opttype(1).Value = True Then
    cmdlisten.Enabled = True
    cmdconnect.Enabled = False
    txtip.Enabled = False
    txtip.BackColor = &HC0C0C0
  ElseIf opttype(0).Value = True Then
    cmdlisten.Enabled = False
    cmdconnect.Enabled = True
    txtip.Enabled = True
    txtip.BackColor = vbWhite
  End If
  For k = 0 To 1
    opttype(k).Enabled = True
  Next
  txtsay.Enabled = False
  txtsay.BackColor = &HC0C0C0
End If

If (Right(dat, 5) = "MSG-S") Or (Right(dat, 3) = "MSG") Then
  If Right(dat, 1) = "S" Then
    recv = dec(Mid(dat, 1, Len(dat) - 5), key(2), key(3))
  Else
    recv = Mid(dat, 1, Len(dat) - 3)
  End If
  MsgBox recv, , "Message from " & chatter
End If

If (Right(dat, 4) = "RP-S") Or (Right(dat, 2) = "RP") Then
  If Right(dat, 1) = "S" Then
    chatter = dec(Mid(dat, 1, Len(dat) - 4), key(2), key(3))
  Else
    chatter = Mid(dat, 1, Len(dat) - 2)
  End If
  txtchat.Text = ""
  txtchat.Text = txtchat.Text & "Chat started - " & Date & ", " & Time & vbNewLine
  txtsay.Enabled = True
  txtsay.BackColor = vbWhite
End If

If (Right(dat, 4) = "NM-S") Or (Right(dat, 2) = "NM") Then
  If Right(dat, 1) = "S" Then
    chatter = dec(Mid(dat, 1, Len(dat) - 4), key(2), key(3))
    ws.SendData enc(nm, key(1), key(3)) & "RP-S"
  Else
    chatter = Mid(dat, 1, Len(dat) - 2)
    ws.SendData nm & "RP"
  End If
  txtchat.Text = ""
  txtchat.Text = txtchat.Text & "Chat started - " & Date & ", " & Time & vbNewLine
  txtsay.Enabled = True
  txtsay.BackColor = vbWhite
End If

If (Right(dat, 5) = "TXT-S") Or (Right(dat, 3) = "TXT") Then
  If Right(dat, 1) = "S" Then
    cht = dec(Mid(dat, 1, Len(dat) - 5), key(2), key(3))
    txtchat.Text = txtchat.Text & chatter & ": " & cht & vbNewLine
  Else
    cht = Mid(dat, 1, Len(dat) - 3)
    txtchat.Text = txtchat.Text & chatter & ": " & cht & vbNewLine
  End If
End If
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, "Error !"
If ws.State <> sckClosed Then
  ws.Close
End If
  lblconnection.Caption = " Ready"
  cmddisconnect.Enabled = False
  If opttype(1).Value = True Then
    cmdlisten.Enabled = True
    cmdconnect.Enabled = False
    txtip.Enabled = False
    txtip.BackColor = &HC0C0C0
  ElseIf opttype(0).Value = True Then
    cmdlisten.Enabled = False
    cmdconnect.Enabled = True
    txtip.Enabled = True
    txtip.BackColor = vbWhite
  End If
  For k = 0 To 1
    opttype(k).Enabled = True
  Next
End Sub

