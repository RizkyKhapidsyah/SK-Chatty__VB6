VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Chatty"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "E-mail me at :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblemail 
      Caption         =   "allegro16@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      MouseIcon       =   "frmabout.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   240
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Cyber Chatting In Your Hands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lbltitle 
      Caption         =   "Chatty v1.00 Build 0100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub Form_Load()
With txtinfo
.Text = txtinfo.Text & "Designed and Developed by Benny T." & vbNewLine
.Text = txtinfo.Text & "Modified for Learn by Rizky Khapidsyah." & vbNewLine
.Text = txtinfo.Text & "Data Transmission: TCP/IP - Winsock" & vbNewLine
.Text = txtinfo.Text & "Port: 45660 TCP" & vbNewLine
.Text = txtinfo.Text & "Data Security: RSA 64-bit Public Key Cipher" & vbNewLine
.Text = txtinfo.Text & "Uses Secure Communication Channel (SCC)" & vbNewLine
End With
End Sub

Private Sub lblemail_Click()
d = ShellExecute(0, vbNullString, "mailto:allegro16@hotmail.com", vbNullString, vbNullString, vbNormalFocus)
End Sub
