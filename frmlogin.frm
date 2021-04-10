VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One Moment Please ....."
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblname 
      Caption         =   "Please enter a name to be used in this chat session"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
If txtname.Text = "" Then
  MsgBox "Please enter a name", vbCritical, "Error !"
  Exit Sub
End If
Open "name.cfg" For Output As #1
Print #1, txtname.Text
Close #1
frmmain.Show
Unload Me
End Sub

Private Sub Form_Load()
If Dir("name.cfg") <> "" Then
  Open "name.cfg" For Input As #1
  Line Input #1, nm
  Close #1
  txtname.Text = nm
  txtname.SelLength = Len(txtname.Text)
End If
End Sub
