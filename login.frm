VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "login page"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Frame1"
      ForeColor       =   &H8000000D&
      Height          =   6735
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
      Width           =   13095
      Begin VB.CommandButton Command2 
         Height          =   855
         Left            =   9480
         Picture         =   "login.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7080
         Picture         =   "login.frx":58CC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   6840
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3960
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         TabIndex        =   2
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "CONTACT SYSTEM ADMINISTRATION FOR FORGET PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   855
         Left            =   5400
         TabIndex        =   7
         Top             =   5880
         Width           =   7695
      End
      Begin VB.Image Image2 
         Height          =   6720
         Left            =   0
         Picture         =   "login.frx":B282
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5400
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Caption         =   "             PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   6840
         TabIndex        =   5
         Top             =   3240
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H008080FF&
         Caption         =   "            USERNAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   6840
         TabIndex        =   3
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   6960
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5280
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "                    PATNA MEDICINE AND RESEARCH HOSPITAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   1800
      Width           =   13095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Private Sub Command1_Click()
Dim username As String
Dim password As String
username = "admin"
password = "admin"
If Trim(LCase(Text1.Text)) = Trim(LCase("admin")) And Trim(LCase(Text2.Text)) = Trim(LCase("admin")) Then
MsgBox "login sucessful"
frmmain.Show
c = 0
Else
MsgBox "sorry......login faield......try again......"
c = c + 1
End If
If c = 3 Then
MsgBox "account is blocked"
End If

End Sub


Private Sub Command2_Click()
Text1.Text = " "
Text2.Text = " "
If Text1.Text = " " Then
MsgBox ("enter user name")
If Text2.Text = " " Then
MsgBox ("enter password")
End If
End If

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus

End Sub


Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus

End Sub
