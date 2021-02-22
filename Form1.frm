VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "LOGIN PAGE"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   9135
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   14775
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   8760
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   5040
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
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
         Left            =   8760
         TabIndex        =   2
         Top             =   2640
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "         PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8760
         TabIndex        =   4
         Top             =   4080
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "          USERNAME"
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
         Left            =   8760
         TabIndex        =   3
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   12000
         Left            =   240
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   -360
         Width           =   7560
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "                      SCHOOL MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   14760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim username As String
Dim password As String
username = "admin"
password = "admin"
If Trim(LCase(Text1.Text)) = Trim(LCase("admin")) And Trim(LCase(Text2.Text)) = Trim(LCase("admin")) Then
MsgBox "login sucessful"
main.Show
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

