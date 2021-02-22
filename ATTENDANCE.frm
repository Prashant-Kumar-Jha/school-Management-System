VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FF8080&
   Caption         =   "ATTENDANCE"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BACK"
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
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9840
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1320
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\school management\atta.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\school management\atta.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "attandance"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   11040
      Width           =   4455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9840
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF80FF&
      DataField       =   "AFTERNOON"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   14
      Top             =   6600
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF80FF&
      DataField       =   "MORNING"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   13
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF80FF&
      DataField       =   "YEAR"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   4920
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF80FF&
      DataField       =   "MONTH"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "ATTENDANCE.frx":0000
      Left            =   5520
      List            =   "ATTENDANCE.frx":0028
      TabIndex        =   11
      Text            =   "CHOOSE MONTH"
      Top             =   4200
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF80FF&
      DataField       =   "DAY"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   10
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF80FF&
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF80FF&
      DataField       =   "REG NO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5520
      TabIndex        =   8
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   8820
      Left            =   9840
      Picture         =   "ATTENDANCE.frx":0090
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9600
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      Caption         =   "AFTERNOON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   6600
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      Caption         =   "MORNING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   5760
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Caption         =   "MONTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "DAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "REG NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "               ATTENDANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command10_Click()
main.Show
Form4.Hide
End Sub

Private Sub Command2_Click()
 Adodc1.Recordset.MoveNext
End Sub

Private Sub Command3_Click()
 Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command4_Click()
 Adodc1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
 Adodc1.Recordset.AddNew
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Fields("REG NO") = Text1.Text
Adodc1.Recordset.Fields("NAME") = Text2.Text
Adodc1.Recordset.Fields("DAY") = Text3.Text
Adodc1.Recordset.Fields("MONTH") = Combo1.Text
Adodc1.Recordset.Fields("YEAR") = Text4.Text
Adodc1.Recordset.Fields("MORNING") = Text5.Text
Adodc1.Recordset.Fields("AFTERNOON") = Text6.Text
End Sub

Private Sub Command7_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Combo1.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
End Sub

Private Sub Command8_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "delete record confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been delete sucessfully", vbInformation, "Message"
Else
MsgBox " Record not delete !!!", vbInformation, "Message"
End If
Adodc1.Recordset.Delete
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub
 Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo1.SetFocus
End Sub

Private Sub COMBO1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus
End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus
End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text6.SetFocus
End Sub
Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus

End Sub
