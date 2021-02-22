VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FF8080&
   Caption         =   " FEE DETAILS"
   ClientHeight    =   10725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19935
   LinkTopic       =   "Form5"
   ScaleHeight     =   10725
   ScaleWidth      =   19935
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      DataField       =   "MONTH"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      ItemData        =   "FEE DATAILS.frx":0000
      Left            =   4800
      List            =   "FEE DATAILS.frx":0028
      TabIndex        =   23
      Text            =   "CHOOSE MONTH"
      Top             =   4440
      Width           =   4215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3360
      Top             =   10200
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\school management\fee details.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\school management\fee details.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "fee "
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
      BackColor       =   &H00FF8080&
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
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF8080&
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
      Height          =   735
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
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
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
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
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9360
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "ADD "
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
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
      Height          =   735
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "LAST "
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "NEXT "
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
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
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9360
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFC0&
      DataField       =   "REASON"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   7560
      Width           =   4215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFC0&
      DataField       =   "FEE PAID"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   6600
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4800
      TabIndex        =   11
      Top             =   5400
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4800
      TabIndex        =   10
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
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
      Left            =   4800
      TabIndex        =   9
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
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
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   8475
      Left            =   9120
      Picture         =   "FEE DATAILS.frx":0090
      Top             =   1320
      Width           =   10215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "REASON"
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
      TabIndex        =   7
      Top             =   7680
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Caption         =   "FEE PAID"
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
      TabIndex        =   6
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
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
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
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
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
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
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
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
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "             FEE DEATAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
 Adodc1.Recordset.MoveNext
End Sub

Private Sub Command3_Click()
 Adodc1.Recordset.MoveLast
End Sub

Private Sub Command4_Click()
 Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command5_Click()
 Adodc1.Recordset.AddNew
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Fields("REG NO") = Text1.Text
Adodc1.Recordset.Fields("NAME") = Text2.Text
Adodc1.Recordset.Fields("DAY") = Text3.Text
Adodc1.Recordset.Fields("MONTH") = Combo1.Text
Adodc1.Recordset.Fields("YEAR") = Text5.Text
Adodc1.Recordset.Fields("FEE PAID") = Text6.Text
Adodc1.Recordset.Fields("REASON") = Text7.Text
End Sub

Private Sub Command7_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "delete record confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been delete sucessfully", vbInformation, "Message"
Else
MsgBox " Record not delete !!!", vbInformation, "Message"
End If
Adodc1.Recordset.Delete
End Sub

Private Sub Command()
End Sub

Private Sub Command9_Click()
main.Show
Form5.Hide
End Sub

Private Sub Command8_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Combo1.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
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



 Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text6.SetFocus
End Sub

 Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text7.SetFocus
End Sub

 Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus
End Sub
Private Sub COMBO1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus
End Sub
 
 

