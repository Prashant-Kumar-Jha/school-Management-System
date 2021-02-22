VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   Caption         =   "MARK DETAILS"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FF8080&
      Caption         =   "CALCULATE"
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9720
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6120
      Top             =   10560
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\school management\mark 3.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\school management\mark 3.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "mark "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF8080&
      Caption         =   "BACK"
      Height          =   495
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF8080&
      Caption         =   "DELETE"
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "ADD"
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "LAST"
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "NEXT "
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "FIRST"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9720
      Width           =   1155
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFF80&
      DataField       =   "DIVESION"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   8760
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFF80&
      DataField       =   "AVG"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   21
      Top             =   8040
      Width           =   3375
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFF80&
      DataField       =   "TOTAL"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   20
      Top             =   7320
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFF80&
      DataField       =   "SOC SCIENCE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   19
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFF80&
      DataField       =   "SCIENCE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFF80&
      DataField       =   "MATH"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   17
      Top             =   5040
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF80&
      DataField       =   "ENGLISH"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF80&
      DataField       =   "HINDI"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      DataField       =   "CLASS"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      DataField       =   "REG NO"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   8400
      Picture         =   "MARK.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   10320
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0FF&
      Caption         =   "DIVESION"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   8760
      Width           =   3855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0FF&
      Caption         =   "AVG "
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   8040
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TOTAL"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   3855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SOC SCIENCE"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SCIENCE"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "MATH"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ENGLISH"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "HINDI"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "NAME"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CLASS"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "REG NO"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "                 MARK DEATILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
Text9.Text = Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
Text10.Text = Text9.Text / 5
If Text10.Text >= 60 Then
Text11.Text = "FIRST"
ElseIf Text10.Text >= 50 Then
Text11.Text = "SECOND"
ElseIf Text10.Text >= 40 Then
Text11.Text = "THIRD"
Else
Text11.Text = "BETTER LUCK NEXT TIME..........."
End If
If Val(Text4.Text And Text5.Text And Text6.Text And Text7.Text And Text8.Text) < 30 Then
Text11.Text = "BETTER LUCK NEXT TIME..."
End If
End Sub

Private Sub Command1_Click()
 Adodc1.Recordset.MoveFirst
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
Adodc1.Recordset.Fields("CLASS") = Text2.Text
Adodc1.Recordset.Fields("NAME") = Text3.Text
Adodc1.Recordset.Fields("HINDI") = Text4.Text
Adodc1.Recordset.Fields("ENGLISH") = Text5.Text
Adodc1.Recordset.Fields("MATH") = Text6.Text
Adodc1.Recordset.Fields("SCIENCE") = Text7.Text
Adodc1.Recordset.Fields("SOC SCIENCE") = Text8.Text
Adodc1.Recordset.Fields("TOTAL") = Text9.Text
Adodc1.Recordset.Fields("AVG") = Text10.Text
Adodc1.Recordset.Fields("DEVESION") = Text11.Text

End Sub


Private Sub Command7_Click()
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text4.Text = " "
Text11.Text = " "
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

Private Sub Command9_Click()
main.Show
Form3.Hide
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus

End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus

End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus

End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus

End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text6.SetFocus

End Sub
Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text7.SetFocus

End Sub
Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text8.SetFocus

End Sub
Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text9.SetFocus

End Sub
Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text10.SetFocus

End Sub
Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text11.SetFocus

End Sub
Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus
End Sub
