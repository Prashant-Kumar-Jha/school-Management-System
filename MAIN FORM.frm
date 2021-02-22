VERSION 5.00
Begin VB.MDIForm main 
   BackColor       =   &H00FF8080&
   Caption         =   "main form"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   Picture         =   "MAIN FORM.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu STU 
      Caption         =   "STUDENT DETAILS"
   End
   Begin VB.Menu MARK 
      Caption         =   "MARK DETAILS"
   End
   Begin VB.Menu FEE 
      Caption         =   "FEE DETAILLS"
   End
   Begin VB.Menu ATTA 
      Caption         =   "ATTENDANCE"
   End
   Begin VB.Menu ABOUT 
      Caption         =   "ABOUT"
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ABOUT_Click()
Form6.Show
End Sub

Private Sub ATTA_Click()
Form4.Show
End Sub

Private Sub EXIT_Click()
confirmation = MsgBox("Do you want to EXIT", vbYesNo + vbCritical, "EXIT confirmation")
If confirmation = vbYes Then
End
MsgBox "EXIT sucessfully", vbInformation, "Message"
Else
MsgBox " EXIT NOT SUCESSFULLY !!!", vbInformation, "Message"
End If
End
End Sub

Private Sub FEE_Click()
Form5.Show
End Sub

Private Sub MARK_Click()
Form3.Show
End Sub

Private Sub STU_Click()
Form2.Show
End Sub
