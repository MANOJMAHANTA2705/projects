VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton previousbtn 
      BackColor       =   &H00808000&
      Caption         =   "PREVIOUS"
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
      Left            =   720
      TabIndex        =   28
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT"
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
      Left            =   9720
      TabIndex        =   27
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton findbtn 
      Caption         =   "SEARCH"
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
      Left            =   6360
      TabIndex        =   26
      Top             =   1560
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton uploadbtn 
      BackColor       =   &H00808000&
      Caption         =   "UPLOAD PICTURE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   25
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
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
      Left            =   8040
      TabIndex        =   24
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "UPDATE"
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
      Left            =   6120
      TabIndex        =   23
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
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
      Left            =   4440
      TabIndex        =   22
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton addnew 
      Caption         =   "ADD"
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
      Left            =   2520
      TabIndex        =   21
      Top             =   7080
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   8880
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110952449
      CurrentDate     =   43647
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Text            =   "Select Semester"
      Top             =   5280
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   15
      Text            =   "Select Course"
      Top             =   4680
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Text            =   "Select Department"
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "STUDENT PROFILE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "PHONE NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "SEMESTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "COURSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "D.O.B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "ROLL NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Dim str As String



Private Sub addnew_Click()
rs.addnew
clear
End Sub
Sub display()
Text1.Text = rs!Rollno
Text2.Text = rs!Name
DTPicker1.Value = rs!DOB
If rs!Gender = "Male" Then
Option1.Value = True
Else
Option2.Value = True
End If
Combo1.Text = rs!Department
Combo2.Text = rs!Course
Combo3.Text = rs!Semester
Text3.Text = rs!Address
Text4.Text = rs!Phone
Picture1.Picture = LoadPicture(rs!Photo)
End Sub
Sub clear()
Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = "01/01/1996"
Option1.Value = False
Option2.Value = False
Combo1.Text = "Select Department"
Combo2.Text = "Select Course"
Combo3.Text = "Select Semester"
Text3.Text = ""
Text4.Text = ""
Picture1.Picture = LoadPicture("")
End Sub
Private Sub Combo1_Click()
Combo2.clear
If Combo1.Text = "Computer Science" Then
Combo2.AddItem "BCA"
Combo2.AddItem "MCA"
Combo2.AddItem "B.SC(COMPUTER SCIENCE)"
Combo2.AddItem "M.SC(COMPUTER SCIENCE)"
ElseIf Combo1.Text = "Geography" Then
Combo2.AddItem "B.SC(Geography)"
Combo2.AddItem "M.SC(Geography)"
Combo2.AddItem "B.A(Geography)"
Combo2.AddItem "M.A(Geography)"
Else
End If
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub deletebtn_Click()
confirm = MsgBox("Do you want to delete the student profile", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "record has been deleted successfully", vbInformation, "Messsge"
rs.Update
refreshdata
Else
MsgBox "Profile not deleted", vbInformation, "message"
End If
End Sub


Sub refreshdata()
rs.Close
rs.Open "select * from PofileTBL", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "no record found"
End If

End Sub

Private Sub findbtn_Click()
rs.Close
rs.Open "Select * from PofileTBL where Rollno = '" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Profile not found", vbInformation
End If
End Sub
Sub reload()
rs.Close
rs.Open "Select * from PofileTBL", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\manoj mahanta\Desktop\vbdbase.mdb;Persist Security Info=False"
rs.Open "Select * from PofileTBL", con, adOpenDynamic, adLockPessimistic
Combo1.AddItem "Computer Science"
Combo1.AddItem "Geography"
Combo3.AddItem "SEMESTER I"
Combo3.AddItem "SEMESTER II"
Combo3.AddItem "SEMESTER III"
Combo3.AddItem "SEMESTER IV"
Combo3.AddItem "SEMESTER V"
Combo3.AddItem "SEMESTER VI"
Combo3.AddItem "SEMESTER VII"
Combo3.AddItem "SEMESTER VIII"
Combo3.AddItem "SEMESTER IX"
Combo3.AddItem "SEMESTER X"
display

End Sub

Private Sub nextbtn_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If

End Sub

Private Sub previousbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If

End Sub

Private Sub savebtn_Click()
rs.Fields("Rollno").Value = Text1.Text
rs.Fields("Name").Value = Text2.Text
rs.Fields("DOB").Value = DTPicker1.Value
If Option1.Value = True Then
rs.Fields("Gender") = Option1.Caption
Else
rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Department").Value = Combo1.Text
rs.Fields("Course").Value = Combo2.Text
rs.Fields("Semester").Value = Combo3.Text
rs.Fields("Address").Value = Text3.Text
rs.Fields("Phone").Value = Text4.Text
rs.Fields("Photo").Value = str

MsgBox "Data is saved", vbInformation
rs.Update


End Sub

Private Sub updatebtn_Click()
rs.Fields("Rollno").Value = Text1.Text
rs.Fields("Name").Value = Text2.Text
rs.Fields("DOB").Value = DTPicker1.Value
If Option1.Value = True Then
rs.Fields("Gender") = Option1.Caption
Else
rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Department").Value = Combo1.Text
rs.Fields("Course").Value = Combo2.Text
rs.Fields("Semester").Value = Combo3.Text
rs.Fields("Address").Value = Text3.Text
rs.Fields("Phone").Value = Text4.Text
rs.Fields("Photo").Value = str

MsgBox "Data is saved", vbInformation
rs.Update
End Sub

Private Sub uploadbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"



str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
Picture1.AutoSize = False

End Sub
