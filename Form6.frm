VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   Caption         =   "NEW EMPLOYEE"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form6"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton prevbtn 
      Caption         =   "PREV"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   23
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   22
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   21
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   20
      Top             =   6120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15480
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton uploadbtn 
      Caption         =   "UPLOAD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15240
      TabIndex        =   19
      Top             =   3840
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   5760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7200
      TabIndex        =   15
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5040
      TabIndex        =   14
      Top             =   7800
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11880
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "employee"
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
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   13
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton addnew 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   960
      TabIndex        =   12
      Top             =   7800
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "DESIG"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11040
      TabIndex        =   11
      Text            =   "SELECT DESIGNATION"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      DataField       =   "EADDRESS"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   11040
      TabIndex        =   10
      Top             =   4320
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "EPHOTO"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   1935
      Left            =   15240
      ScaleHeight     =   1875
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "EMOBILENO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      DataField       =   "EID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "ENAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "MOBILE NO."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "DESIGNATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW EMPLOYEE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form6"
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
Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
Combo1.Text = "SELECT DESIGNATION"
Text4.Text = ""
Picture1.Picture = LoadPicture("")

End Sub

Private Sub deletebtn_Click()
confirm = MsgBox("Do you want to Delete????", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted Successfully", vbInformation, "message"
rs.Update
refreshdata

Else
MsgBox "Record not deleted", vbInformation, "Message"
End Sub
Sub refreshdata()
rs.Close
rs.Open "select * From employee", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No Record Found"
End If
End Sub

Private Sub firstbtn_Click()
rs.MoveFirst
display

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
rs.Open "select * from employee", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "Accountant"
Combo1.AddItem "salesman"
Combo1.AddItem "clerk"

display

End Sub
Sub display()
Text1.Text = rs!ENAME
Text2.Text = rs!EID
Text3.Text = rs!EMOBILENO
If rs!GENDER = "MALE" Then
Option1.Value = True
Else
Option2.Value = True
End If
Combo1.Text = rs!DESIG
Text4.Text = rs!EADDRESS
Picture1.Picture = LoadPicture(rs!EPHOTO)


End Sub



Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub lastbtn_Click()
rs.MoveLast
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

Private Sub prevbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub savebtn_Click()
rs.Fields("ENAME").Value = Text1.Text
rs.Fields("EID").Value = Text2.Text
rs.Fields("EMOBILENO").Value = Text3.Text
If Option1.Value = True Then
rs.Fields("GENDER").Value = Option1.Caption
Else
rs.Fields("GENDER").Value = Option2.Caption
End If
rs.Fields("DESIG").Value = Combo1.Text
rs.Fields("EADDRESS").Value = Text4.Text
rs.Fields("EPHOTO").Value = str
MsgBox "Data is saved Successfully", vbInformation
rs.Update

End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2.Text) Then
MsgBox "Should be Numeric!!"
Text2.Text = ""
End If
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3.Text) Then
MsgBox "Should be Numeric!!"
Text3.Text = ""
End If
End Sub

Private Sub updatebtn_Click()
rs.Fields("ENAME").Value = Text1.Text
rs.Fields("EID").Value = Text2.Text
rs.Fields("EMOBILENO").Value = Text3.Text
If Option1.Value = True Then
rs.Fields("GENDER").Value = Option1.Caption
Else
rs.Fields("GENDER").Value = Option2.Caption
End If
rs.Fields("DESIG").Value = Combo1.Text
rs.Fields("EADDRESS").Value = Text4.Text
rs.Fields("EPHOTO").Value = str
MsgBox "Data is Updated Successfully", vbInformation
rs.Update
End Sub

Private Sub uploadbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)

End Sub
