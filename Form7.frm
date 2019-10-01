VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form7 
   Caption         =   "new customer"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form7"
   ScaleHeight     =   8565
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton searchbtn 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   19
      Top             =   1680
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12480
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   18
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton prevbtn 
      Caption         =   "PREV"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   17
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton uploadbtn 
      Caption         =   "UPLOAD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   16
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   13
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   7680
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11760
      Top             =   600
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
      RecordSource    =   "customer"
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
   Begin VB.CommandButton updatebtn 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   7680
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "ADDRESS"
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
      Height          =   1485
      Left            =   4080
      TabIndex        =   10
      Top             =   5640
      Width           =   4695
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
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   7680
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      DataField       =   "PHOTO"
      DataSource      =   "Adodc1"
      Height          =   1935
      Left            =   10680
      ScaleHeight     =   1875
      ScaleWidth      =   2715
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "MOBILENO"
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
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      DataField       =   "ID"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
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
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label5 
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
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label3 
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
      Left            =   960
      TabIndex        =   4
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "New Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "Form7"
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
End If

End Sub
Sub refreshdata()
rs.Close
rs.Open "select * from customer", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "NO record found"
End If

End Sub

Private Sub firstbtn_Click()
rs.MoveFirst
display

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
rs.Open "Select * from customer", con, adOpenDynamic, adLockPessimistic
display


End Sub
Sub display()
Text1.Text = rs!Name
Text2.Text = rs!ID
Text3.Text = rs!MOBILENO
Text4.Text = rs!ADDRESS
Picture1.Picture = LoadPicture(rs!CPHOTO)

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
rs.Fields("NAME").Value = Text1.Text
rs.Fields("ID").Value = Text2.Text
rs.Fields("MOBILENO").Value = Text3.Text
rs.Fields("ADDRESS").Value = Text4.Text
rs.Fields("CPHOTO").Value = str
MsgBox "Data is Saved....!!!!", vbInformation
rs.Update

End Sub

Private Sub searchbtn_Click()
rs.Close
rs.Open "select * from customer where ID='" + Text2.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload

Else
MsgBox "Record Not Found...!!", vbInformation
End If

End Sub
Sub reload()
rs.Close
rs.Open "select * from customer", con, adOpenDynamic, adLockPessimistic
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
rs.Fields("NAME").Value = Text1.Text
rs.Fields("ID").Value = Text2.Text
rs.Fields("MOBILENO").Value = Text3.Text
rs.Fields("ADDRESS").Value = Text4.Text
rs.Fields("CPHOTO").Value = str
MsgBox "Data is Updated Successfully", vbInformation
rs.Update

End Sub

Private Sub uploadbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)

End Sub
