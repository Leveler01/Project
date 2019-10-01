VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form18 
   Caption         =   "new products"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form18"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   3000
      TabIndex        =   16
      Top             =   8160
      Width           =   10095
      Begin VB.CommandButton Command2 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   6120
         TabIndex        =   18
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         Height          =   615
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   1920
      TabIndex        =   1
      Top             =   2160
      Width           =   12615
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   9120
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   2640
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   9120
         TabIndex        =   15
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   9120
         TabIndex        =   13
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   9120
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "MAXIMUM RETAIL PRICE"
         Height          =   495
         Left            =   6360
         TabIndex        =   14
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "SELLING PRICE"
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "PURCHASE PRICE"
         Height          =   375
         Left            =   6240
         TabIndex        =   10
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "COMPANY ID"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "QUANTITY"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PRODUCT ID"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "PRODUCT NAME"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "COMPANY NAME"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10800
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
      CommandType     =   1
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
      RecordSource    =   "select * from products"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PRODUCTS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection


Private Sub Command1_Click()
rs.addnew
clear

End Sub

Private Sub Command2_Click()
rs.Fields("pname").Value = Text1.Text
rs.Fields("PID").Value = Text2.Text
rs.Fields("qty").Value = Text3.Text
rs.Fields("pprice").Value = Text4.Text
rs.Fields("rprice").Value = Text5.Text
rs.Fields("mrp").Value = Text6.Text
rs.Fields("companyname").Value = Text7.Text
rs.Fields("compid").Value = Text8.Text
MsgBox "Data Saved Successfully!!", vbInformation


End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""


End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
rs.Open "Select * from products", con, adOpenDynamic, adLockPessimistic

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close

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

Private Sub Text4_Change()
If Not IsNumeric(Text4.Text) Then
MsgBox "Should be Numeric!!"
Text4.Text = ""
End If
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5.Text) Then
MsgBox "Should be Numeric!!"
Text5.Text = ""
End If
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6.Text) Then
MsgBox "Should be Numeric!!"
Text6.Text = ""
End If

End Sub

Private Sub Text8_Change()
If Not IsNumeric(Text8.Text) Then
MsgBox "Should be Numeric!!"
Text8.Text = ""
End If
End Sub
