VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form13 
   Caption         =   "EXISTING EMP"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form13"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   7215
      Begin VB.CommandButton Command6 
         Caption         =   "SHOW GENDER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         Left            =   4920
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "SHOW DESIG"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         Left            =   4920
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   11
         Text            =   " SELECT"
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   10
         Text            =   "SELECT"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "DESIGNATION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   14640
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SHOW ALL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EMPLOYEE ID                                                           EMPLOYEE NAME"
      Height          =   1575
      Left            =   8400
      TabIndex        =   2
      Top             =   2160
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "SEARCH"
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
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10680
      Top             =   360
      Visible         =   0   'False
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
      RecordSource    =   "select * from employee"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form13.frx":0000
      Height          =   3975
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   26
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Show

End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "Select * from employee where EID='" + Text1.Text + "'or ENAME='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data Not Found!!", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "select * from employee"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource

End Sub

Private Sub Command4_Click()
DataGrid1.AllowDelete = True
 
End Sub

Private Sub Command5_Click()
If Combo1.Text = "clerk" Then
Adodc1.RecordSource = "Select * from employee where DESIG='clerk'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
ElseIf Combo1.Text = "accountant" Then
Adodc1.RecordSource = "Select * from employee where DESIG='accountant'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
ElseIf Combo1.Text = "salesman" Then
Adodc1.RecordSource = "Select * from employee where DESIG='salesman'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End If

Combo1.Text = ""


End Sub

Private Sub Command6_Click()
If Combo2.Text = "MALE" Then
Adodc1.RecordSource = "Select * from employee where GENDER='MALE'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
Else
Adodc1.RecordSource = "Select * from employee where GENDER='FEMALE'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End If

Combo2.Text = ""
End Sub

Private Sub Form_Load()
Combo1.AddItem "clerk"
Combo1.AddItem "accountant"
Combo1.AddItem "salesman"
Combo2.AddItem "MALE"
Combo2.AddItem "FEMALE"

End Sub

