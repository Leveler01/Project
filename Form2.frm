VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "orders"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1455
      Left            =   5880
      TabIndex        =   14
      Top             =   6960
      Width           =   9735
      Begin VB.CommandButton Command8 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   6600
         TabIndex        =   17
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton Command7 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   3480
         TabIndex        =   16
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ADD"
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   1095
      Left            =   240
      TabIndex        =   11
      Top             =   7320
      Width           =   5175
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   5175
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   5175
      Begin VB.CommandButton Command3 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "   SALES     ORDER"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PURCHASE ORDER"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   3495
      Left            =   5760
      TabIndex        =   1
      Top             =   2880
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11280
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
      RecordSource    =   "select * from orders"
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
      Caption         =   "ORDERS"
      Height          =   855
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "Select * from orders where otype='purchase order'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data not Found!!!", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "Select * from orders where otype='sales order'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data not Found!!!", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "Select * from orders where oid='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data not Found!!!", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "Select * from orders where cid='" + Text2.Text + "' or cname='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data not Found!!!", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command5_Click()
Adodc1.RecordSource = "Select * from orders where custid='" + Text3.Text + "' or custname='" + Text3.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data not Found!!!", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command6_Click()
Form10.Show

End Sub

Private Sub Command7_Click()
DataGrid1.AllowUpdate = True

End Sub

Private Sub Command8_Click()
DataGrid1.AllowDelete = True
End Sub
