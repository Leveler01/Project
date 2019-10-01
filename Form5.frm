VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   Caption         =   "new ORDER"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   1815
      Left            =   720
      TabIndex        =   22
      Top             =   7320
      Width           =   5655
      Begin VB.TextBox Text8 
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
         Left            =   2520
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text7 
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
         Left            =   2520
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "QUANTITY"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "PRODUCTS"
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
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2535
      Left            =   6960
      TabIndex        =   17
      Top             =   4440
      Width           =   5775
      Begin VB.TextBox Text5 
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
         Left            =   2880
         TabIndex        =   19
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text6 
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
         Left            =   2880
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "CUSTOMER NAME"
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
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "CUSTOMER ID"
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
         TabIndex        =   20
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2535
      Left            =   840
      TabIndex        =   12
      Top             =   4440
      Width           =   5775
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text3 
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
         Left            =   2880
         TabIndex        =   13
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label3 
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
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "COMPANY ID"
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
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   11760
      Top             =   720
      Width           =   1560
      _ExtentX        =   2752
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "orders"
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
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   9720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   9720
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   12615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9360
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   42860545
         CurrentDate     =   42634
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Left            =   9360
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label11 
         Caption         =   "DATE"
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
         Left            =   6960
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "SR NO"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "ORDER ID"
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
         Left            =   6960
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "ORDER TYPE"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ORDER"
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
      Width           =   6255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
rs.addnew
clear
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
DTPicker1.Value = "21/09/2016"
Combo1.Text = "Select Order Type"
End Sub

Private Sub Command2_Click()
rs.Fields("oid") = Text1.Text
rs.Fields("srno") = Text2.Text
rs.Fields("cname") = Text3.Text
rs.Fields("cid") = Text4.Text
rs.Fields("custname") = Text5.Text
rs.Fields("custid") = Text6.Text
rs.Fields("products") = Text7.Text
rs.Fields("quantity") = Text8.Text
rs.Fields("otype") = Combo1.Text
rs.Fields("ddate") = DTPicker1.Value
MsgBox "Data Saved Successfully!!"
rs.Update

End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\Database3.mdb;Persist Security Info=False"
rs.Open "Select * from orders", con, adOpenDynamic, adLockPessimistic


Combo1.AddItem "Purchase Order"
Combo1.AddItem "Sales Order"
End Sub



Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub Text1_Change()
If Not IsNumeric(Text1.Text) Then
MsgBox "Should be Numeric!!", vbCritical
Text1.Text = ""
End If
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2.Text) Then
MsgBox "Should be Numeric!!", vbCritical
Text2.Text = ""
End If

End Sub

Private Sub Text3_Click()
If Combo1.Text = "Sales Order" Then
Text3.Enabled = False
Else
Text3.Enabled = True
End If
End Sub



Private Sub Text4_Click()
If Not IsNumeric(Text4.Text) Then
MsgBox "Should be Numeric!!", vbCritical
Text4.Text = ""
End If
If Combo1.Text = "Sales Order" Then
Text4.Enabled = False
Else
Text4.Enabled = True
End If
End Sub


Private Sub Text5_Click()
If Combo1.Text = "Purchase Order" Then
Text5.Enabled = False
Else
Text5.Enabled = True
End If
End Sub



Private Sub Text6_Click()
If Not IsNumeric(Text6.Text) Then
MsgBox "Should be Numeric!!", vbCritical
Text6.Text = ""
End If
If Combo1.Text = "Purchase Order" Then
Text6.Enabled = False
Else
Text6.Enabled = True
End If
End Sub

Private Sub Text8_Change()
If Not IsNumeric(Text8.Text) Then
MsgBox "Should be Numeric!!", vbCritical
Text8.Text = ""
End If
End Sub
